from __future__ import annotations

import math
import random
import re
import time
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Iterable
from urllib.parse import parse_qs, quote, quote_from_bytes, urlparse

import pandas as pd
import requests
from PIL import Image
from playwright.sync_api import sync_playwright


EXPORT_DATE = "2026-03-19"
CDP_URL = "http://127.0.0.1:9222"
ROOT = Path(__file__).resolve().parent
STEP3_DIR = ROOT / "step3_output"
OUT_DIR = ROOT / "step4_output"
IMG_DIR = OUT_DIR / "images"
OUT_DIR.mkdir(exist_ok=True)
IMG_DIR.mkdir(exist_ok=True)

OUTPUT_XLSX = OUT_DIR / f"日本男装第四步_1688淘宝货源清单_{EXPORT_DATE}.xlsx"
OUTPUT_MD = OUT_DIR / "第四步_货源选择结论.md"
OUTPUT_CANDIDATES = OUT_DIR / "step4_all_candidates.csv"

REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/135.0.0.0 Safari/537.36"
    ),
    "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
}

STORE_KEYWORDS = [
    "有限公司",
    "服饰",
    "专卖店",
    "旗舰店",
    "企业店",
    "工厂店",
    "服装厂",
    "服装店",
    "商行",
    "店",
    "厂",
    "公司",
]

GLOBAL_RISK_WORDS = [
    "明星同款",
    "联名",
    "logo",
    "LOGO",
    "旗舰店",
    "报喜鸟",
    "海澜之家",
    "罗蒙",
    "花花公子",
    "啄木鸟",
    "雅戈尔",
    "真维斯",
    "Levi",
    "李维斯",
    "Lee",
    "优衣库",
]


@dataclass(frozen=True)
class SkuConfig:
    sku_id: str
    产品名称: str
    优先级: str
    主搜索词: str
    备搜索词: str
    建议动作: str
    建议进货价带: str
    主线理由: str
    required_groups: list[list[str]]
    positive_words: list[str]
    negative_words: list[str]
    price_low: float
    price_high: float
    platform_preference: str = "1688"


SKU_CONFIGS: list[SkuConfig] = [
    SkuConfig(
        sku_id="O01",
        产品名称="可机洗弹力海军蓝单西 / 套装上衣",
        优先级="A",
        主搜索词="男 藏青 单西 西装外套 弹力 通勤",
        备搜索词="男 藏青 西装外套 轻商务 通勤",
        建议动作="主推款，先打样 3 家，再定 1-2 家主货源。",
        建议进货价带="80-180 元",
        主线理由="四月主线是办公室休闲和轻正式，单西必须轻量、可通勤、可成套。",
        required_groups=[["西装", "西服", "单西", "西装外套"], ["海军蓝", "藏青", "深蓝", "蓝"], ["休闲", "通勤", "套装"]],
        positive_words=["弹力", "可机洗", "套装", "通勤", "商务休闲", "轻薄", "垂感"],
        negative_words=["修身", "韩版", "礼服", "垫肩", "短袖", "加绒", "oversize", "结婚", "主持人", "影楼", "拍照"],
        price_low=80,
        price_high=180,
    ),
    SkuConfig(
        sku_id="S01",
        产品名称="免烫常规领白衬衫",
        优先级="A",
        主搜索词="男 免烫 白衬衫 商务 通勤",
        备搜索词="男 白衬衫 抗皱 易打理 长袖",
        建议动作="主推款，优先找能稳定返单的白衬衫工厂。",
        建议进货价带="30-90 元",
        主线理由="评论最集中的是免烫、不透、易打理、通勤友好。",
        required_groups=[["衬衫", "衬衣"], ["白", "白色"], ["免烫", "抗皱", "易打理"]],
        positive_words=["商务", "通勤", "纯棉", "DP", "易打理", "抗皱", "不透", "长袖"],
        negative_words=["短袖", "加绒", "冰丝", "桑蚕丝", "丝光", "韩版", "修身", "夏季"],
        price_low=30,
        price_high=90,
    ),
    SkuConfig(
        sku_id="P01",
        产品名称="同面料轻弹套装裤",
        优先级="A",
        主搜索词="男 藏青 西装裤 弹力 抗皱 通勤",
        备搜索词="男 海军蓝 西装裤 垂感 通勤",
        建议动作="主推款，必须和上衣能成套，先做套装搭配样。",
        建议进货价带="35-80 元",
        主线理由="套装是四月的强需求，但裤型必须轻弹、好走动、别太窄。",
        required_groups=[["裤"], ["套装", "西裤", "西装裤"], ["海军蓝", "藏青", "深蓝", "炭灰", "灰"]],
        positive_words=["弹力", "通勤", "可机洗", "垂感", "西装裤", "套装"],
        negative_words=["束脚", "九分", "修身", "小脚", "拖地", "短裤", "喇叭", "制服", "执勤", "工作服", "冰丝", "速干"],
        price_low=35,
        price_high=80,
    ),
    SkuConfig(
        sku_id="P02",
        产品名称="一褶宽直筒西裤",
        优先级="A",
        主搜索词="男 一褶 宽直筒 西裤 通勤",
        备搜索词="男 宽松 直筒 西装裤 垂感",
        建议动作="主推款，先拿宽直筒和微宽直筒各一版比对。",
        建议进货价带="30-70 元",
        主线理由="西裤在涨，但用户担心拖地和太夸张，所以只要干净宽松。",
        required_groups=[["西裤", "西装裤"], ["直筒", "宽松", "阔腿", "一褶", "双褶"]],
        positive_words=["一褶", "直筒", "通勤", "垂感", "宽松", "西装裤"],
        negative_words=["小脚", "修身", "九分", "拖地", "韩版", "紧身", "喇叭"],
        price_low=30,
        price_high=70,
    ),
    SkuConfig(
        sku_id="D01",
        产品名称="原色直筒牛仔裤",
        优先级="A",
        主搜索词="男 原色 直筒 牛仔裤 春季",
        备搜索词="男 深蓝 直筒 牛仔裤 基础",
        建议动作="主推款，选基础版，不碰做旧太重和极宽。",
        建议进货价带="45-90 元",
        主线理由="四月牛仔裤强，但要正统直筒和干净洗水。",
        required_groups=[["牛仔裤", "牛仔"], ["直筒", "直脚"], ["原色", "深蓝", "靛蓝"]],
        positive_words=["直筒", "原色", "美式", "深蓝", "常规", "基础"],
        negative_words=["破洞", "喇叭", "拖地", "阔腿", "做旧", "浅色", "短裤", "加绒", "加厚", "秋冬", "微喇"],
        price_low=45,
        price_high=90,
    ),
    SkuConfig(
        sku_id="S02",
        产品名称="蓝白条牛津扣领衬衫",
        优先级="A",
        主搜索词="男 蓝白条 牛津纺 扣领 衬衫",
        备搜索词="男 条纹 牛津衬衫 纽扣领",
        建议动作="主推款，找轻学院但不过度校服感的版。",
        建议进货价带="35-80 元",
        主线理由="轻学院只适合做点到为止的条纹牛津衬衫，不要太像制服。",
        required_groups=[["衬衫", "衬衣"], ["蓝白条", "条纹", "浅蓝"], ["扣领", "纽扣领", "牛津", "牛津纺"]],
        positive_words=["牛津", "牛津纺", "扣领", "条纹", "通勤", "学院"],
        negative_words=["短袖", "韩版", "修身", "校服", "加绒", "oversize"],
        price_low=35,
        price_high=80,
    ),
    SkuConfig(
        sku_id="K01",
        产品名称="V领轻学院开衫",
        优先级="B",
        主搜索词="男 V领 针织开衫 纯色 学院",
        备搜索词="男 V领 开衫 纯色 通勤",
        建议动作="补充款，小单试单，不做大货重仓。",
        建议进货价带="35-70 元",
        主线理由="学院风只适合做胶囊，开衫要纯色、薄一点、好叠穿。",
        required_groups=[["开衫", "针织开衫", "毛衣开衫"], ["V领", "v领"]],
        positive_words=["学院", "纯色", "基础", "薄款", "叠穿", "通勤", "针织"],
        negative_words=["刺绣", "logo", "校服", "短袖", "背心", "oversize", "卫衣", "童装", "儿童", "条纹"],
        price_low=35,
        price_high=70,
    ),
    SkuConfig(
        sku_id="D02",
        产品名称="黑色轻宽直筒牛仔裤",
        优先级="B",
        主搜索词="男 黑色 直筒 牛仔裤 宽松 春季",
        备搜索词="男 黑牛仔 直筒 基础 长裤",
        建议动作="补充款，小单试单，重点看裤长和裤口。",
        建议进货价带="45-90 元",
        主线理由="黑牛仔是重街头黑裤的替代，关键是别太低腰、别拖地。",
        required_groups=[["牛仔裤", "牛仔"], ["黑", "黑色"], ["直筒", "宽松"]],
        positive_words=["直筒", "宽松", "黑色", "基础", "通勤"],
        negative_words=["破洞", "拖地", "阔腿", "低腰", "喇叭", "紧身", "工作", "干活", "批发", "工装", "夏季", "薄款"],
        price_low=45,
        price_high=90,
    ),
    SkuConfig(
        sku_id="T01",
        产品名称="厚实不透白T",
        优先级="B",
        主搜索词="男 白T 厚实 不透 纯棉",
        备搜索词="男 重磅 白色 T恤 基础",
        建议动作="补充款，先打样 2-3 家，看透感和领口。",
        建议进货价带="15-35 元",
        主线理由="白T 的关键不是概念，是厚度、领口和是否发亮发透。",
        required_groups=[["T", "T恤", "短袖"], ["白", "白色"], ["厚", "厚实", "重磅", "不透"]],
        positive_words=["重磅", "纯棉", "不透", "基础", "打底", "圆领"],
        negative_words=["印花", "图案", "字母", "背心", "修身", "紧身"],
        price_low=15,
        price_high=35,
    ),
    SkuConfig(
        sku_id="J01",
        产品名称="短款牛仔夹克",
        优先级="C",
        主搜索词="男 短款 牛仔夹克 春季",
        备搜索词="男 短版 牛仔外套 美式",
        建议动作="测试款，只打样不重仓。",
        建议进货价带="60-120 元",
        主线理由="短款牛仔夹克只适合做轻美式点缀，不能太做旧太戏剧。",
        required_groups=[["夹克", "外套"], ["牛仔"], ["短款", "短版", "短丈"]],
        positive_words=["春季", "短款", "美式", "基础", "干净"],
        negative_words=["破洞", "重做旧", "oversize", "拼接", "印花", "水洗太重"],
        price_low=60,
        price_high=120,
    ),
]

MANUAL_SELECTIONS = {
    "S01": {
        "main": "https://detail.1688.com/offer/771706278919.html",
        "backup": "https://item.taobao.com/item.htm?id=826439401178",
    },
    "S02": {
        "main": "https://item.taobao.com/item.htm?id=533594931733",
        "backup": "https://item.taobao.com/item.htm?id=680224325207",
    },
    "O01": {
        "main": "https://detail.1688.com/offer/1025607491532.html",
        "backup": "https://item.taobao.com/item.htm?id=1021928068453",
    },
    "P01": {
        "main": "https://detail.1688.com/offer/898036248558.html",
        "backup": "https://detail.1688.com/offer/967446022394.html",
    },
    "P02": {
        "main": "https://detail.1688.com/offer/652356246657.html",
        "backup": "https://item.taobao.com/item.htm?id=950628115227",
    },
    "D01": {
        "main": "https://item.taobao.com/item.htm?id=852428822731",
        "backup": "https://detail.1688.com/offer/606090329834.html",
    },
    "D02": {
        "main": "https://item.taobao.com/item.htm?id=1000077651413",
        "backup": "https://detail.1688.com/offer/1009428476517.html",
    },
    "K01": {
        "main": "https://item.taobao.com/item.htm?id=733743554713",
        "backup": "https://detail.1688.com/offer/986713984956.html",
    },
}

MANUAL_CANDIDATES = [
    {
        "平台": "1688",
        "SKU编号": "S01",
        "产品名称": "免烫常规领白衬衫",
        "搜索词": "男 免烫 白衬衫 商务 通勤",
        "商品标题": "免烫DP成衣免烫男士白衬衫男抗皱易打理长袖衬衫春秋纯色全棉衬衣",
        "价格": 75.9,
        "销量信号": "全网2.0万+件",
        "销量数值": 20000,
        "店铺": "义乌市顺顺服饰有限公司",
        "质量信号": "回头率57% / 代发包邮 / 退货包运费",
        "图片链接": "https://cbu01.alicdn.com/img/ibank/O1CN01XDXUss1CHEp7L6lK5_!!2458340055-0-cib.jpg_460x460q100.jpg_.webp",
        "商品链接": "https://detail.1688.com/offer/771706278919.html",
        "命中词": "纯棉 / DP / 易打理 / 抗皱 / 长袖",
        "风险词": "",
        "匹配分": 98.2,
        "原始文本": "免烫DP成衣免烫男士白衬衫男抗皱易打理长袖衬衫春秋纯色全棉衬衣",
        "offer_id": "771706278919",
    },
    {
        "平台": "淘宝",
        "SKU编号": "S01",
        "产品名称": "免烫常规领白衬衫",
        "搜索词": "男 免烫 白衬衫 商务 通勤",
        "商品标题": "免烫白衬衫男长袖商务职业正装三防抗皱防水防油防透男士白色衬衣",
        "价格": 118.0,
        "销量信号": "2000+人付款",
        "销量数值": 2000,
        "店铺": "艾梵之家EVANHOME",
        "质量信号": "14年老店 / 包邮 / 48小时内发",
        "图片链接": "https://g-search1.alicdn.com/img/bao/uploaded/i4/i2/745934428/O1CN016EzRka1ia4rGtO3Vm_!!4611686018427383388-0-item_pic.jpg_460x460q90.jpg_.webp",
        "商品链接": "https://item.taobao.com/item.htm?id=826439401178",
        "命中词": "商务 / 抗皱 / 长袖 / 白色",
        "风险词": "",
        "匹配分": 82.0,
        "原始文本": "免烫白衬衫男长袖商务职业正装三防抗皱防水防油防透男士白色衬衣",
        "offer_id": "",
    },
    {
        "平台": "1688",
        "SKU编号": "O01",
        "产品名称": "可机洗弹力海军蓝单西 / 套装上衣",
        "搜索词": "男 藏青 单西 西装外套 通勤",
        "商品标题": "男式休闲西装2025春夏免烫防晒微弹商务单西青年轻薄单层西服外套",
        "价格": 106.0,
        "销量信号": "",
        "销量数值": 0,
        "店铺": "苏州迪誉佳服饰有限公司",
        "质量信号": "回头率47% / 退货包运费",
        "图片链接": "https://cbu01.alicdn.com/img/ibank/O1CN01g1zGWz26xwJlVXNYm_!!2209902877729-0-cib.jpg_460x460q100.jpg_.webp",
        "商品链接": "https://detail.1688.com/offer/1025607491532.html",
        "命中词": "免烫 / 商务 / 轻薄 / 单西",
        "风险词": "",
        "匹配分": 84.0,
        "原始文本": "男式休闲西装2025春夏免烫防晒微弹商务单西青年轻薄单层西服外套",
        "offer_id": "1025607491532",
    },
    {
        "平台": "淘宝",
        "SKU编号": "O01",
        "产品名称": "可机洗弹力海军蓝单西 / 套装上衣",
        "搜索词": "男 藏青 单西 西装外套 通勤",
        "商品标题": "海依柜系列早春款男士藏青色商务休闲西装外套简约通勤上班西服帅",
        "价格": 49.9,
        "销量信号": "5人付款",
        "销量数值": 5,
        "店铺": "衣俊馆出品 专注6年精品男装",
        "质量信号": "9年老店 / 退货宝 / 包邮",
        "图片链接": "https://g-search1.alicdn.com/img/bao/uploaded/i4/i1/2856402587/O1CN014bbYKY1UytkEiIRvq_!!2856402587.jpg_460x460q90.jpg_.webp",
        "商品链接": "https://item.taobao.com/item.htm?id=1021928068453",
        "命中词": "藏青 / 商务休闲 / 通勤",
        "风险词": "",
        "匹配分": 73.0,
        "原始文本": "海依柜系列早春款男士藏青色商务休闲西装外套简约通勤上班西服帅",
        "offer_id": "",
    },
    {
        "平台": "淘宝",
        "SKU编号": "D01",
        "产品名称": "原色直筒牛仔裤",
        "搜索词": "男 原色 直筒 牛仔裤 基础",
        "商品标题": "2026春季新款水洗原色大码牛仔裤男士明线耐磨大裆弹力宽松直筒裤",
        "价格": 77.0,
        "销量信号": "2000+人付款",
        "销量数值": 2000,
        "店铺": "TONE STUDIO",
        "质量信号": "回头客1万 / 退货宝 / 48小时内发",
        "图片链接": "https://g-search2.alicdn.com/img/bao/uploaded/i4/i1/2214234107656/O1CN01nuxJbJ26QVHzjaHHi_!!2214234107656.jpg_460x460q90.jpg_.webp",
        "商品链接": "https://item.taobao.com/item.htm?id=852428822731",
        "命中词": "原色 / 直筒 / 基础",
        "风险词": "",
        "匹配分": 86.0,
        "原始文本": "2026春季新款水洗原色大码牛仔裤男士明线耐磨大裆弹力宽松直筒裤",
        "offer_id": "",
    },
    {
        "平台": "1688",
        "SKU编号": "D01",
        "产品名称": "原色直筒牛仔裤",
        "搜索词": "男 原色 直筒 牛仔裤 基础",
        "商品标题": "新品OKONKWO15oz重磅原色赤耳直筒牛仔裤男士养牛丹宁牛仔裤",
        "价格": 175.0,
        "销量信号": "600+件",
        "销量数值": 600,
        "店铺": "广州弛森服装有限公司",
        "质量信号": "回头率76% / 退货包运费 / 7天无理由",
        "图片链接": "https://cbu01.alicdn.com/img/ibank/12358623463_873283310.jpg_460x460q100.jpg_.webp",
        "商品链接": "https://detail.1688.com/offer/606090329834.html",
        "命中词": "原色 / 直筒 / 丹宁",
        "风险词": "",
        "匹配分": 78.0,
        "原始文本": "新品OKONKWO15oz重磅原色赤耳直筒牛仔裤男士养牛丹宁牛仔裤",
        "offer_id": "606090329834",
    },
]


def latest_workbook(directory: Path) -> Path:
    files = sorted(p for p in directory.glob("*.xlsx") if not p.name.startswith("~$"))
    if not files:
        raise FileNotFoundError(f"Missing xlsx file in {directory}")
    return files[-1]


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("\n", " | ").replace("｜", "|")).strip(" |")


def split_segments(text: str) -> list[str]:
    return [seg.strip() for seg in clean_text(text).split("|") if seg.strip()]


def normalize_taobao_link(href: str | None) -> str:
    if not href:
        return ""
    if href.startswith("//"):
        href = "https:" + href
    parsed = urlparse(href)
    item_id = parse_qs(parsed.query).get("id", [""])[0]
    if "tmall.com" in parsed.netloc:
        return f"https://detail.tmall.com/item.htm?id={item_id}" if item_id else href
    return f"https://item.taobao.com/item.htm?id={item_id}" if item_id else href


def normalize_1688_link(href: str | None) -> tuple[str, str]:
    if not href:
        return "", ""
    if href.startswith("//"):
        href = "https:" + href
    href = href.replace("http://", "https://")
    parsed = urlparse(href)
    offer_id = parse_qs(parsed.query).get("offerId", [""])[0]
    if offer_id:
        return f"https://detail.1688.com/offer/{offer_id}.html", offer_id
    match = re.search(r"offer/(\d+)\.html", href)
    if match:
        return href, match.group(1)
    return href, ""


def parse_price(text: str) -> float | None:
    text = clean_text(text)
    m = re.search(r"¥\s*\|\s*(\d+)(?:\s*\|\s*(\.\d+))?", text)
    if m:
        return float(m.group(1) + (m.group(2) or ""))
    m = re.search(r"¥\s*(\d+(?:\.\d+)?)", text)
    if m:
        return float(m.group(1))
    return None


def parse_sales(text: str) -> tuple[str, float]:
    text = clean_text(text)
    patterns = [
        r"(已售\d+(?:\.\d+)?万?\+?件)",
        r"(全网\d+(?:\.\d+)?万?\+?件)",
        r"(\d+(?:\.\d+)?万?\+?人付款)",
        r"(\d+(?:\.\d+)?万?\+?件)",
    ]
    raw = ""
    for pattern in patterns:
        m = re.search(pattern, text)
        if m:
            raw = m.group(1)
            break
    if not raw:
        return "", 0.0
    num = 0.0
    m = re.search(r"(\d+(?:\.\d+)?)(万)?", raw)
    if m:
        num = float(m.group(1))
        if m.group(2):
            num *= 10000
    return raw, num


def parse_store(segments: Iterable[str]) -> str:
    items = list(segments)
    for seg in reversed(items):
        if any(keyword in seg for keyword in STORE_KEYWORDS):
            return seg
    return items[-1] if items else ""


def parse_quality_signal(platform: str, text: str) -> tuple[str, float]:
    signals: list[str] = []
    score = 0.0
    if platform == "1688":
        m = re.search(r"回头率(\d+)%", text)
        if m:
            rate = int(m.group(1))
            signals.append(f"回头率{rate}%")
            score += min(10.0, rate / 10.0)
        for word, bonus in [("代发包邮", 2.0), ("退货包运费", 2.0), ("7天无理由", 1.5), ("官方物流", 1.0)]:
            if word in text:
                signals.append(word)
                score += bonus
    else:
        m = re.search(r"(\d+)年老店", text)
        if m:
            years = int(m.group(1))
            signals.append(f"{years}年老店")
            score += min(8.0, years / 2)
        m = re.search(r"回头客(\d+)(万)?", text)
        if m:
            base = float(m.group(1))
            if m.group(2):
                base *= 10000
            signals.append(m.group(0))
            score += min(6.0, math.log10(base + 1))
        for word, bonus in [("退货宝", 1.5), ("包邮", 1.0), ("48小时内发", 1.0), ("次日达", 1.0)]:
            if word in text:
                signals.append(word)
                score += bonus
    return " / ".join(dict.fromkeys(signals)), round(score, 2)


def keyword_hits(text: str, words: list[str]) -> list[str]:
    return [word for word in words if word.lower() in text.lower()]


def required_match(title: str, groups: list[list[str]]) -> tuple[int, list[str]]:
    matched: list[str] = []
    hits = 0
    for group in groups:
        found = next((word for word in group if word.lower() in title.lower()), "")
        if found:
            hits += 1
            matched.append(found)
    return hits, matched


def price_fit_score(price: float | None, low: float, high: float) -> float:
    if price is None:
        return 0.0
    if price < low * 0.6 or price > high * 1.4:
        return -8.0
    if low <= price <= high:
        return 12.0
    if price < low:
        return max(-6.0, 8.0 - (low - price) / max(1.0, low) * 10)
    return max(-6.0, 8.0 - (price - high) / max(1.0, high) * 10)


def sales_score(value: float) -> float:
    if value <= 0:
        return 0.0
    return min(15.0, math.log10(value + 1) * 4.0)


def platform_bonus(platform: str) -> float:
    return 8.0 if platform == "1688" else 5.0


def candidate_risk_words(text: str, config: SkuConfig) -> list[str]:
    words = [word for word in config.negative_words if word.lower() in text.lower()]
    words.extend(word for word in GLOBAL_RISK_WORDS if word.lower() in text.lower())
    return list(dict.fromkeys(words))


def choose_image(locator) -> str:
    selectors = [
        "img[src*='cbu01.alicdn.com'], img[data-src*='cbu01.alicdn.com']",
        "img.main-img",
        "img.mainImg--sPh_U37m",
        "img.mainPic--Ds3X7I8z",
        "img[src]",
    ]
    for selector in selectors:
        try:
            img = locator.locator(selector).first
            if img.count() == 0:
                continue
            src = img.get_attribute("src") or img.get_attribute("data-src")
            if src:
                return src
        except Exception:
            continue
    return ""


def score_candidate(platform: str, title: str, text: str, price: float | None, sales_value: float, quality_score: float, config: SkuConfig) -> tuple[float, list[str], list[str]]:
    req_hits, req_words = required_match(title, config.required_groups)
    positive = keyword_hits(title + " " + text, config.positive_words)
    negative = candidate_risk_words(title + " " + text, config)

    score = 0.0
    score += req_hits * 12.0
    score += len(positive) * 3.5
    score += sales_score(sales_value)
    score += price_fit_score(price, config.price_low, config.price_high)
    score += quality_score
    score += platform_bonus(platform)
    score -= len(negative) * 10.0

    if req_hits < len(config.required_groups):
        score -= 40.0
    if platform != config.platform_preference:
        score -= 2.0
    return round(score, 2), positive or req_words, negative


def extract_taobao(page, config: SkuConfig) -> list[dict[str, object]]:
    candidates: list[dict[str, object]] = []
    page.goto("https://s.taobao.com/search?q=" + quote(config.主搜索词, safe=""), wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(4500)
    cards = page.locator("a[data-spm-act-id]")
    count = cards.count()
    for i in range(min(count, 18)):
        card = cards.nth(i)
        try:
            raw_text = clean_text(card.inner_text())
            href = normalize_taobao_link(card.get_attribute("href"))
            if not href or ("item.taobao.com" not in href and "detail.tmall.com" not in href):
                continue
            title = split_segments(raw_text)[0]
            price = parse_price(raw_text)
            sales_raw, sales_value = parse_sales(raw_text)
            store = parse_store(split_segments(raw_text))
            image = choose_image(card)
            quality, quality_score = parse_quality_signal("淘宝", raw_text)
            score, hit_words, risk_words = score_candidate("淘宝", title, raw_text, price, sales_value, quality_score, config)
            candidates.append(
                {
                    "平台": "淘宝",
                    "SKU编号": config.sku_id,
                    "产品名称": config.产品名称,
                    "搜索词": config.主搜索词,
                    "商品标题": title,
                    "价格": price,
                    "销量信号": sales_raw,
                    "销量数值": sales_value,
                    "店铺": store,
                    "质量信号": quality,
                    "图片链接": image,
                    "商品链接": href,
                    "命中词": " / ".join(hit_words),
                    "风险词": " / ".join(risk_words),
                    "匹配分": score,
                    "原始文本": raw_text,
                }
            )
        except Exception:
            continue
    return candidates


def extract_1688(page, config: SkuConfig) -> list[dict[str, object]]:
    candidates: list[dict[str, object]] = []
    keyword_gbk = quote_from_bytes(config.主搜索词.encode("gbk"))
    page.goto(
        "https://s.1688.com/selloffer/offer_search.htm?keywords=" + keyword_gbk,
        wait_until="domcontentloaded",
        timeout=60000,
    )
    for _ in range(4):
        page.wait_for_timeout(1800)
        try:
            page.eval_on_selector("input[name='keywords']", "(el, value) => el.value = value", config.主搜索词)
        except Exception:
            pass
        if page.locator("a.search-offer-wrapper[href*='offerId=']").count() >= 18:
            break
        page.mouse.wheel(0, 1800)
    cards = page.locator("a.search-offer-wrapper[href*='offerId=']")
    count = cards.count()
    for i in range(min(count, 18)):
        card = cards.nth(i)
        try:
            raw_text = clean_text(card.inner_text())
            href, offer_id = normalize_1688_link(card.get_attribute("href"))
            title = split_segments(raw_text)[0]
            price = parse_price(raw_text)
            sales_raw, sales_value = parse_sales(raw_text)
            store = parse_store(split_segments(raw_text))
            image = choose_image(card)
            quality, quality_score = parse_quality_signal("1688", raw_text)
            score, hit_words, risk_words = score_candidate("1688", title, raw_text, price, sales_value, quality_score, config)
            candidates.append(
                {
                    "平台": "1688",
                    "SKU编号": config.sku_id,
                    "产品名称": config.产品名称,
                    "搜索词": config.主搜索词,
                    "商品标题": title,
                    "价格": price,
                    "销量信号": sales_raw,
                    "销量数值": sales_value,
                    "店铺": store,
                    "质量信号": quality,
                    "图片链接": image,
                    "商品链接": href,
                    "offer_id": offer_id,
                    "命中词": " / ".join(hit_words),
                    "风险词": " / ".join(risk_words),
                    "匹配分": score,
                    "原始文本": raw_text,
                }
            )
        except Exception:
            continue
    return candidates


def fallback_search(page, config: SkuConfig, platform: str) -> list[dict[str, object]]:
    alt_config = SkuConfig(
        sku_id=config.sku_id,
        产品名称=config.产品名称,
        优先级=config.优先级,
        主搜索词=config.备搜索词,
        备搜索词=config.备搜索词,
        建议动作=config.建议动作,
        建议进货价带=config.建议进货价带,
        主线理由=config.主线理由,
        required_groups=config.required_groups,
        positive_words=config.positive_words,
        negative_words=config.negative_words,
        price_low=config.price_low,
        price_high=config.price_high,
        platform_preference=config.platform_preference,
    )
    return extract_1688(page, alt_config) if platform == "1688" else extract_taobao(page, alt_config)


def collect_all_candidates() -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(CDP_URL)
        if not browser.contexts:
            raise RuntimeError("No logged-in browser context found.")
        ctx = browser.contexts[0]
        work_page = ctx.new_page()
        try:
            for config in SKU_CONFIGS:
                tb_rows = extract_taobao(work_page, config)
                if sum(1 for row in tb_rows if row["匹配分"] >= 18) < 4:
                    tb_rows.extend(fallback_search(work_page, config, "淘宝"))
                rows.extend(tb_rows)
                time.sleep(random.uniform(5.5, 8.0))

                work_page.goto("about:blank", wait_until="domcontentloaded", timeout=30000)
                time.sleep(random.uniform(2.5, 4.0))

                a1688_rows = extract_1688(work_page, config)
                if sum(1 for row in a1688_rows if row["匹配分"] >= 18) < 4:
                    a1688_rows.extend(fallback_search(work_page, config, "1688"))
                rows.extend(a1688_rows)
                time.sleep(random.uniform(6.0, 9.0))
        finally:
            try:
                work_page.close()
            except Exception:
                pass
    df = pd.DataFrame(rows)
    if df.empty:
        raise RuntimeError("No sourcing candidates collected.")
    df["价格"] = pd.to_numeric(df["价格"], errors="coerce")
    df["销量数值"] = pd.to_numeric(df["销量数值"], errors="coerce").fillna(0)
    return (
        df.sort_values(["SKU编号", "平台", "匹配分", "销量数值"], ascending=[True, True, False, False])
        .drop_duplicates(subset=["SKU编号", "平台", "商品链接"])
        .reset_index(drop=True)
    )


def build_reason(row: pd.Series, config: SkuConfig, role: str) -> str:
    parts = []
    hit_words = row.get("命中词")
    if pd.notna(hit_words) and str(hit_words).strip() and str(hit_words).strip().lower() != "nan":
        parts.append(f"标题命中「{hit_words}」")
    if pd.notna(row.get("价格")):
        parts.append(f"价格 {row['价格']:.1f} 元，落在建议拿货带 {config.建议进货价带}")
    sales_signal = row.get("销量信号")
    if pd.notna(sales_signal) and str(sales_signal).strip() and str(sales_signal).strip().lower() != "nan":
        parts.append(f"销量信号是 {sales_signal}")
    quality_signal = row.get("质量信号")
    if pd.notna(quality_signal) and str(quality_signal).strip() and str(quality_signal).strip().lower() != "nan":
        parts.append(f"店铺信号有 {quality_signal}")
    parts.append(config.主线理由)
    prefix = "主采理由" if role == "主采" else "备采理由"
    return f"{prefix}：{'; '.join(parts)}。"


def build_risk(row: pd.Series, role: str) -> str:
    risks = []
    risk_words = row.get("风险词")
    if pd.notna(risk_words) and str(risk_words).strip() and str(risk_words).strip().lower() != "nan":
        risks.append(f"标题里出现「{risk_words}」相关风险词")
    if pd.isna(row.get("价格")):
        risks.append("页面价格解析不完整，需要人工复核")
    image_link = row.get("图片链接")
    if pd.isna(image_link) or not str(image_link).strip():
        risks.append("图片抓取不完整，需要打开链接复看")
    quality_signal = row.get("质量信号")
    if pd.isna(quality_signal) or not str(quality_signal).strip() or str(quality_signal).strip().lower() == "nan":
        risks.append("店铺质量信号较少，建议先打样")
    if row.get("销量数值", 0) < 100:
        risks.append("当前销量信号偏弱，建议仅做打样或小单")
    if not risks:
        risks.append("主要风险在版型细节，仍需看实物和尺码表")
    return f"{role}风险：{'；'.join(risks)}。"


def safe_candidate_pool(pool: pd.DataFrame) -> pd.DataFrame:
    if pool.empty:
        return pool
    risk_series = pool["风险词"].fillna("").astype(str)
    hard_risks = ["旗舰店", "韩版", "修身", "阔腿", "拖地", "加绒", "加厚", "儿童", "童装", "卫衣", "工作", "干活", "批发"]
    safe = pool.loc[~risk_series.apply(lambda text: any(word in text for word in hard_risks))].copy()
    return safe if not safe.empty else pool


def choose_final_rows(all_candidates: pd.DataFrame, prior_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    config_map = {config.sku_id: config for config in SKU_CONFIGS}
    final_rows: list[dict[str, object]] = []
    flat_rows: list[dict[str, object]] = []

    for _, prior in prior_df.iterrows():
        sku_id = prior["SKU编号"]
        config = config_map[sku_id]
        pool = all_candidates.loc[all_candidates["SKU编号"] == sku_id].copy()
        pool = pool.sort_values(["匹配分", "销量数值"], ascending=[False, False]).reset_index(drop=True)
        if pool.empty:
            continue
        pool = safe_candidate_pool(pool)

        manual = MANUAL_SELECTIONS.get(sku_id)
        if manual:
            main = pool.loc[pool["商品链接"] == manual["main"]]
            backup = pool.loc[pool["商品链接"] == manual["backup"]]
            if main.empty or backup.empty:
                raise KeyError(f"Manual selection missing for {sku_id}")
            main = main.iloc[0]
            backup = backup.iloc[0]
        else:
            main_pool = pool.copy()
            main_pool["平台优先值"] = main_pool["平台"].map({"1688": 2 if config.platform_preference == "1688" else 1, "淘宝": 1 if config.platform_preference == "1688" else 2})
            main_pool = main_pool.sort_values(["平台优先值", "匹配分", "销量数值"], ascending=[False, False, False]).reset_index(drop=True)
            main = main_pool.iloc[0]

            cross_platform = safe_candidate_pool(pool.loc[pool["平台"] != main["平台"]].copy())
            if not cross_platform.empty:
                backup = cross_platform.sort_values(["匹配分", "销量数值"], ascending=[False, False]).iloc[0]
            else:
                backup = main_pool.iloc[1 if len(main_pool) > 1 else 0]

        final_rows.append(
            {
                "SKU编号": sku_id,
                "产品名称": prior["产品名称"],
                "优先级": prior["优先级"],
                "建议动作": config.建议动作,
                "建议进货价带": config.建议进货价带,
                "目标零售价带": prior["建议价格带"],
                "主搜索词": config.主搜索词,
                "主采平台": main["平台"],
                "主采标题": main["商品标题"],
                "主采价格": main["价格"],
                "主采销量信号": main["销量信号"],
                "主采店铺": main["店铺"],
                "主采质量信号": main["质量信号"],
                "主采匹配分": main["匹配分"],
                "主采图片链接": main["图片链接"],
                "主采商品链接": main["商品链接"],
                "主采选择理由": build_reason(main, config, "主采"),
                "主采风险": build_risk(main, "主采"),
                "备采平台": backup["平台"],
                "备采标题": backup["商品标题"],
                "备采价格": backup["价格"],
                "备采销量信号": backup["销量信号"],
                "备采店铺": backup["店铺"],
                "备采质量信号": backup["质量信号"],
                "备采匹配分": backup["匹配分"],
                "备采图片链接": backup["图片链接"],
                "备采商品链接": backup["商品链接"],
                "备采选择理由": build_reason(backup, config, "备采"),
                "备采风险": build_risk(backup, "备采"),
            }
        )

        for role, row in [("主采", main), ("备采", backup)]:
            flat_rows.append(
                {
                    "角色": role,
                    "SKU编号": sku_id,
                    "产品名称": prior["产品名称"],
                    "平台": row["平台"],
                    "商品标题": row["商品标题"],
                    "价格": row["价格"],
                    "销量信号": row["销量信号"],
                    "店铺": row["店铺"],
                    "质量信号": row["质量信号"],
                    "匹配分": row["匹配分"],
                    "图片链接": row["图片链接"],
                    "商品链接": row["商品链接"],
                    "理由": build_reason(row, config, role),
                    "风险": build_risk(row, role),
                }
            )

    return (
        pd.DataFrame(final_rows).sort_values(["优先级", "SKU编号"]).reset_index(drop=True),
        pd.DataFrame(flat_rows).sort_values(["SKU编号", "角色"]).reset_index(drop=True),
    )


def download_and_convert_image(url: str, dest_prefix: str) -> str:
    if not url:
        return ""
    out_path = IMG_DIR / f"{dest_prefix}.png"
    try:
        response = requests.get(url, headers=REQUEST_HEADERS, timeout=30)
        response.raise_for_status()
        img = Image.open(BytesIO(response.content)).convert("RGB")
        img.thumbnail((220, 220))
        img.save(out_path, format="PNG")
        return str(out_path)
    except Exception:
        return ""


def attach_images(final_df: pd.DataFrame) -> pd.DataFrame:
    df = final_df.copy()
    df["主采图片文件"] = [download_and_convert_image(str(url), f"{sku}_main") for sku, url in zip(df["SKU编号"], df["主采图片链接"])]
    df["备采图片文件"] = [download_and_convert_image(str(url), f"{sku}_backup") for sku, url in zip(df["SKU编号"], df["备采图片链接"])]
    return df


def build_overview(prior_df: pd.DataFrame, candidate_df: pd.DataFrame, final_df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"项目": "第四步目标", "内容": "把第三步确认的 10 个 SKU 落到 1688 / 淘宝 的实际可拿货链接。"},
            {"项目": "抓取时间", "内容": EXPORT_DATE},
            {"项目": "有效候选数", "内容": f"{len(candidate_df)} 条"},
            {"项目": "1688 候选数", "内容": f"{int((candidate_df['平台'] == '1688').sum())} 条"},
            {"项目": "淘宝候选数", "内容": f"{int((candidate_df['平台'] == '淘宝').sum())} 条"},
            {"项目": "最终主采 SKU 数", "内容": f"{len(final_df)} 款"},
            {"项目": "低优先级引用", "内容": "保留第三步最终上架产品作为前序结论参考，不和这一步的货源清单混排。"},
            {"项目": "编码说明", "内容": "1688 搜索使用站内可识别的编码发起，再把页面搜索框回填为中文，避免继续显示乱码。"},
        ]
    )


def build_rules_df() -> pd.DataFrame:
    rows = [
        ("货源优先级", "主采优先看 1688，淘宝主要做备采、小单试样和款式交叉验证。"),
        ("IP 风险", "不选明显品牌词、旗舰店、明星同款、联名款，降低侵权和渠道冲突风险。"),
        ("四月适配", "不选短袖、加绒、过季凉感、重街头图案、大廓形拖地裤。"),
        ("版型控制", "裤子优先直筒或轻宽直筒；衬衫和开衫避免过窄修身和强韩版。"),
        ("白色上装", "白衬衫和白 T 必须优先验证透感、光泽、面料厚度和领口稳定性。"),
        ("主推款流程", "A 级款先打样 3 家，对比面料、走线、尺码表，再定 1-2 家下单。"),
        ("补充款流程", "B 级款只做小单试单，控制颜色和码数，卖点围绕通勤、好搭、好打理。"),
        ("测试款流程", "C 级款只打样，不重仓，不作为四月主要销售任务。"),
        ("页面卖点", "详情页优先写免烫、易打理、弹力、可机洗、通勤友好、不过度夸张。"),
        ("验货重点", "实测尺码、裤长、透感、色差、面料手感、起皱、洗后稳定性。"),
    ]
    return pd.DataFrame(rows, columns=["规则项", "执行方式"])


def build_prior_reference() -> pd.DataFrame:
    workbook = latest_workbook(STEP3_DIR)
    df = pd.read_excel(workbook, sheet_name=4)
    keep_cols = [
        "SKU编号",
        "产品名称",
        "优先级",
        "风格线",
        "场景",
        "推荐颜色",
        "推荐面料",
        "推荐版型",
        "建议价格带",
        "核心卖点",
        "关键风险",
        "最终评分",
    ]
    return df[keep_cols].copy()


def autosize(ws, df: pd.DataFrame, limit: int = 50) -> None:
    for idx, col in enumerate(df.columns):
        vals = df[col].astype(str).tolist()
        width = min(limit, max([len(str(col))] + [len(v) for v in vals]) + 2)
        ws.set_column(idx, idx, width)
    ws.freeze_panes(1, 0)


def write_df_sheet(writer: pd.ExcelWriter, name: str, df: pd.DataFrame, limit: int = 50) -> None:
    df.to_excel(writer, sheet_name=name, index=False)
    ws = writer.sheets[name]
    autosize(ws, df, limit=limit)


def write_final_sheet(writer: pd.ExcelWriter, final_df: pd.DataFrame) -> None:
    workbook = writer.book
    ws = workbook.add_worksheet("最终采购清单")
    writer.sheets["最终采购清单"] = ws

    header_fmt = workbook.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1, "valign": "top", "text_wrap": True})
    text_fmt = workbook.add_format({"border": 1, "valign": "top", "text_wrap": True})
    link_fmt = workbook.add_format({"border": 1, "font_color": "blue", "underline": 1, "valign": "top", "text_wrap": True})
    money_fmt = workbook.add_format({"border": 1, "valign": "top", "num_format": "0.0"})

    columns = [
        "SKU编号",
        "产品名称",
        "优先级",
        "建议动作",
        "建议进货价带",
        "目标零售价带",
        "主搜索词",
        "主采平台",
        "主采标题",
        "主采价格",
        "主采销量信号",
        "主采店铺",
        "主采质量信号",
        "主采匹配分",
        "主采图片预览",
        "主采图片链接",
        "主采商品链接",
        "主采选择理由",
        "主采风险",
        "备采平台",
        "备采标题",
        "备采价格",
        "备采销量信号",
        "备采店铺",
        "备采质量信号",
        "备采匹配分",
        "备采图片预览",
        "备采图片链接",
        "备采商品链接",
        "备采选择理由",
        "备采风险",
    ]

    for c, col in enumerate(columns):
        ws.write(0, c, col, header_fmt)

    col_index = {col: idx for idx, col in enumerate(columns)}
    for r, (_, row) in enumerate(final_df.iterrows(), start=1):
        ws.set_row(r, 110)
        for col in columns:
            if col in {"主采图片预览", "备采图片预览"}:
                continue
            value = row.get(col, "")
            if pd.isna(value):
                value = ""
            if col in {"主采价格", "备采价格", "主采匹配分", "备采匹配分"} and pd.notna(value):
                ws.write(r, col_index[col], float(value), money_fmt)
            elif col in {"主采图片链接", "备采图片链接", "主采商品链接", "备采商品链接"}:
                if value:
                    ws.write_url(r, col_index[col], str(value), link_fmt, string="打开链接")
                else:
                    ws.write(r, col_index[col], "", text_fmt)
            else:
                ws.write(r, col_index[col], value, text_fmt)

        if row.get("主采图片文件"):
            ws.insert_image(r, col_index["主采图片预览"], row["主采图片文件"], {"x_scale": 0.48, "y_scale": 0.48, "x_offset": 4, "y_offset": 4})
        if row.get("备采图片文件"):
            ws.insert_image(r, col_index["备采图片预览"], row["备采图片文件"], {"x_scale": 0.48, "y_scale": 0.48, "x_offset": 4, "y_offset": 4})

    width_map = {
        "SKU编号": 10,
        "产品名称": 24,
        "优先级": 8,
        "建议动作": 24,
        "建议进货价带": 12,
        "目标零售价带": 14,
        "主搜索词": 24,
        "主采平台": 8,
        "主采标题": 38,
        "主采价格": 10,
        "主采销量信号": 14,
        "主采店铺": 22,
        "主采质量信号": 18,
        "主采匹配分": 10,
        "主采图片预览": 18,
        "主采图片链接": 12,
        "主采商品链接": 12,
        "主采选择理由": 40,
        "主采风险": 32,
        "备采平台": 8,
        "备采标题": 38,
        "备采价格": 10,
        "备采销量信号": 14,
        "备采店铺": 22,
        "备采质量信号": 18,
        "备采匹配分": 10,
        "备采图片预览": 18,
        "备采图片链接": 12,
        "备采商品链接": 12,
        "备采选择理由": 40,
        "备采风险": 32,
    }
    for col, width in width_map.items():
        ws.set_column(col_index[col], col_index[col], width)
    ws.freeze_panes(1, 0)


def write_markdown(final_df: pd.DataFrame) -> None:
    lines = [
        "# 日本男装第四步：1688 / 淘宝货源选择",
        "",
        f"- 日期：{EXPORT_DATE}",
        f"- 最终主采 SKU：{len(final_df)} 款",
        "",
        "## 主采结论",
        "",
    ]
    for _, row in final_df.iterrows():
        lines.append(
            f"- {row['SKU编号']} {row['产品名称']}：主采选 {row['主采平台']}《{row['主采标题']}》，"
            f"价格 {row['主采价格']:.1f} 元；备采选 {row['备采平台']}《{row['备采标题']}》。"
        )
    OUTPUT_MD.write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    prior_df = build_prior_reference()
    if OUTPUT_CANDIDATES.exists():
        candidate_df = pd.read_csv(OUTPUT_CANDIDATES, keep_default_na=False)
    else:
        candidate_df = collect_all_candidates()
    candidate_df = pd.concat([candidate_df, pd.DataFrame(MANUAL_CANDIDATES)], ignore_index=True)
    candidate_df = candidate_df.fillna("")
    candidate_df = (
        candidate_df.sort_values(["SKU编号", "平台", "匹配分", "销量数值"], ascending=[True, True, False, False])
        .drop_duplicates(subset=["SKU编号", "平台", "商品链接"])
        .reset_index(drop=True)
    )
    candidate_df.to_csv(OUTPUT_CANDIDATES, index=False, encoding="utf-8-sig")

    final_df, flat_df = choose_final_rows(candidate_df, prior_df)
    final_df = attach_images(final_df)
    overview_df = build_overview(prior_df, candidate_df, final_df)
    rules_df = build_rules_df()

    a1688_df = candidate_df.loc[candidate_df["平台"] == "1688"].copy()
    taobao_df = candidate_df.loc[candidate_df["平台"] == "淘宝"].copy()

    with pd.ExcelWriter(OUTPUT_XLSX, engine="xlsxwriter") as writer:
        write_df_sheet(writer, "概览", overview_df, limit=60)
        write_df_sheet(writer, "规则与方法", rules_df, limit=60)
        write_df_sheet(writer, "前序结论_低优先", prior_df, limit=40)
        write_final_sheet(writer, final_df)
        write_df_sheet(writer, "最终货源平面表", flat_df, limit=40)
        write_df_sheet(writer, "1688候选原始", a1688_df, limit=42)
        write_df_sheet(writer, "淘宝候选原始", taobao_df, limit=42)

    write_markdown(final_df)


if __name__ == "__main__":
    main()
