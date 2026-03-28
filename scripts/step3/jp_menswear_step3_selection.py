from __future__ import annotations

import math
import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd


EXPORT_DATE = "2026-03-19"
APRIL_DIR = Path("april_sell_output")
STEP2_DIR = Path("step2_plus_output")
OUT_DIR = Path("step3_output")
OUT_DIR.mkdir(exist_ok=True)

OUTPUT_XLSX = OUT_DIR / f"日本男装第三步_具体选品清单_{EXPORT_DATE}.xlsx"
OUTPUT_MD = OUT_DIR / "第三步_具体选品结论.md"


JUDGMENT_POINTS = {
    "四月主推": 95,
    "四月主线风格": 92,
    "四月稳卖": 82,
    "四月加分风格": 76,
    "四月季节机会": 68,
    "成长风格": 62,
    "观察项": 48,
    "大盘仍大但在回落": 44,
    "四月谨慎": 30,
    "低搜索量，不单独决策": 15,
}


POSITIVE_WORDS = [
    "良い",
    "いい",
    "満足",
    "便利",
    "着やす",
    "快適",
    "おすすめ",
    "オススメ",
    "リピ",
    "ジャスト",
    "ぴったり",
    "ピッタリ",
    "助かる",
    "問題なし",
    "好き",
]

NEGATIVE_WORDS = [
    "透け",
    "難し",
    "ミスマッチ",
    "ダサ",
    "大きすぎ",
    "小さめ",
    "苦し",
    "薄い",
    "シワ",
    "不安",
    "避け",
    "交換",
    "長すぎ",
    "ダボ",
    "浮く",
    "傷む",
]


SIGNAL_TAG_RULES = {
    "免烫/易护理": ["ノーアイロン", "形態安定", "お手入れ", "免烫", "易护理"],
    "快干/透气": ["速乾", "快干", "通気", "吸水速乾", "サラサラ"],
    "弹力/舒适": ["ストレッチ", "伸び", "伸縮", "動きやす", "快適", "舒服", "舒适"],
    "不透/低光泽": ["透け", "光沢", "白T", "白シャツ", "乳首", "不透"],
    "尺寸/版型": ["サイズ", "サイズ感", "小さめ", "大きめ", "ジャスト", "スリム"],
    "裤长/拖地": ["丈", "股下", "裾", "裾直し", "拖地"],
    "宽度风险": ["ワイド", "ダボ", "太い", "ルーズ", "太宽"],
    "办公室休闲": ["オフィスカジュアル", "ビジカジ", "仕事", "通勤", "入社式", "就活"],
    "学院风": ["プレッピー", "アイビー", "スクール", "紺ブレ", "Vネック", "学院"],
    "牛仔": ["デニム", "ジーンズ", "505", "503"],
    "价格接受": ["価格", "コスパ", "お手頃", "安い", "值", "合理"],
    "可机洗": ["洗える", "ウォッシャブル", "可机洗"],
    "防泼水": ["撥水", "防泼水"],
}


@dataclass(frozen=True)
class Candidate:
    sku_id: str
    产品名称: str
    上架结论: str
    优先级: str
    品类: str
    风格线: str
    场景: str
    对应第一步商品: str | None
    对应第一步风格: str | None
    对应第一步场景: str | None
    对应第一步功能: str | None
    amazon_group: str | None
    amazon_title_patterns: str
    signal_patterns: str
    推荐颜色: str
    推荐面料: str
    推荐版型: str
    建议价格带: str
    核心卖点: str
    关键风险: str
    日文搜索词: str


CANDIDATES: list[Candidate] = [
    Candidate(
        sku_id="S01",
        产品名称="免烫常规领白衬衫",
        上架结论="上架",
        优先级="A",
        品类="衬衫",
        风格线="办公室休闲",
        场景="通勤 / 入职 / 面试",
        对应第一步商品="男士衬衫",
        对应第一步风格="办公室休闲男装",
        对应第一步场景="办公室休闲男装",
        对应第一步功能="弹力男装",
        amazon_group="衬衫",
        amazon_title_patterns=r"ノーアイロン|形態安定|白ワイシャツ|レギュラー",
        signal_patterns=r"ワイシャツ|白シャツ|白T|ノーアイロン|透け|光沢",
        推荐颜色="白",
        推荐面料="低光泽棉涤混纺，微弹，非透",
        推荐版型="常规偏宽松，不做过窄修身",
        建议价格带="JPY 2,990-3,990",
        核心卖点="免烫、快干、低光泽、不易透、通勤友好",
        关键风险="面料太薄、光泽过强、领型太商务化",
        日文搜索词="ノーアイロン シャツ メンズ / 白シャツ メンズ / オフィスカジュアル シャツ メンズ",
    ),
    Candidate(
        sku_id="S02",
        产品名称="蓝白条牛津扣领衬衫",
        上架结论="上架",
        优先级="A",
        品类="衬衫",
        风格线="轻美式 / 轻学院",
        场景="通勤 / 周末 / 学院胶囊",
        对应第一步商品="男士衬衫",
        对应第一步风格="美式休闲男装",
        对应第一步场景="办公室休闲男装",
        对应第一步功能=None,
        amazon_group="衬衫",
        amazon_title_patterns=r"オックスフォード|ボタンダウン|カジュアル",
        signal_patterns=r"ブルーのシャツ|オックスフォード|ボタンダウン|ブルックス|J.PRESS|プレッピー",
        推荐颜色="蓝白条 / 浅蓝",
        推荐面料="中等厚度牛津纺，挺但不硬",
        推荐版型="常规宽松，衣长可内外穿",
        建议价格带="JPY 3,290-4,290",
        核心卖点="好搭配、轻学院、通勤和周末都能穿",
        关键风险="太像校服、条纹过粗、领子过软",
        日文搜索词="オックスフォード シャツ メンズ / ボタンダウン シャツ メンズ / ストライプ シャツ メンズ",
    ),
    Candidate(
        sku_id="O01",
        产品名称="可机洗弹力海军蓝单西 / 套装上衣",
        上架结论="上架",
        优先级="A",
        品类="西装外套",
        风格线="办公室休闲",
        场景="通勤 / 入职 / 轻正式",
        对应第一步商品="男士套装",
        对应第一步风格="办公室休闲男装",
        对应第一步场景="办公室休闲男装",
        对应第一步功能="弹力男装",
        amazon_group="西装外套",
        amazon_title_patterns=r"紺ブレ|ブレザー|オフィス|ビジネス|ジャケパン",
        signal_patterns=r"ブレザー|紺ブレ|ジャケット|感動ブレザー|ジャケパン|オフィスカジュアル",
        推荐颜色="海军蓝 / 炭灰",
        推荐面料="聚酯+粘胶+氨纶，轻量可机洗",
        推荐版型="两粒扣，肩部自然，非厚垫肩",
        建议价格带="JPY 7,990-10,990",
        核心卖点="可机洗、轻弹、通勤、可成套",
        关键风险="太正式、太薄、肩线过硬",
        日文搜索词="紺ブレ メンズ / 感動ジャケット / セットアップ メンズ",
    ),
    Candidate(
        sku_id="P01",
        产品名称="同面料轻弹套装裤",
        上架结论="上架",
        优先级="A",
        品类="长裤",
        风格线="办公室休闲",
        场景="通勤 / 轻正式 / 成套",
        对应第一步商品="男士套装",
        对应第一步风格="办公室休闲男装",
        对应第一步场景="通勤男装",
        对应第一步功能="弹力男装",
        amazon_group="长裤",
        amazon_title_patterns=r"ストレッチ|ビジカジ|パンツ|オフィス|テーパード",
        signal_patterns=r"感動パンツ|セットアップ|パンツ|通勤|動きやす|ウォッシャブル|ビジカジ",
        推荐颜色="海军蓝 / 炭灰",
        推荐面料="与 O01 同面料，微弹可机洗",
        推荐版型="微锥形或直筒，不做贴腿裤",
        建议价格带="JPY 3,990-5,990",
        核心卖点="轻弹、可机洗、成套、通勤友好",
        关键风险="裤型太窄、裆太长、面料发亮",
        日文搜索词="感動パンツ / セットアップ パンツ メンズ / ビジカジ パンツ メンズ",
    ),
    Candidate(
        sku_id="P02",
        产品名称="一褶宽直筒西裤",
        上架结论="上架",
        优先级="A",
        品类="长裤",
        风格线="办公室休闲 / 简约",
        场景="通勤 / 日常",
        对应第一步商品="男士西裤",
        对应第一步风格="办公室休闲男装",
        对应第一步场景="办公室休闲男装",
        对应第一步功能="弹力男装",
        amazon_group="长裤",
        amazon_title_patterns=r"スラックス|ビジカジ|ストレート|セミワイド|ストレッチ",
        signal_patterns=r"スラックス|ワイドパンツ|低身長|丈|ダボ|オフィスカジュアル",
        推荐颜色="炭灰 / 中灰",
        推荐面料="哑光混纺，轻弹",
        推荐版型="一褶宽直筒，控制裤口和裤长",
        建议价格带="JPY 3,990-4,990",
        核心卖点="不难搭、通勤可穿、显利落、走动舒服",
        关键风险="过宽、拖地、前裆过长",
        日文搜索词="スラックス メンズ / ワンタック スラックス メンズ / オフィスカジュアル パンツ メンズ",
    ),
    Candidate(
        sku_id="D01",
        产品名称="原色直筒牛仔裤",
        上架结论="上架",
        优先级="A",
        品类="牛仔裤",
        风格线="轻美式",
        场景="日常 / 周末 / 轻通勤",
        对应第一步商品="男士牛仔裤",
        对应第一步风格="美式休闲男装",
        对应第一步场景=None,
        对应第一步功能=None,
        amazon_group="牛仔裤",
        amazon_title_patterns=r"ストレート|505|レギュラー",
        signal_patterns=r"デニム|ジーンズ|505|503|ストレート|履きやす",
        推荐颜色="原色靛蓝",
        推荐面料="11.5-12.5oz 轻弹或软化牛仔",
        推荐版型="正统直筒，不做极宽",
        建议价格带="JPY 4,990-6,990",
        核心卖点="好搭、不过时、长度友好、舒服",
        关键风险="太硬、太窄、洗水太复古",
        日文搜索词="デニム メンズ ストレート / リーバイス 505 メンズ / レギュラーストレート デニム メンズ",
    ),
    Candidate(
        sku_id="D02",
        产品名称="黑色轻宽直筒牛仔裤",
        上架结论="上架",
        优先级="B",
        品类="牛仔裤",
        风格线="简约 / 轻时装感",
        场景="日常 / 晚间 / 轻通勤",
        对应第一步商品="男士牛仔裤",
        对应第一步风格="简约男装",
        对应第一步场景=None,
        对应第一步功能=None,
        amazon_group="牛仔裤",
        amazon_title_patterns=r"ブラック|ワイド|ルーズ|ストレート",
        signal_patterns=r"ブラックデニム|黒デニム|ワイド|ルーズ|ストレート",
        推荐颜色="黑",
        推荐面料="11.5-12oz 软化牛仔",
        推荐版型="轻宽直筒，避免极端宽",
        建议价格带="JPY 4,990-6,990",
        核心卖点="更好配西装外套和针织，替代重街头黑裤",
        关键风险="过宽、过低腰、拖地",
        日文搜索词="ブラックデニム メンズ / 黒 デニム メンズ ストレート / ワイドすぎない デニム メンズ",
    ),
    Candidate(
        sku_id="T01",
        产品名称="厚实不透白T",
        上架结论="上架",
        优先级="B",
        品类="T恤",
        风格线="基础内搭",
        场景="内搭 / 周末 / 通勤打底",
        对应第一步商品=None,
        对应第一步风格="简约男装",
        对应第一步场景="办公室休闲男装",
        对应第一步功能=None,
        amazon_group=None,
        amazon_title_patterns=r"",
        signal_patterns=r"白T|透け|ヘビーウェイト|タンクトップ|乳首|インナー",
        推荐颜色="白 / 浅灰",
        推荐面料="230-260gsm 棉质，偏干爽",
        推荐版型="常规略宽，领口不松",
        建议价格带="JPY 1,990-2,990",
        核心卖点="不透、可单穿、适合西装内搭",
        关键风险="太薄、领口软塌、发亮",
        日文搜索词="白T メンズ 厚手 / 白T 透けない メンズ / インナー Tシャツ メンズ",
    ),
    Candidate(
        sku_id="K01",
        产品名称="V领轻学院开衫",
        上架结论="上架",
        优先级="B",
        品类="开衫",
        风格线="轻学院胶囊",
        场景="学院风 / 通勤叠穿 / 周末",
        对应第一步商品=None,
        对应第一步风格="学院派男装",
        对应第一步场景=None,
        对应第一步功能=None,
        amazon_group="开衫",
        amazon_title_patterns=r"Vネック|スクールカーディガン|ビジネス|薄手",
        signal_patterns=r"カーディガン|スクール|プレッピー|アイビー|Vネック|紺ブレ",
        推荐颜色="海军蓝 / 深灰 / 米白",
        推荐面料="轻薄棉混纺或抗起球针织",
        推荐版型="合体但不紧身，门襟不窄",
        建议价格带="JPY 3,990-4,990",
        核心卖点="轻学院、好叠穿、和衬衫/牛仔兼容",
        关键风险="过度校服感、面料太薄、扣子廉价",
        日文搜索词="Vネック カーディガン メンズ / スクールカーディガン メンズ / プレッピー メンズ",
    ),
    Candidate(
        sku_id="J01",
        产品名称="短款牛仔夹克",
        上架结论="上架",
        优先级="C",
        品类="夹克",
        风格线="轻美式",
        场景="周末 / 叠穿 / 春季外套",
        对应第一步商品="男士夹克",
        对应第一步风格="美式休闲男装",
        对应第一步场景=None,
        对应第一步功能=None,
        amazon_group=None,
        amazon_title_patterns=r"",
        signal_patterns=r"デニムジャケット|Gジャン|短丈|アメカジ|春アウター",
        推荐颜色="中蓝 / 原色",
        推荐面料="11-12oz 牛仔布，洗水干净",
        推荐版型="短款、肩线自然、不过分宽大",
        建议价格带="JPY 5,990-7,990",
        核心卖点="把轻美式落成上身单品，适合四月层搭",
        关键风险="过短、做旧太重、偏复古戏剧化",
        日文搜索词="デニムジャケット メンズ 春 / 短丈 ジャケット メンズ / アメカジ メンズ",
    ),
    Candidate(
        sku_id="X01",
        产品名称="工装裤",
        上架结论="不上架",
        优先级="X",
        品类="长裤",
        风格线="重街头",
        场景="点缀",
        对应第一步商品=None,
        对应第一步风格="街头风男装",
        对应第一步场景=None,
        对应第一步功能=None,
        amazon_group="长裤",
        amazon_title_patterns=r"カーゴ",
        signal_patterns=r"カーゴ|cargo",
        推荐颜色="军绿 / 黑",
        推荐面料="",
        推荐版型="",
        建议价格带="",
        核心卖点="",
        关键风险="趋势回落，四月主线不匹配",
        日文搜索词="カーゴパンツ メンズ",
    ),
    Candidate(
        sku_id="X02",
        产品名称="重街头图案T",
        上架结论="不上架",
        优先级="X",
        品类="T恤",
        风格线="重街头",
        场景="点缀",
        对应第一步商品=None,
        对应第一步风格="街头风男装",
        对应第一步场景=None,
        对应第一步功能=None,
        amazon_group=None,
        amazon_title_patterns=r"",
        signal_patterns=r"ロゴ|グラフィック|ストリート",
        推荐颜色="",
        推荐面料="",
        推荐版型="",
        建议价格带="",
        核心卖点="",
        关键风险="内容热度还在，但购买主线已转向通勤和基础款",
        日文搜索词="ストリート Tシャツ メンズ",
    ),
    Candidate(
        sku_id="X03",
        产品名称="极宽拖地裤",
        上架结论="不上架",
        优先级="X",
        品类="长裤",
        风格线="极端宽松",
        场景="点缀",
        对应第一步商品=None,
        对应第一步风格=None,
        对应第一步场景=None,
        对应第一步功能=None,
        amazon_group="长裤",
        amazon_title_patterns=r"サルエル|袴|超ワイド",
        signal_patterns=r"ワイド|ダボ|袴|低身長",
        推荐颜色="",
        推荐面料="",
        推荐版型="",
        建议价格带="",
        核心卖点="",
        关键风险="评论明确担心显矮、拖地、太夸张",
        日文搜索词="超ワイド パンツ メンズ",
    ),
    Candidate(
        sku_id="X04",
        产品名称="过早凉感主卖点",
        上架结论="不上架",
        优先级="X",
        品类="功能",
        风格线="季节错位",
        场景="点缀",
        对应第一步商品=None,
        对应第一步风格=None,
        对应第一步场景=None,
        对应第一步功能="凉感男装",
        amazon_group=None,
        amazon_title_patterns=r"",
        signal_patterns=r"接触冷感|凉感",
        推荐颜色="",
        推荐面料="",
        推荐版型="",
        建议价格带="",
        核心卖点="",
        关键风险="四月仍偏早，适合五月后再放大",
        日文搜索词="接触冷感 メンズ",
    ),
]


def find_one(directory: Path, pattern: str) -> Path:
    matches = sorted(directory.glob(pattern))
    if not matches:
        raise FileNotFoundError(f"Missing file matching {pattern} in {directory}")
    return matches[0]


def latest_step2_workbook() -> Path:
    matches = sorted(p for p in STEP2_DIR.glob("*v3*.xlsx") if not p.name.startswith("~$"))
    if not matches:
        raise FileNotFoundError("Missing step2 v3 workbook")
    return matches[-1]


def load_step1_tables() -> dict[str, pd.DataFrame]:
    return {
        "商品": pd.read_csv(find_one(APRIL_DIR, "*商品机会表*.csv")),
        "风格": pd.read_csv(find_one(APRIL_DIR, "*风格机会表*.csv")),
        "功能": pd.read_csv(find_one(APRIL_DIR, "*功能机会表*.csv")),
        "场景": pd.read_csv(find_one(APRIL_DIR, "*场景机会表*.csv")),
    }


def load_step2_tables() -> dict[str, pd.DataFrame]:
    workbook = latest_step2_workbook()
    return {
        "综合判断": pd.read_excel(workbook, sheet_name=4),
        "来源总结": pd.read_excel(workbook, sheet_name=5),
        "Amazon": pd.read_excel(workbook, sheet_name=6),
        "YouTube评论": pd.read_excel(workbook, sheet_name=9),
        "论坛": pd.read_excel(workbook, sheet_name=10),
        "购买评论": pd.read_excel(workbook, sheet_name=11),
        "学院风专项": pd.read_excel(workbook, sheet_name=12),
        "最终建议": pd.read_excel(workbook, sheet_name=13),
    }


def normalize(value: float, lower: float, upper: float) -> float:
    if pd.isna(value):
        return 0.0
    if upper <= lower:
        return 0.0
    clipped = max(lower, min(float(value), upper))
    return (clipped - lower) / (upper - lower)


def row_score(row: pd.Series) -> float:
    april_score = normalize(row.get("4月季节指数", 0), 70, 160) * 40
    yoy_score = normalize(row.get("同比变化（%）", 0), -20, 25) * 20
    recent_score = normalize(row.get("近期变化（%）", 0), -25, 25) * 10
    anchor_score = normalize(row.get("相对锚点指数", 0), 0, 600) * 10
    judgment_score = JUDGMENT_POINTS.get(str(row.get("判断", "")), 20) * 0.2
    return round(april_score + yoy_score + recent_score + anchor_score + judgment_score, 2)


def get_row(table: pd.DataFrame, keyword: str | None) -> pd.Series | None:
    if not keyword:
        return None
    matches = table.loc[table["中文关键词"] == keyword]
    if matches.empty:
        return None
    return matches.iloc[0]


def parse_price(value: object) -> int | None:
    digits = re.findall(r"\d+", str(value).replace(",", ""))
    return int("".join(digits)) if digits else None


def sentiment_from_text(text: str) -> str:
    pos = sum(word in text for word in POSITIVE_WORDS)
    neg = sum(word in text for word in NEGATIVE_WORDS)
    if pos and neg:
        return "混合"
    if neg:
        return "风险"
    if pos:
        return "正向"
    return "中性"


def extract_tags(text: str) -> list[str]:
    tags = [tag for tag, keywords in SIGNAL_TAG_RULES.items() if any(word in text for word in keywords)]
    return tags or ["其他"]


def scene_from_text(text: str) -> str:
    if any(word in text for word in ["オフィスカジュアル", "ビジカジ", "通勤", "入社式", "就活", "工作", "通勤"]):
        return "办公室休闲 / 通勤"
    if any(word in text for word in ["プレッピー", "アイビー", "スクール", "学院"]):
        return "学院风"
    if any(word in text for word in ["デニム", "ジーンズ", "アメカジ"]):
        return "轻美式 / 日常"
    if any(word in text for word in ["春", "春服", "ライトアウター"]):
        return "春季换季"
    return "综合"


def action_from_tags(tags: list[str]) -> str:
    if "不透/低光泽" in tags:
        return "白色上装必须控透感和控光泽。"
    if "尺寸/版型" in tags or "裤长/拖地" in tags:
        return "版型、裤长和尺码说明必须写清。"
    if "免烫/易护理" in tags or "快干/透气" in tags:
        return "详情页优先写易护理、快干和舒适。"
    if "办公室休闲" in tags:
        return "先用通勤场景图和搭配建议成交。"
    if "学院风" in tags:
        return "只做轻学院胶囊，不做整套校服感。"
    return "作为辅助验证信号保留。"


def unified_signal_rows(step2_tables: dict[str, pd.DataFrame]) -> pd.DataFrame:
    rows: list[dict[str, str]] = []

    comments = step2_tables["YouTube评论"]
    for _, row in comments.iterrows():
        original = str(row.iloc[6])
        translated = str(row.iloc[8])
        combined = f"{row.iloc[0]} {original} {translated}"
        tags = extract_tags(combined)
        rows.append(
            {
                "来源": "YouTube评论",
                "来源主题": str(row.iloc[0]),
                "原文": original,
                "中文": translated,
                "场景判断": scene_from_text(combined),
                "标签": " / ".join(tags),
                "信号方向": sentiment_from_text(combined),
                "行动建议": action_from_tags(tags),
            }
        )

    forum = step2_tables["论坛"]
    for _, row in forum.iterrows():
        original = f"{row.iloc[3]} / {row.iloc[4]}"
        translated = f"{row.iloc[7]} / {row.iloc[8]}"
        combined = f"{row.iloc[2]} {original} {translated}"
        tags = extract_tags(combined)
        rows.append(
            {
                "来源": "论坛与问答",
                "来源主题": str(row.iloc[2]),
                "原文": original,
                "中文": translated,
                "场景判断": scene_from_text(combined),
                "标签": " / ".join(tags),
                "信号方向": sentiment_from_text(combined),
                "行动建议": action_from_tags(tags),
            }
        )

    reviews = step2_tables["购买评论"]
    for _, row in reviews.iterrows():
        original = f"{row.iloc[3]} / {row.iloc[4]}"
        translated = f"{row.iloc[7]} / {row.iloc[8]}"
        combined = f"{row.iloc[2]} {original} {translated}"
        tags = extract_tags(combined)
        rows.append(
            {
                "来源": "真实购买评论",
                "来源主题": str(row.iloc[2]),
                "原文": original,
                "中文": translated,
                "场景判断": scene_from_text(combined),
                "标签": " / ".join(tags),
                "信号方向": sentiment_from_text(combined),
                "行动建议": action_from_tags(tags),
            }
        )

    return pd.DataFrame(rows)


def candidate_relevant_rows(signal_df: pd.DataFrame, candidate: Candidate) -> pd.DataFrame:
    if not candidate.signal_patterns:
        return signal_df.iloc[0:0]
    text = (
        signal_df["来源主题"].fillna("")
        + " "
        + signal_df["原文"].fillna("")
        + " "
        + signal_df["中文"].fillna("")
        + " "
        + signal_df["标签"].fillna("")
    )
    mask = text.str.contains(candidate.signal_patterns, case=False, regex=True)
    return signal_df.loc[mask].copy()


def amazon_support_rows(amazon_df: pd.DataFrame, candidate: Candidate) -> pd.DataFrame:
    if not candidate.amazon_group:
        return amazon_df.iloc[0:0]
    group_mask = amazon_df.iloc[:, 2] == candidate.amazon_group
    rows = amazon_df.loc[group_mask].copy()
    if not candidate.amazon_title_patterns:
        return rows
    title_mask = rows.iloc[:, 4].fillna("").str.contains(candidate.amazon_title_patterns, case=False, regex=True)
    matched = rows.loc[title_mask]
    return matched if not matched.empty else rows.head(5)


def build_signal_detail(signal_df: pd.DataFrame) -> pd.DataFrame:
    candidates = []
    for _, row in signal_df.iterrows():
        matching = [candidate.产品名称 for candidate in CANDIDATES if candidate_relevant_rows(pd.DataFrame([row]), candidate).shape[0]]
        candidates.append(" / ".join(matching) if matching else "主线判断辅助")
    detail = signal_df.copy()
    detail.insert(4, "关联产品", candidates)
    return detail


def build_tag_summary(signal_df: pd.DataFrame) -> pd.DataFrame:
    exploded = signal_df.assign(标签=signal_df["标签"].str.split(" / ")).explode("标签")
    summary = (
        exploded.groupby(["来源", "标签", "信号方向"], dropna=False)
        .size()
        .reset_index(name="样本数")
        .sort_values(["样本数", "来源"], ascending=[False, True])
        .reset_index(drop=True)
    )
    return summary


def candidate_score_row(
    candidate: Candidate,
    step1_tables: dict[str, pd.DataFrame],
    step2_tables: dict[str, pd.DataFrame],
    signal_df: pd.DataFrame,
) -> dict[str, object]:
    supporting_rows = []
    support_texts = []
    for table_name, keyword in [
        ("商品", candidate.对应第一步商品),
        ("风格", candidate.对应第一步风格),
        ("场景", candidate.对应第一步场景),
        ("功能", candidate.对应第一步功能),
    ]:
        row = get_row(step1_tables[table_name], keyword)
        if row is not None:
            supporting_rows.append(row_score(row))
            support_texts.append(f"{table_name}:{keyword}={row['判断']}")

    step1_score = round(sum(supporting_rows) / len(supporting_rows), 2) if supporting_rows else 35.0

    amazon_rows = amazon_support_rows(step2_tables["Amazon"], candidate)
    amazon_count = len(amazon_rows)
    amazon_score = min(20.0, amazon_count * 2.0)

    relevant = candidate_relevant_rows(signal_df, candidate)
    signal_count = len(relevant)
    positive_count = int((relevant["信号方向"] == "正向").sum())
    mixed_count = int((relevant["信号方向"] == "混合").sum())
    risk_count = int((relevant["信号方向"] == "风险").sum())
    comment_score = min(25.0, signal_count * 1.2 + positive_count * 0.8 + mixed_count * 0.5)
    risk_penalty = min(15.0, risk_count * 1.5)

    final_score = round(step1_score * 0.55 + amazon_score + comment_score - risk_penalty, 2)

    key_tags = (
        relevant["标签"]
        .str.split(" / ")
        .explode()
        .value_counts()
        .head(4)
        .index.tolist()
    )

    return {
        "SKU编号": candidate.sku_id,
        "产品名称": candidate.产品名称,
        "上架结论": candidate.上架结论,
        "优先级": candidate.优先级,
        "第一步基础分": step1_score,
        "榜单验证分": round(amazon_score, 2),
        "评论验证分": round(comment_score, 2),
        "风险扣分": round(risk_penalty, 2),
        "最终评分": final_score,
        "相关评论样本": signal_count,
        "正向样本": positive_count,
        "风险样本": risk_count,
        "Amazon匹配商品数": amazon_count,
        "第一步支撑": "；".join(support_texts) if support_texts else "主要依赖第二步评论和榜单",
        "评论高频标签": " / ".join(key_tags) if key_tags else "",
        "建议价格带": candidate.建议价格带,
        "核心卖点": candidate.核心卖点,
        "关键风险": candidate.关键风险,
    }


def build_candidate_scores(
    step1_tables: dict[str, pd.DataFrame],
    step2_tables: dict[str, pd.DataFrame],
    signal_df: pd.DataFrame,
) -> pd.DataFrame:
    rows = [candidate_score_row(candidate, step1_tables, step2_tables, signal_df) for candidate in CANDIDATES]
    return pd.DataFrame(rows).sort_values(["上架结论", "优先级", "最终评分"], ascending=[True, True, False]).reset_index(drop=True)


def build_final_selection(candidate_scores: pd.DataFrame) -> pd.DataFrame:
    selected = candidate_scores.loc[candidate_scores["上架结论"] == "上架"].copy()
    meta = {candidate.sku_id: candidate for candidate in CANDIDATES}
    rows: list[dict[str, object]] = []
    for _, row in selected.iterrows():
        candidate = meta[row["SKU编号"]]
        rows.append(
            {
                "SKU编号": candidate.sku_id,
                "产品名称": candidate.产品名称,
                "优先级": candidate.优先级,
                "风格线": candidate.风格线,
                "场景": candidate.场景,
                "推荐颜色": candidate.推荐颜色,
                "推荐面料": candidate.推荐面料,
                "推荐版型": candidate.推荐版型,
                "建议价格带": candidate.建议价格带,
                "核心卖点": candidate.核心卖点,
                "关键风险": candidate.关键风险,
                "日文搜索词": candidate.日文搜索词,
                "最终评分": row["最终评分"],
            }
        )
    return pd.DataFrame(rows).sort_values(["优先级", "最终评分"], ascending=[True, False]).reset_index(drop=True)


def build_assortment_plan(final_df: pd.DataFrame) -> pd.DataFrame:
    distribution = {
        "A": ("主推库存", "60%", "衬衫 / 西裤 / 套装上衣 / 套装裤 / 原色直筒牛仔"),
        "B": ("补充库存", "30%", "黑色牛仔 / 厚实白T / V领开衫"),
        "C": ("测试库存", "10%", "短款牛仔夹克"),
    }
    rows = []
    for priority, (layer, budget, focus) in distribution.items():
        rows.append({"优先级": priority, "库存层级": layer, "建议预算占比": budget, "适用品类": focus})
    return pd.DataFrame(rows)


def build_exclusions(candidate_scores: pd.DataFrame) -> pd.DataFrame:
    meta = {candidate.sku_id: candidate for candidate in CANDIDATES}
    excluded = candidate_scores.loc[candidate_scores["上架结论"] == "不上架"].copy()
    rows = []
    for _, row in excluded.iterrows():
        candidate = meta[row["SKU编号"]]
        rows.append(
            {
                "方向": candidate.产品名称,
                "原因": candidate.关键风险,
                "来自第一步的依据": row["第一步支撑"],
                "来自第二步的依据": row["评论高频标签"],
            }
        )
    return pd.DataFrame(rows)


def build_overview(final_df: pd.DataFrame, signal_df: pd.DataFrame, candidate_scores: pd.DataFrame) -> pd.DataFrame:
    summary = final_df.groupby("优先级").size().to_dict()
    return pd.DataFrame(
        [
            {"项目": "第三步目标", "内容": "把第一步趋势、第二步多来源验证、评论拆分结果统一到最终上架 SKU。"},
            {"项目": "有效评论与帖子样本", "内容": f"{len(signal_df)} 条（YouTube / 论坛 / 购买评论合并）"},
            {"项目": "最终上架数量", "内容": f"{len(final_df)} 款，其中 A 级 {summary.get('A', 0)} 款，B 级 {summary.get('B', 0)} 款，C 级 {summary.get('C', 0)} 款。"},
            {"项目": "不上架方向", "内容": f"{len(candidate_scores.loc[candidate_scores['上架结论']=='不上架'])} 类，已单独列出。"},
        ]
    )


def build_source_reference(step2_tables: dict[str, pd.DataFrame]) -> pd.DataFrame:
    workbook = latest_step2_workbook()
    return pd.DataFrame(
        [
            {"来源层级": "第一步", "文件": str(find_one(APRIL_DIR, "*商品机会表*.csv")), "用途": "四月商品机会判断"},
            {"来源层级": "第一步", "文件": str(find_one(APRIL_DIR, "*风格机会表*.csv")), "用途": "四月风格方向"},
            {"来源层级": "第一步", "文件": str(find_one(APRIL_DIR, "*功能机会表*.csv")), "用途": "功能卖点方向"},
            {"来源层级": "第一步", "文件": str(find_one(APRIL_DIR, "*场景机会表*.csv")), "用途": "通勤和四月场景"},
            {"来源层级": "第二步", "文件": str(workbook), "用途": "Amazon / YouTube / 论坛 / 购买评论多来源验证"},
        ]
    )


def autosize(worksheet, dataframe: pd.DataFrame) -> None:
    for idx, col in enumerate(dataframe.columns):
        values = dataframe[col].astype(str).tolist()
        max_len = max([len(str(col))] + [len(v) for v in values])
        worksheet.set_column(idx, idx, min(max_len + 2, 90))


def write_sheet(writer: pd.ExcelWriter, name: str, df: pd.DataFrame) -> None:
    df.to_excel(writer, sheet_name=name, index=False)
    ws = writer.sheets[name]
    autosize(ws, df)
    ws.freeze_panes(1, 0)


def write_markdown(final_df: pd.DataFrame, candidate_scores: pd.DataFrame) -> None:
    lines = [
        "# 日本男装第三步：具体选品",
        "",
        f"- 分析日期：{EXPORT_DATE}",
        f"- 最终上架：{len(final_df)} 款",
        "",
        "## 主推结论",
        "",
    ]
    for _, row in final_df.iterrows():
        lines.append(f"- {row['产品名称']}：{row['核心卖点']}；价格带 {row['建议价格带']}。")

    lines.extend(["", "## 不上架方向", ""])
    for _, row in candidate_scores.loc[candidate_scores["上架结论"] == "不上架"].iterrows():
        lines.append(f"- {row['产品名称']}：{row['关键风险']}。")

    OUTPUT_MD.write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    step1_tables = load_step1_tables()
    step2_tables = load_step2_tables()

    signal_df = unified_signal_rows(step2_tables)
    signal_detail_df = build_signal_detail(signal_df)
    tag_summary_df = build_tag_summary(signal_detail_df)
    candidate_scores_df = build_candidate_scores(step1_tables, step2_tables, signal_detail_df)
    final_df = build_final_selection(candidate_scores_df)
    assortment_df = build_assortment_plan(final_df)
    exclusions_df = build_exclusions(candidate_scores_df)
    overview_df = build_overview(final_df, signal_detail_df, candidate_scores_df)
    source_ref_df = build_source_reference(step2_tables)

    with pd.ExcelWriter(OUTPUT_XLSX, engine="xlsxwriter") as writer:
        write_sheet(writer, "概览", overview_df)
        write_sheet(writer, "评论拆分_明细", signal_detail_df)
        write_sheet(writer, "评论标签汇总", tag_summary_df)
        write_sheet(writer, "候选产品评分", candidate_scores_df)
        write_sheet(writer, "最终上架产品", final_df)
        write_sheet(writer, "组合结构建议", assortment_df)
        write_sheet(writer, "不上架方向", exclusions_df)
        write_sheet(writer, "来源引用", source_ref_df)

    write_markdown(final_df, candidate_scores_df)


if __name__ == "__main__":
    main()
