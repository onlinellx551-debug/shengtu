from __future__ import annotations

import hashlib
import json
import math
import re
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
import requests
from PIL import Image, ImageDraw, ImageOps
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
STEP5_DIR = ROOT / "step5_output"
STEP6_DIR = ROOT / "step6_output"
STEP6_DIR.mkdir(exist_ok=True)

CONFIRM_XLSX = next(path for path in STEP5_DIR.glob("T01_*确认版.xlsx") if not path.name.startswith("~$"))
EXPORT_DATE = "2026-03-20"
CDP_URL = "http://127.0.0.1:9222"

PACK_DIR = STEP6_DIR / "T01_白T素材包"
PACK_DIR.mkdir(exist_ok=True)
REPORT_XLSX = STEP6_DIR / f"T01_白T_第6步素材包_{EXPORT_DATE}.xlsx"
REPORT_MD = STEP6_DIR / f"T01_白T_第6步素材包说明_{EXPORT_DATE}.md"
REPORT_JSON = STEP6_DIR / f"T01_白T_第6步素材包明细_{EXPORT_DATE}.json"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/146.0.0.0 Safari/537.36"
    ),
    "Accept": "*/*",
}


def clean_lines(text: str) -> list[str]:
    return [line.strip() for line in re.split(r"[\r\n]+", str(text)) if line.strip()]


def label_value(lines: list[str], label: str) -> str:
    for index, line in enumerate(lines):
        if line == label and index + 1 < len(lines):
            return lines[index + 1]
        if line.startswith(f"{label}:") or line.startswith(f"{label}："):
            return re.split(r"[:：]", line, maxsplit=1)[1].strip()
    return ""


def number_from_text(text: str) -> float:
    content = str(text or "").strip().replace("%", "")
    if not content or content == "-":
        return 0.0
    if content.startswith("<"):
        matched = re.search(r"[0-9.]+", content)
        return float(matched.group(0)) / 2 if matched else 0.0
    matched = re.search(r"[0-9.]+", content)
    return float(matched.group(0)) if matched else 0.0


def sales_to_num(text: str) -> float:
    content = str(text or "").strip()
    multiplier = 1.0
    if "万" in content:
        multiplier = 10000.0
    elif "千" in content:
        multiplier = 1000.0
    return number_from_text(content) * multiplier


def canonical_asset_url(url: str) -> str:
    clean = str(url or "").split("?")[0]
    clean = clean.replace(".220x220.jpg", ".jpg")
    clean = clean.replace(".310x310.jpg", ".jpg")
    clean = clean.replace(".search.jpg", ".jpg")
    clean = clean.replace(".summ.jpg", ".jpg")
    clean = clean.replace(".jpg_.webp", ".jpg")
    clean = clean.replace(".png_.webp", ".png")
    clean = clean.replace(".jpg_b.jpg", ".jpg")
    clean = clean.replace(".png_b.jpg", ".png")
    clean = re.sub(r"(\.(?:jpg|png|webp))_250x250\.jpg$", r"\1", clean)
    return clean


def classify_asset(url: str) -> str:
    if url.endswith(".mp4"):
        return "视频"
    if "_sum.jpg" in url:
        return "详情长图"
    return "主图细节图"


def download_file(url: str, referer: str, out_dir: Path, prefix: str) -> str:
    ext_match = re.search(r"\.(jpg|jpeg|png|webp|mp4)$", canonical_asset_url(url), flags=re.I)
    ext = f".{ext_match.group(1).lower()}" if ext_match else ".bin"
    ext = ".jpg" if ext == ".jpeg" else ext
    digest = hashlib.md5(url.encode("utf-8")).hexdigest()[:12]
    out_path = out_dir / f"{prefix}_{digest}{ext}"
    if out_path.exists():
        return str(out_path)
    headers = dict(HEADERS)
    headers["Referer"] = referer
    response = requests.get(url, headers=headers, timeout=60)
    response.raise_for_status()
    out_path.write_bytes(response.content)
    return str(out_path)


def extract_probe(page, supplier: str, product_link: str) -> dict[str, Any]:
    page.goto(product_link, wait_until="domcontentloaded", timeout=90000)
    page.wait_for_timeout(5000)
    for _ in range(6):
        page.mouse.wheel(0, 1400)
        page.wait_for_timeout(900)
    body = page.locator("body").inner_text()
    html = page.content()

    if "访问行为存在异常" in body or "安全验证" in body:
        return {
            "供应商": supplier,
            "商品链接": product_link,
            "probe_status": "blocked",
            "页面标题": page.title(),
            "正文前段": body[:500],
        }

    lines = clean_lines(body)
    image_urls = sorted(
        {
            canonical_asset_url(url)
            for url in re.findall(r"https?://(?:cbu\d+|img)\.alicdn\.com/[^\"'\s>]+(?:jpg|jpeg|png|webp)", html)
            if ("img/ibank/" in url or "-0-cib." in url or "_sum.jpg" in url)
        }
    )
    video_urls = sorted(set(re.findall(r"https?://[^\"'\s>]+\.mp4", html)))

    review_section = ""
    review_match = re.search(r"商品评价(.*?)商品属性", body, flags=re.S)
    if review_match:
        review_section = review_match.group(1).strip()[:1200]

    return {
        "供应商": supplier,
        "商品链接": product_link,
        "probe_status": "ok",
        "页面标题": page.title(),
        "正文前段": body[:2400],
        "店铺回头率": label_value(lines, "店铺回头率"),
        "店铺服务分": label_value(lines, "店铺服务分"),
        "准时发货率": label_value(lines, "准时发货率"),
        "店铺好评率": label_value(lines, "店铺好评率"),
        "面料名称": label_value(lines, "面料名称"),
        "主面料成分": label_value(lines, "主面料成分"),
        "克重": label_value(lines, "克重"),
        "版型": label_value(lines, "版型"),
        "领型": label_value(lines, "领型") or label_value(lines, "领口形状"),
        "厚薄": label_value(lines, "厚薄"),
        "图案": label_value(lines, "图案"),
        "颜色": label_value(lines, "颜色"),
        "尺码": label_value(lines, "尺码"),
        "有无质检报告": label_value(lines, "有无质检报告"),
        "是否跨境专供": label_value(lines, "是否跨境出口专供货源"),
        "有可授权品牌": label_value(lines, "有可授权的自有品牌"),
        "评价条数": re.search(r"(\d+\+?)条评价", body).group(1) if re.search(r"(\d+\+?)条评价", body) else "",
        "好评率": label_value(lines, "好评率"),
        "贴牌换标": "是" if ("贴牌换标" in body or "更换客户标" in body) else "否",
        "包装定制": "是" if "包装定制" in body else "否",
        "个性定制": "是" if ("个性定制" in body or "Logo定制" in body or "来图加工" in body) else "否",
        "素材图片数": len(image_urls),
        "素材视频数": len(video_urls),
        "素材图片链接_json": json.dumps(image_urls, ensure_ascii=False),
        "素材视频链接_json": json.dumps(video_urls, ensure_ascii=False),
        "评论摘录": review_section,
    }


def compute_scores(df: pd.DataFrame) -> pd.DataFrame:
    scored = df.copy()
    scored["价格数"] = pd.to_numeric(scored["价格"], errors="coerce").fillna(0.0)
    scored["店铺服务分数"] = scored["店铺服务分"].apply(number_from_text)
    scored["店铺回头率数"] = scored["店铺回头率"].apply(number_from_text)
    scored["店铺好评率数"] = scored["店铺好评率"].apply(number_from_text)
    scored["好评率数"] = scored["好评率"].apply(number_from_text)
    scored["评价条数数"] = scored["评价条数"].apply(sales_to_num)
    scored["厚度匹配分"] = 0.0
    scored.loc[scored["克重"].astype(str).str.contains("300", na=False), "厚度匹配分"] += 18
    scored.loc[scored["克重"].astype(str).str.contains("260", na=False), "厚度匹配分"] += 14
    scored.loc[scored["克重"].astype(str).str.contains("230", na=False), "厚度匹配分"] += 12
    scored.loc[scored["厚薄"].astype(str).str.contains("加厚", na=False), "厚度匹配分"] += 10
    scored.loc[scored["厚薄"].astype(str).str.contains("薄款", na=False), "厚度匹配分"] -= 18
    scored.loc[scored["面料名称"].astype(str).str.contains("纯棉|精梳棉|棉", na=False), "厚度匹配分"] += 8
    scored.loc[scored["面料名称"].astype(str).str.contains("牛奶丝", na=False), "厚度匹配分"] -= 20

    scored["质量稳定分"] = (
        scored["店铺服务分数"] * 5
        + scored["店铺好评率数"] / 6
        + scored["好评率数"] / 8
        + scored["店铺回头率数"] / 10
        + scored["评价条数数"].apply(lambda value: min(14.0, math.log10(value + 1) * 5))
    )
    scored.loc[scored["有无质检报告"].eq("是"), "质量稳定分"] += 8
    scored.loc[scored["评论摘录"].astype(str).str.contains("尺寸要正常偏小一码|版型 质量 售后", na=False), "质量稳定分"] -= 8

    scored["供货友好分"] = 0.0
    scored.loc[scored["起批量"].astype(str).str.contains("1件|2件", na=False), "供货友好分"] += 8
    scored.loc[scored["贴牌换标"].eq("是"), "供货友好分"] += 8
    scored.loc[scored["包装定制"].eq("是"), "供货友好分"] += 5
    scored.loc[scored["个性定制"].eq("是"), "供货友好分"] += 5
    scored.loc[scored["是否跨境专供"].eq("是"), "供货友好分"] += 4

    scored["素材完整度分"] = (
        scored["素材图片数"].fillna(0).apply(lambda value: min(24.0, value * 0.35))
        + scored["素材视频数"].fillna(0) * 8
        + scored["评论摘录"].astype(str).str.len().clip(upper=400) / 50
    )

    scored["价格竞争力分"] = 0.0
    scored.loc[(scored["价格数"] >= 10) & (scored["价格数"] <= 18), "价格竞争力分"] = 12
    scored.loc[(scored["价格数"] > 18) & (scored["价格数"] <= 28), "价格竞争力分"] = 8
    scored.loc[(scored["价格数"] > 28) | (scored["价格数"] < 10), "价格竞争力分"] = 5

    scored["综合总分"] = (
        scored["厚度匹配分"]
        + scored["质量稳定分"]
        + scored["供货友好分"]
        + scored["素材完整度分"]
        + scored["价格竞争力分"]
    ).round(2)
    return scored.sort_values(["综合总分", "素材完整度分"], ascending=[False, False]).reset_index(drop=True)


def role_for_supplier(row: pd.Series, ranked: pd.DataFrame) -> str:
    supplier = row["供应商"]
    if supplier == ranked.iloc[0]["供应商"]:
        return "主供货商"
    if supplier == ranked.iloc[1]["供应商"]:
        return "素材补充源1"
    if supplier == ranked.iloc[2]["供应商"]:
        return "素材补充源2"
    return "低优先同款源"


def contact_sheet(title: str, image_paths: list[str], out_path: Path) -> None:
    valid = [Path(path) for path in image_paths if path and Path(path).exists()]
    if not valid:
        return
    cards: list[Image.Image] = []
    for index, path in enumerate(valid[:12], start=1):
        image = Image.open(path).convert("RGB")
        image = ImageOps.contain(image, (220, 220))
        canvas = Image.new("RGB", (240, 260), "white")
        canvas.paste(image, ((240 - image.width) // 2, 10))
        draw = ImageDraw.Draw(canvas)
        draw.text((8, 232), str(index), fill="black")
        cards.append(canvas)
    cols = 3
    rows = math.ceil(len(cards) / cols)
    sheet = Image.new("RGB", (cols * 240, rows * 260 + 40), "#f4f4f4")
    draw = ImageDraw.Draw(sheet)
    draw.text((12, 10), title, fill="black")
    for index, image in enumerate(cards):
        x = (index % cols) * 240
        y = 40 + (index // cols) * 260
        sheet.paste(image, (x, y))
    sheet.save(out_path, quality=92)


def download_assets(ranked: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for _, supplier_row in ranked.iterrows():
        role = role_for_supplier(supplier_row, ranked)
        folder = PACK_DIR / role / supplier_row["供应商"]
        folder.mkdir(parents=True, exist_ok=True)
        main_dir = folder / "主图细节图"
        detail_dir = folder / "详情长图"
        video_dir = folder / "视频"
        for sub in [main_dir, detail_dir, video_dir]:
            sub.mkdir(exist_ok=True)

        image_urls = json.loads(supplier_row["素材图片链接_json"]) if supplier_row["素材图片链接_json"] else []
        video_urls = json.loads(supplier_row["素材视频链接_json"]) if supplier_row["素材视频链接_json"] else []
        main_paths: list[str] = []
        detail_paths: list[str] = []
        video_paths: list[str] = []

        for index, url in enumerate(image_urls, start=1):
            asset_type = classify_asset(url)
            target_dir = detail_dir if asset_type == "详情长图" else main_dir
            local_path = download_file(url, supplier_row["商品链接"], target_dir, f"{asset_type}_{index:02d}")
            rows.append(
                {
                    "角色": role,
                    "供应商": supplier_row["供应商"],
                    "商品链接": supplier_row["商品链接"],
                    "素材类型": asset_type,
                    "素材链接": url,
                    "本地文件": local_path,
                }
            )
            if asset_type == "详情长图":
                detail_paths.append(local_path)
            else:
                main_paths.append(local_path)

        for index, url in enumerate(video_urls, start=1):
            local_path = download_file(url, supplier_row["商品链接"], video_dir, f"视频_{index:02d}")
            rows.append(
                {
                    "角色": role,
                    "供应商": supplier_row["供应商"],
                    "商品链接": supplier_row["商品链接"],
                    "素材类型": "视频",
                    "素材链接": url,
                    "本地文件": local_path,
                }
            )
            video_paths.append(local_path)

        contact_sheet(f"{role} | {supplier_row['供应商']} | 主图细节图", main_paths, folder / "主图联系表.jpg")
        contact_sheet(f"{role} | {supplier_row['供应商']} | 详情长图", detail_paths, folder / "详情联系表.jpg")
    return pd.DataFrame(rows)


def build_copy(ranked: pd.DataFrame) -> pd.DataFrame:
    main = ranked.iloc[0]
    title_zh = "厚实不透基础白T"
    title_ja = "透けにくい ヘビーウェイト 白T メンズ"
    subtitle_ja = "厚手コットンで1枚でも着やすい、春夏の定番ベーシックTシャツ"
    bullets = [
        "白でも透けにくい厚手生地で、1枚でも着やすい",
        "圆领利落、版型干净，适合单穿也适合西装内搭",
        "基础宽松而不过分夸张，适合日本四月日常和通勤场景",
        "同款素材已覆盖主图、细节图、详情长图和视频",
    ]
    return pd.DataFrame(
        [
            {"字段": "中文标题", "内容": title_zh},
            {"字段": "日文标题", "内容": title_ja},
            {"字段": "日文副标题", "内容": subtitle_ja},
            {"字段": "建议主卖点1", "内容": bullets[0]},
            {"字段": "建议主卖点2", "内容": bullets[1]},
            {"字段": "建议主卖点3", "内容": bullets[2]},
            {"字段": "页面素材来源", "内容": f"主供货商 {main['供应商']}，其余同款源补细节和视频"},
            {"字段": "主供参数", "内容": f"{main['面料名称']} / {main['主面料成分']} / {main['克重']} / {main['尺码']}"},
        ]
    )


def page_structure() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"模块": "首屏主图", "建议": "8-10张，优先白色平铺主图、上身图、领口细节、厚度图"},
            {"模块": "视频", "建议": "放1个短视频在主图区域后半段"},
            {"模块": "卖点图", "建议": "从同款详情长图中选厚度、领口、面料、版型说明"},
            {"模块": "参数区", "建议": "放颜色、尺码、克重、洗护、材质"},
            {"模块": "长图详情", "建议": "优先使用主供和素材补充源的 `_sum.jpg` 长图"},
        ]
    )


def markdown_summary(ranked: pd.DataFrame, assets: pd.DataFrame) -> str:
    top = ranked.iloc[0]
    second = ranked.iloc[1]
    third = ranked.iloc[2]
    return "\n".join(
        [
            "# T01 白T 第6步素材包",
            "",
            f"- 主供货商：{top['供应商']} | {top['商品链接']}",
            f"- 素材补充源1：{second['供应商']} | {second['商品链接']}",
            f"- 素材补充源2：{third['供应商']} | {third['商品链接']}",
            "",
            "## 结论",
            "",
            f"- 主供货商定为 `{top['供应商']}`，原因是厚度和面料匹配度最高，300g 纯棉更符合“厚实不透白T”的目标。",
            f"- `{second['供应商']}` 作为素材补充源1，原因是视频、评论、质检和贴牌能力更完整，适合补页面素材。",
            f"- `{third['供应商']}` 作为素材补充源2，原因是同款素材量大，评论正向，适合补主图和细节图。",
            "",
            "## 素材统计",
            "",
            f"- 已下载素材总数：{len(assets)}",
            f"- 其中图片/长图：{len(assets[assets['素材类型'] != '视频'])}",
            f"- 其中视频：{len(assets[assets['素材类型'] == '视频'])}",
        ]
    )


def main() -> None:
    confirmed_df = pd.read_excel(CONFIRM_XLSX, sheet_name=1)
    confirmed_df = confirmed_df[["搜索来源", "AlphaShop排名", "标题", "供应商", "价格", "销量信号", "起批量", "发货地", "AI亮点", "图片链接", "商品链接"]].copy()

    probes: list[dict[str, Any]] = []
    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(CDP_URL)
        context = browser.contexts[0]
        page = context.new_page()
        try:
            for _, row in confirmed_df.iterrows():
                probes.append(extract_probe(page, row["供应商"], row["商品链接"]))
                page.wait_for_timeout(5000)
        finally:
            page.close()

    probe_df = pd.DataFrame(probes)
    merged = confirmed_df.merge(probe_df, on=["供应商", "商品链接"], how="left")
    ranked = compute_scores(merged)
    assets = download_assets(ranked)
    copy_df = build_copy(ranked)
    structure_df = page_structure()

    with pd.ExcelWriter(REPORT_XLSX, engine="xlsxwriter") as writer:
        ranked.to_excel(writer, sheet_name="供应商评分", index=False)
        assets.to_excel(writer, sheet_name="素材清单", index=False)
        copy_df.to_excel(writer, sheet_name="文案建议", index=False)
        structure_df.to_excel(writer, sheet_name="页面结构建议", index=False)
        merged.to_excel(writer, sheet_name="详情探测原始", index=False)
        for ws in writer.sheets.values():
            ws.freeze_panes(1, 0)

    REPORT_MD.write_text(markdown_summary(ranked, assets), encoding="utf-8")
    REPORT_JSON.write_text(
        json.dumps(
            {
                "ranked_suppliers": ranked.to_dict(orient="records"),
                "assets": assets.to_dict(orient="records"),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(REPORT_XLSX)
    print(REPORT_MD)
    print(REPORT_JSON)


if __name__ == "__main__":
    main()
