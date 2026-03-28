from __future__ import annotations

import hashlib
import json
import math
import re
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
import requests
from PIL import Image, ImageDraw, ImageOps
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
STEP4_DIR = ROOT / "step4_output"
STEP5_DIR = ROOT / "step5_output"
STEP5_DIR.mkdir(exist_ok=True)
EXPORT_DATE = "2026-03-20"
CDP_URL = "http://127.0.0.1:9222"
HOME_URL = "https://create.alphashop.cn/"
RAW_STEP4_CSV = STEP4_DIR / "alphashop_supplier_raw_2026-03-20.csv"
SELECTED_XLS = next(path for path in STEP5_DIR.glob("*.xls") if not path.name.startswith("~$"))

T01_ASSET_DIR = STEP5_DIR / "T01_素材包"
SEARCH_DIR = T01_ASSET_DIR / "搜索缓存"
DOWNLOAD_DIR = T01_ASSET_DIR / "下载素材"
REPORT_XLSX = STEP5_DIR / f"T01_供应商与素材总表_{EXPORT_DATE}.xlsx"
REPORT_JSON = STEP5_DIR / f"T01_供应商与素材明细_{EXPORT_DATE}.json"
REPORT_MD = STEP5_DIR / f"T01_商品详情页素材包_{EXPORT_DATE}.md"

for folder in [T01_ASSET_DIR, SEARCH_DIR, DOWNLOAD_DIR]:
    folder.mkdir(parents=True, exist_ok=True)


REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/146.0.0.0 Safari/537.36"
    ),
    "Accept": "*/*",
}


@dataclass(frozen=True)
class SearchTask:
    search_id: str
    search_type: str
    prompt: str
    image_path: Path | None = None


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
    content = str(text or "").replace("%", "").strip()
    if not content or content == "-":
        return 0.0
    if content.startswith("<"):
        match = re.search(r"[0-9.]+", content)
        return float(match.group(0)) / 2 if match else 0.0
    match = re.search(r"[0-9.]+", content)
    return float(match.group(0)) if match else 0.0


def sales_to_num(text: str) -> float:
    content = str(text or "").strip()
    if not content:
        return 0.0
    multiplier = 1.0
    if "万" in content:
        multiplier = 10000.0
    elif "千" in content:
        multiplier = 1000.0
    return number_from_text(content) * multiplier


def pct_to_num(text: str) -> float:
    return number_from_text(text)


def moq_to_num(text: str) -> float:
    return number_from_text(text)


def parse_search_row(row_text: str) -> dict[str, Any]:
    lines = clean_lines(row_text)
    rank = lines[0] if lines and lines[0].isdigit() else ""
    title = lines[1] if len(lines) > 1 else ""
    price = ""
    price_match = re.search(r"¥\s*([0-9.]+)", row_text)
    if price_match:
        price = price_match.group(1)

    supplier = ""
    supplier_tags: list[str] = []
    if title:
        title_index = lines.index(title)
        for line in lines[title_index + 1 :]:
            if line.startswith("¥"):
                continue
            if line.startswith("✨") or line.startswith("❗"):
                break
            supplier = line
            break
        if supplier:
            supplier_index = lines.index(supplier)
            for line in lines[supplier_index + 1 :]:
                if line.startswith("✨") or line.startswith("❗"):
                    break
                if "年" in line or "工厂" in line or "商家" in line:
                    supplier_tags.append(line)

    return {
        "alpha_rank": rank,
        "title": title,
        "price": price,
        "supplier": supplier,
        "supplier_tags": " / ".join(supplier_tags[:4]),
        "highlights": " / ".join([line for line in lines if line.startswith(("✨", "❗"))][:2]),
        "sales_signal": label_value(lines, "近一年全网销量"),
        "moq": label_value(lines, "起批量"),
        "origin": label_value(lines, "发货地"),
        "service_score": label_value(lines, "综合服务分"),
        "response_rate": label_value(lines, "客服响应率"),
        "repeat_rate": label_value(lines, "90天回头率"),
        "refund_rate": label_value(lines, "品质退款率"),
        "raw_text": row_text,
    }


def match_score(title: str, preferred: list[str], avoided: list[str]) -> float:
    title_lower = str(title).lower()
    score = 0.0
    for word in preferred:
        if word.lower() in title_lower:
            score += 8.0
    for word in avoided:
        if word.lower() in title_lower:
            score -= 20.0
    return score


def build_candidate_key(row: pd.Series) -> str:
    link = str(row.get("product_link", "")).strip()
    if "detail.1688.com/offer/" in link:
        return link
    return f"{row.get('supplier', '')}||{row.get('title', '')}"


def latest_selected_row() -> pd.Series:
    selected_df = pd.read_excel(SELECTED_XLS, sheet_name=0)
    return selected_df.iloc[0]


def prepare_local_image(image_url: str, filename: str, referer: str | None = None) -> Path:
    out_path = SEARCH_DIR / filename
    if out_path.exists():
        return out_path
    headers = dict(REQUEST_HEADERS)
    if referer:
        headers["Referer"] = referer
    response = requests.get(image_url, headers=headers, timeout=30)
    response.raise_for_status()
    out_path.write_bytes(response.content)
    return out_path


def open_findshop(page) -> None:
    page.goto(HOME_URL, wait_until="domcontentloaded", timeout=90000)
    page.wait_for_timeout(7000)
    page.mouse.click(620, 257)
    page.wait_for_timeout(2500)


def fill_findshop_prompt(page, prompt: str) -> None:
    editor = page.locator('[contenteditable="true"]').first
    editor.click()
    editor.evaluate(
        """
        (el, value) => {
          el.focus();
          el.innerHTML = '';
          el.textContent = value;
          el.dispatchEvent(new InputEvent('input', { bubbles: true, data: value, inputType: 'insertText' }));
        }
        """,
        prompt,
    )


def click_send_button(page) -> None:
    button = page.locator("button.sendButton--zLQHd42x")
    if button.count():
        button.click()
        return
    buttons = page.locator("button")
    for index in range(buttons.count()):
        box = buttons.nth(index).bounding_box()
        if box and box["width"] <= 40 and box["height"] <= 40 and box["y"] < 500:
            buttons.nth(index).click()
            return
    raise RuntimeError("找不到 AlphaShop 发送按钮。")


def submit_search(page, task: SearchTask) -> str:
    open_findshop(page)
    fill_findshop_prompt(page, task.prompt)
    if task.image_path:
        page.locator('input[type="file"]').set_input_files(str(task.image_path))
        page.wait_for_timeout(4000)
    click_send_button(page)
    page.wait_for_url("**/sourcing**", timeout=90000)
    page.wait_for_timeout(6000)
    confirm = page.locator("button.preSelectBusinessBtn--HBrU7smX")
    if confirm.count():
        confirm.click()
    for _ in range(90):
        body = page.locator("body").inner_text()
        if "供应商推荐" in body and page.locator("tr").count() >= 4:
            return page.url
        page.wait_for_timeout(2000)
    raise TimeoutError(f"{task.search_id} 未在预期时间内返回找商结果。")


def extract_product_link(page, row_locator) -> str:
    image = row_locator.locator("img").first
    if image.count() == 0:
        return ""
    try:
        with page.expect_popup(timeout=12000) as popup_info:
            image.click()
        popup = popup_info.value
        popup.wait_for_timeout(2500)
        url = popup.url
        popup.close()
        page.wait_for_timeout(1200)
        return url
    except PlaywrightTimeoutError:
        return ""


def extract_search_results(page, task: SearchTask) -> list[dict[str, Any]]:
    rows = page.locator("tr")
    results: list[dict[str, Any]] = []
    for index in range(2, rows.count()):
        row = rows.nth(index)
        row_text = row.inner_text().strip()
        if not row_text:
            continue
        parsed = parse_search_row(row_text)
        parsed.update(
            {
                "search_id": task.search_id,
                "search_type": task.search_type,
                "prompt": task.prompt,
                "search_session_url": page.url,
                "image_search_path": str(task.image_path) if task.image_path else "",
                "image_url": row.locator("img").first.get_attribute("src") or "",
                "product_link": extract_product_link(page, row),
            }
        )
        results.append(parsed)
    return results


def run_alphashop_searches(tasks: list[SearchTask]) -> tuple[pd.DataFrame, pd.DataFrame]:
    rows: list[dict[str, Any]] = []
    sessions: list[dict[str, Any]] = []
    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(CDP_URL)
        context = browser.contexts[0]
        page = context.new_page()
        try:
            for task in tasks:
                session_url = submit_search(page, task)
                results = extract_search_results(page, task)
                rows.extend(results)
                sessions.append(
                    {
                        "search_id": task.search_id,
                        "search_type": task.search_type,
                        "prompt": task.prompt,
                        "image_path": str(task.image_path) if task.image_path else "",
                        "result_count": len(results),
                        "session_url": session_url,
                    }
                )
                page.wait_for_timeout(6000)
        finally:
            page.close()
    return pd.DataFrame(rows), pd.DataFrame(sessions)


def parse_detail_page_text(text: str) -> dict[str, Any]:
    lines = clean_lines(text)
    review_count_match = re.search(r"(\d+\+?)条评价", text)
    return {
        "店铺回头率": label_value(lines, "店铺回头率"),
        "店铺服务分": label_value(lines, "店铺服务分"),
        "准时发货率": label_value(lines, "准时发货率"),
        "店铺好评率": label_value(lines, "店铺好评率"),
        "商品评分": number_from_text(text.split("商品属性")[0]),
        "评价条数": review_count_match.group(1) if review_count_match else "",
        "好评率": label_value(lines, "好评率"),
        "面料名称": label_value(lines, "面料名称"),
        "主面料成分": label_value(lines, "主面料成分"),
        "克重": label_value(lines, "克重"),
        "版型": label_value(lines, "版型"),
        "领型": label_value(lines, "领型") or label_value(lines, "领口形状"),
        "图案": label_value(lines, "图案"),
        "货源类型": label_value(lines, "货源类型"),
        "是否跨境专供": label_value(lines, "是否跨境出口专供货源"),
        "有无质检报告": label_value(lines, "有无质检报告"),
        "有可授权品牌": label_value(lines, "有可授权的自有品牌"),
        "颜色": label_value(lines, "颜色"),
        "尺码": label_value(lines, "尺码"),
        "包装定制": "是" if "包装定制" in text else "否",
        "贴牌换标": "是" if ("贴牌换标" in text or "更换客户标" in text) else "否",
        "个性定制": "是" if ("个性定制" in text or "Logo定制" in text or "来图加工" in text) else "否",
    }


def extract_material_urls(html: str) -> tuple[list[str], list[str]]:
    image_candidates = sorted(
        set(
            re.findall(
                r"https?://(?:cbu\d+|img)\.alicdn\.com/[^\"'\s>]+(?:jpg|jpeg|png|webp)",
                html,
            )
        )
    )
    video_candidates = sorted(set(re.findall(r"https?://[^\"'\s>]+\.mp4", html)))
    images = [
        url
        for url in image_candidates
        if (
            "img/ibank/" in url
            or "-0-cib." in url
            or "_sum.jpg" in url
            or "/ii" in url
        )
    ]
    return images, video_candidates


def probe_detail_pages(candidates: pd.DataFrame) -> pd.DataFrame:
    probed_rows: list[dict[str, Any]] = []
    unique_candidates = candidates.drop_duplicates(subset=["candidate_key"]).copy()

    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(CDP_URL)
        context = browser.contexts[0]
        page = context.new_page()
        try:
            for _, row in unique_candidates.iterrows():
                product_link = str(row["product_link"] or "").strip()
                if "detail.1688.com/offer/" not in product_link:
                    probed_rows.append(
                        {
                            "candidate_key": row["candidate_key"],
                            "product_link": product_link,
                            "probe_status": "invalid_link",
                        }
                    )
                    continue

                page.goto(product_link, wait_until="domcontentloaded", timeout=90000)
                page.wait_for_timeout(5000)
                body = page.locator("body").inner_text()
                if "访问行为存在异常" in body or "安全验证" in body:
                    probed_rows.append(
                        {
                            "candidate_key": row["candidate_key"],
                            "product_link": product_link,
                            "probe_status": "blocked",
                            "body_head": body[:500],
                        }
                    )
                    break

                for _ in range(7):
                    page.mouse.wheel(0, 1800)
                    page.wait_for_timeout(1000)

                body = page.locator("body").inner_text()
                html = page.content()
                image_urls, video_urls = extract_material_urls(html)
                parsed = parse_detail_page_text(body)
                parsed.update(
                    {
                        "candidate_key": row["candidate_key"],
                        "product_link": product_link,
                        "probe_status": "ok",
                        "body_head": body[:2000],
                        "素材图片数": len(image_urls),
                        "素材视频数": len(video_urls),
                        "素材图片链接_json": json.dumps(image_urls, ensure_ascii=False),
                        "素材视频链接_json": json.dumps(video_urls, ensure_ascii=False),
                    }
                )
                probed_rows.append(parsed)
                page.wait_for_timeout(5000)
        finally:
            page.close()

    return pd.DataFrame(probed_rows)


def compute_scores(candidates: pd.DataFrame, detail_df: pd.DataFrame) -> pd.DataFrame:
    preferred = ["白", "白色", "纯棉", "重磅", "厚实", "不透", "圆领", "短袖", "纯色"]
    avoided = ["印花", "字母", "长袖", "德绒", "卫衣", "立领", "中式", "高街", "情侣装"]

    df = candidates.merge(detail_df, on=["candidate_key", "product_link"], how="left")
    df["alpha_rank_num"] = pd.to_numeric(df["alpha_rank"], errors="coerce").fillna(99)
    df["sales_num"] = df["sales_signal"].apply(sales_to_num)
    df["service_num"] = df["service_score"].apply(number_from_text)
    df["response_num"] = df["response_rate"].apply(pct_to_num)
    df["repeat_num"] = df["repeat_rate"].apply(pct_to_num)
    df["refund_num"] = df["refund_rate"].apply(pct_to_num)
    df["moq_num"] = df["moq"].apply(moq_to_num)
    df["price_num"] = pd.to_numeric(df["price"], errors="coerce").fillna(0.0)
    df["评价条数_num"] = df["评价条数"].apply(lambda x: sales_to_num(str(x).replace("条评价", "")))
    df["好评率_num"] = df["好评率"].apply(pct_to_num)
    df["店铺好评率_num"] = df["店铺好评率"].apply(pct_to_num)
    df["店铺回头率_num"] = df["店铺回头率"].apply(pct_to_num)
    df["店铺服务分_num"] = df["店铺服务分"].apply(number_from_text)

    df["标题匹配分"] = df["title"].apply(lambda title: match_score(title, preferred, avoided))
    df["AlphaShop分"] = (
        (21 - df["alpha_rank_num"].clip(lower=1, upper=20)) * 2
        + df["service_num"] * 6
        + df["response_num"] / 8
        + df["repeat_num"] / 10
        - df["refund_num"] * 4
        + df["sales_num"].apply(lambda value: min(12.0, math.log10(value + 1) * 3.5))
    )

    df["详情质量分"] = (
        df["店铺服务分_num"] * 4
        + df["店铺好评率_num"] / 8
        + df["店铺回头率_num"] / 12
        + df["好评率_num"] / 10
        + df["评价条数_num"].apply(lambda value: min(12.0, math.log10(value + 1) * 5))
    )

    df["素材完整度分"] = (
        df["素材图片数"].fillna(0).apply(lambda value: min(20.0, value * 0.35))
        + df["素材视频数"].fillna(0) * 8
        + df["贴牌换标"].eq("是") * 6
        + df["个性定制"].eq("是") * 5
        + df["有无质检报告"].eq("是") * 4
        + df["是否跨境专供"].eq("是") * 4
    )

    def price_score(price: float) -> float:
        if price <= 0:
            return 0.0
        if 10 <= price <= 35:
            return 12.0
        if 8 <= price < 10 or 35 < price <= 45:
            return 8.0
        return 3.0

    df["价格分"] = df["price_num"].apply(price_score)
    df["MOQ分"] = df["moq_num"].apply(lambda value: 6.0 if 0 < value <= 2 else 4.0 if value <= 5 else 1.0)
    df["最终总分"] = (
        df["标题匹配分"]
        + df["AlphaShop分"]
        + df["详情质量分"]
        + df["素材完整度分"]
        + df["价格分"]
        + df["MOQ分"]
    ).round(2)
    return df.sort_values(["最终总分", "素材完整度分", "AlphaShop分"], ascending=[False, False, False]).reset_index(drop=True)


def normalize_file_ext(url: str) -> str:
    lower = url.lower()
    for ext in [".jpg", ".jpeg", ".png", ".webp", ".mp4"]:
        if ext in lower:
            return ext.replace(".jpeg", ".jpg")
    return ".bin"


def download_asset(url: str, referer: str, out_dir: Path, prefix: str) -> str:
    ext = normalize_file_ext(url)
    digest = hashlib.md5(url.encode("utf-8")).hexdigest()[:12]
    out_path = out_dir / f"{prefix}_{digest}{ext}"
    if out_path.exists():
        return str(out_path)
    headers = dict(REQUEST_HEADERS)
    headers["Referer"] = referer
    response = requests.get(url, headers=headers, timeout=60)
    response.raise_for_status()
    out_path.write_bytes(response.content)
    return str(out_path)


def create_contact_sheet(image_paths: list[str], out_path: Path, title: str) -> None:
    valid = [Path(path) for path in image_paths if path and Path(path).exists()]
    if not valid:
        return
    thumbs: list[Image.Image] = []
    for index, path in enumerate(valid[:12], start=1):
        image = Image.open(path).convert("RGB")
        image = ImageOps.contain(image, (220, 220))
        canvas = Image.new("RGB", (240, 260), "white")
        canvas.paste(image, ((240 - image.width) // 2, 10))
        draw = ImageDraw.Draw(canvas)
        draw.text((8, 232), f"{index}", fill="black")
        thumbs.append(canvas)

    cols = 3
    rows = math.ceil(len(thumbs) / cols)
    sheet = Image.new("RGB", (cols * 240, rows * 260 + 40), "#f2f2f2")
    draw = ImageDraw.Draw(sheet)
    draw.text((12, 10), title, fill="black")
    for index, image in enumerate(thumbs):
        x = (index % cols) * 240
        y = 40 + (index // cols) * 260
        sheet.paste(image, (x, y))
    sheet.save(out_path, quality=92)


def download_source_materials(source_rows: pd.DataFrame) -> pd.DataFrame:
    asset_rows: list[dict[str, Any]] = []
    for _, row in source_rows.iterrows():
        source_folder = DOWNLOAD_DIR / row["source_role"] / row["candidate_key"].replace("https://detail.1688.com/offer/", "").replace(".html", "").replace("/", "_")
        source_folder.mkdir(parents=True, exist_ok=True)
        image_urls = json.loads(row["素材图片链接_json"]) if row.get("素材图片链接_json") else []
        video_urls = json.loads(row["素材视频链接_json"]) if row.get("素材视频链接_json") else []
        gallery_paths: list[str] = []
        detail_paths: list[str] = []
        for idx, url in enumerate(image_urls, start=1):
            asset_type = "详情图" if "_sum.jpg" in url else "主图/细节图"
            prefix = f"{asset_type}_{idx:02d}".replace("/", "_")
            saved = download_asset(url, str(row["product_link"]), source_folder, prefix)
            asset_rows.append(
                {
                    "source_role": row["source_role"],
                    "candidate_key": row["candidate_key"],
                    "supplier": row["supplier"],
                    "product_link": row["product_link"],
                    "asset_type": asset_type,
                    "asset_url": url,
                    "local_path": saved,
                }
            )
            if asset_type == "详情图":
                detail_paths.append(saved)
            else:
                gallery_paths.append(saved)
        for idx, url in enumerate(video_urls, start=1):
            saved = download_asset(url, str(row["product_link"]), source_folder, f"视频_{idx:02d}")
            asset_rows.append(
                {
                    "source_role": row["source_role"],
                    "candidate_key": row["candidate_key"],
                    "supplier": row["supplier"],
                    "product_link": row["product_link"],
                    "asset_type": "视频",
                    "asset_url": url,
                    "local_path": saved,
                }
            )

        create_contact_sheet(
            gallery_paths,
            source_folder / "主图联系表.jpg",
            f"{row['source_role']} | {row['supplier']} | 主图/细节图",
        )
        create_contact_sheet(
            detail_paths,
            source_folder / "详情图联系表.jpg",
            f"{row['source_role']} | {row['supplier']} | 详情图",
        )
    return pd.DataFrame(asset_rows)


def reference_page_summary() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"模块": "首图轮播", "说明": "参考 Oakln 页面，主图轮播约 16 张，包含正面图、细节图、功能图、颜色图。"},
            {"模块": "标题区", "说明": "长标题 + 售价 + 颜色 + 尺码。"},
            {"模块": "卖点文案", "说明": "页面 Meta 里直接写了产地、材质、尺码、季节、洗涤等基础信息。"},
            {"模块": "详情承接", "说明": "详情区延续同一商品素材，偏电商长图结构。"},
            {"模块": "需要准备的素材", "说明": "主图、颜色图、细节图、尺寸/参数图、视频、标题、副标题、参数表。"},
        ]
    )


def build_final_selection(df: pd.DataFrame) -> pd.DataFrame:
    main_supply = df.iloc[0].copy()
    material_only = df.sort_values(["素材完整度分", "最终总分"], ascending=[False, False]).copy()
    material_only = material_only[material_only["candidate_key"] != main_supply["candidate_key"]].head(2)
    rows = [main_supply.to_dict()]
    for _, row in material_only.iterrows():
        rows.append(row.to_dict())
    final = pd.DataFrame(rows)
    final.insert(0, "source_role", ["最终供货主选", "素材补充源1", "素材补充源2"][: len(final)])
    return final


def build_copy(final_rows: pd.DataFrame) -> str:
    main = final_rows.iloc[0]
    weight = str(main.get("克重", "")).strip() or "280g级"
    sizes = str(main.get("尺码", "")).replace(",", " / ")
    colors = str(main.get("颜色", "")).replace(",", " / ")
    title_ja = f"透けにくい ヘビーウェイト 白T メンズ {weight} コットン クルーネック 半袖"
    subtitle_ja = "厚手で1枚でも着やすく、インナー使いもしやすい日本向け定番白T"
    lines = [
        f"# T01 商品详情页素材包",
        "",
        f"- 供货主选：{main['supplier']}",
        f"- 主商品链接：{main['product_link']}",
        f"- 建议日文标题：{title_ja}",
        f"- 建议日文副标题：{subtitle_ja}",
        "",
        "## 日文卖点",
        "",
        "- 透け感を抑えた厚手コットンで、白T1枚でも着やすい",
        "- クルーネックの立ち上がりがきれいで、ジャケットのインナーにも合わせやすい",
        "- ゆるすぎないベーシック寄りのシルエットで、通勤カジュアルにも使いやすい",
        "- 春夏の単品使いだけでなく、4月のレイヤードにも対応しやすい",
        "- 無地ベースなので、ページ展開や色展開の横展開もしやすい",
        "",
        "## 参数建议",
        "",
        f"- 颜色：{colors}",
        f"- 尺码：{sizes}",
        f"- 面料名称：{main.get('面料名称', '')}",
        f"- 主面料成分：{main.get('主面料成分', '')}",
        f"- 克重：{weight}",
        f"- 版型：{main.get('版型', '')}",
        f"- 领型：{main.get('领型', '')}",
        "",
        "## 页面结构建议",
        "",
        "- 主图轮播：选 8-12 张，优先白色主图、面料细节、领口细节、厚度表现、尺码图、卖点图",
        "- 视频：保留 1 个主视频放在主图区或详情区前段",
        "- 详情长图：优先下载的 `_sum.jpg` 详情图，再补 1-2 张参数说明图",
        "- 素材补充：如果主选供应商主图不够丰富，用素材补充源的细节图和视频补齐",
        "",
    ]
    return "\n".join(lines)


def write_excel(
    sessions_df: pd.DataFrame,
    all_candidates_df: pd.DataFrame,
    detail_df: pd.DataFrame,
    scored_df: pd.DataFrame,
    final_df: pd.DataFrame,
    assets_df: pd.DataFrame,
) -> None:
    with pd.ExcelWriter(REPORT_XLSX, engine="xlsxwriter") as writer:
        reference_page_summary().to_excel(writer, sheet_name="参考页结构", index=False)
        sessions_df.to_excel(writer, sheet_name="搜索记录", index=False)
        all_candidates_df.to_excel(writer, sheet_name="全量候选", index=False)
        detail_df.to_excel(writer, sheet_name="详情探测", index=False)
        scored_df.to_excel(writer, sheet_name="综合评分", index=False)
        final_df.to_excel(writer, sheet_name="最终供应商与素材源", index=False)
        assets_df.to_excel(writer, sheet_name="素材清单", index=False)
        for worksheet in writer.sheets.values():
            worksheet.freeze_panes(1, 0)


def main() -> None:
    selected = latest_selected_row()
    raw_step4 = pd.read_csv(RAW_STEP4_CSV)
    t01_history = raw_step4[raw_step4["sku_id"] == "T01"].copy()
    t01_history["search_id"] = "step4_history"
    t01_history["search_type"] = "历史文字找商"
    t01_history["prompt"] = "来自第4步 AlphaShop 文字找商历史候选"
    t01_history["search_session_url"] = ""
    t01_history["image_search_path"] = ""

    main_image_path = Path(str(selected["主采图片文件"]))
    backup_image_path = Path(str(selected["备采图片文件"]))

    raw_621 = t01_history[t01_history["product_link"].astype(str).str.contains("621147723281", na=False)].head(1)
    third_image_path = None
    if not raw_621.empty:
        third_image_path = prepare_local_image(
            str(raw_621.iloc[0]["image_url"]),
            "t01_621_image.png",
        )

    search_tasks = [
        SearchTask(
            search_id="text_main",
            search_type="文字找商",
            prompt=(
                "Find 1688 suppliers for men's heavyweight opaque plain white t-shirt for Japan April sales. "
                "Need round neck, no print, 250g-300g cotton, clean silhouette, office casual friendly, small MOQ, strong factory."
            ),
        ),
        SearchTask(
            search_id="text_alt",
            search_type="文字找商",
            prompt=(
                "Find 1688 suppliers for men's premium white tee. Prefer 280g or 300g cotton, not see-through, stable collar, "
                "supports relabeling, packaging customization, small MOQ, and export friendly."
            ),
        ),
        SearchTask(
            search_id="image_main",
            search_type="图片+文字找商",
            prompt=(
                "Find 1688 suppliers for this exact men's heavyweight plain white t-shirt image. "
                "Prefer opaque, clean round neck, no print, 250g-300g cotton, small MOQ, strong factory."
            ),
            image_path=main_image_path,
        ),
        SearchTask(
            search_id="image_backup",
            search_type="图片+文字找商",
            prompt=(
                "Find 1688 suppliers for this plain white short-sleeve t-shirt image. "
                "Prefer thick cotton, no logo, good collar shape, Japan-friendly basic style, and complete product materials."
            ),
            image_path=backup_image_path,
        ),
    ]
    if third_image_path:
        search_tasks.append(
            SearchTask(
                search_id="image_621",
                search_type="图片+文字找商",
                prompt=(
                    "Find 1688 suppliers for this plain white relaxed fit t-shirt image. "
                    "Prefer complete gallery images, detail long images, video, quality report, relabel support, and export support."
                ),
                image_path=third_image_path,
            )
        )

    search_df, sessions_df = run_alphashop_searches(search_tasks)
    combined = pd.concat([t01_history, search_df], ignore_index=True, sort=False)
    combined["candidate_key"] = combined.apply(build_candidate_key, axis=1)

    detail_df = probe_detail_pages(combined[["candidate_key", "product_link"]].drop_duplicates().merge(
        combined[["candidate_key", "supplier"]].drop_duplicates(), on="candidate_key", how="left"
    ))

    scored_df = compute_scores(combined, detail_df)
    final_df = build_final_selection(scored_df)
    assets_df = download_source_materials(final_df)

    REPORT_JSON.write_text(
        json.dumps(
            {
                "selected_sku": selected.to_dict(),
                "search_sessions": sessions_df.to_dict(orient="records"),
                "final_selection": final_df.to_dict(orient="records"),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    REPORT_MD.write_text(build_copy(final_df), encoding="utf-8")
    write_excel(sessions_df, combined, detail_df, scored_df, final_df, assets_df)
    print(REPORT_XLSX)
    print(REPORT_MD)


if __name__ == "__main__":
    main()
