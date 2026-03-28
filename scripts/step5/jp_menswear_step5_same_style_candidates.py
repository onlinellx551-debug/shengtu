from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd
import requests
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
STEP4_DIR = ROOT / "step4_output"
STEP5_DIR = ROOT / "step5_output"
STEP5_DIR.mkdir(exist_ok=True)
RAW_STEP4_CSV = STEP4_DIR / "alphashop_supplier_raw_2026-03-20.csv"
SELECTED_XLS = next(path for path in STEP5_DIR.glob("*.xls") if not path.name.startswith("~$"))
OUT_XLSX = STEP5_DIR / "T01_同款候选确认版.xlsx"
OUT_CSV = STEP5_DIR / "T01_同款候选确认版.csv"
IMG_DIR = STEP5_DIR / "same_style_preview"
IMG_DIR.mkdir(exist_ok=True)

CDP_URL = "http://127.0.0.1:9222"
HOME_URL = "https://create.alphashop.cn/"


@dataclass(frozen=True)
class SearchTask:
    search_id: str
    prompt: str
    image_path: Path


def clean_lines(text: str) -> list[str]:
    return [line.strip() for line in re.split(r"[\r\n]+", str(text)) if line.strip()]


def label_value(lines: list[str], label: str) -> str:
    for index, line in enumerate(lines):
        if line == label and index + 1 < len(lines):
            return lines[index + 1]
        if line.startswith(f"{label}:") or line.startswith(f"{label}："):
            return re.split(r"[:：]", line, maxsplit=1)[1].strip()
    return ""


def parse_search_row(row_text: str) -> dict[str, Any]:
    lines = clean_lines(row_text)
    rank = lines[0] if lines and lines[0].isdigit() else ""
    title = lines[1] if len(lines) > 1 else ""
    price_match = re.search(r"¥\s*([0-9.]+)", row_text)
    price = price_match.group(1) if price_match else ""

    supplier = ""
    supplier_tags: list[str] = []
    if title:
        title_index = lines.index(title)
        for line in lines[title_index + 1 :]:
            if line.startswith("¥"):
                continue
            if line.startswith(("✨", "❗")):
                break
            supplier = line
            break
        if supplier:
            supplier_index = lines.index(supplier)
            for line in lines[supplier_index + 1 :]:
                if line.startswith(("✨", "❗")):
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
        "response_rate": label_value(lines, "客服响应率"),
        "repeat_rate": label_value(lines, "90天回头率"),
        "refund_rate": label_value(lines, "品质退款率"),
        "raw_text": row_text,
    }


def open_findshop(page) -> None:
    page.goto(HOME_URL, wait_until="domcontentloaded", timeout=90000)
    page.wait_for_timeout(7000)
    page.mouse.click(620, 257)
    page.wait_for_timeout(2500)


def fill_prompt(page, prompt: str) -> None:
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


def click_send(page) -> None:
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
    raise RuntimeError("找不到发送按钮")


def submit_search(page, task: SearchTask) -> str:
    open_findshop(page)
    fill_prompt(page, task.prompt)
    page.locator('input[type="file"]').set_input_files(str(task.image_path))
    page.wait_for_timeout(3500)
    click_send(page)
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
    raise TimeoutError(f"{task.search_id} 没有正常返回结果")


def extract_product_link(page, row) -> str:
    image = row.locator("img").first
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


def run_searches(tasks: list[SearchTask]) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(CDP_URL)
        context = browser.contexts[0]
        page = context.new_page()
        try:
            for task in tasks:
                submit_search(page, task)
                table_rows = page.locator("tr")
                for idx in range(2, table_rows.count()):
                    row = table_rows.nth(idx)
                    row_text = row.inner_text().strip()
                    if not row_text:
                        continue
                    parsed = parse_search_row(row_text)
                    parsed.update(
                        {
                            "search_id": task.search_id,
                            "prompt": task.prompt,
                            "search_image": str(task.image_path),
                            "search_session_url": page.url,
                            "image_url": row.locator("img").first.get_attribute("src") or "",
                            "product_link": extract_product_link(page, row),
                        }
                    )
                    rows.append(parsed)
                page.wait_for_timeout(5000)
        finally:
            page.close()
    return pd.DataFrame(rows)


def same_style_score(title: str) -> tuple[int, str]:
    text = str(title or "")
    score = 0
    reasons: list[str] = []
    positives = ["白", "白色", "纯色", "短袖", "T恤", "圆领", "纯棉", "重磅", "厚实", "打底", "基础"]
    negatives = ["长袖", "印花", "字母", "文化衫", "班服", "卫衣", "立领", "中式", "裤", "裙", "外套"]
    for word in positives:
        if word in text:
            score += 2
            reasons.append(f"+{word}")
    for word in negatives:
        if word in text:
            score -= 3
            reasons.append(f"-{word}")
    return score, " ".join(reasons)


def ensure_local_preview(url: str, row_id: str) -> str:
    if not str(url).startswith("http"):
        return ""
    out_path = IMG_DIR / f"{row_id}.jpg"
    if out_path.exists():
        return str(out_path)
    response = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=30)
    response.raise_for_status()
    out_path.write_bytes(response.content)
    return str(out_path)


def apply_hyperlinks(workbook, worksheet, df: pd.DataFrame) -> None:
    link_format = workbook.add_format({"font_color": "blue", "underline": 1})
    for column in ["商品链接", "图片链接", "搜索会话链接"]:
        if column not in df.columns:
            continue
        col = df.columns.get_loc(column)
        for row_idx, value in enumerate(df[column], start=1):
            text = str(value or "").strip()
            if text.startswith("http"):
                worksheet.write_url(row_idx, col, text, link_format, string=text)


def write_image_sheet(writer: pd.ExcelWriter, df: pd.DataFrame) -> None:
    export_df = df.copy()
    export_df.insert(0, "图片预览", "")
    export_df.to_excel(writer, sheet_name="同款候选", index=False)
    workbook = writer.book
    worksheet = writer.sheets["同款候选"]
    worksheet.freeze_panes(1, 1)
    worksheet.set_column(0, 0, 14)
    width_map = {
        "标题": 40,
        "供应商": 24,
        "同款初筛说明": 24,
        "搜索来源": 16,
        "商品链接": 18,
        "图片链接": 18,
        "搜索会话链接": 18,
    }
    for idx, column in enumerate(export_df.columns):
        worksheet.set_column(idx, idx, width_map.get(column, 14))
    apply_hyperlinks(workbook, worksheet, export_df)

    preview_col = 0
    helper_col = export_df.columns.get_loc("图片文件")
    worksheet.set_column(helper_col, helper_col, 2, None, {"hidden": True})
    for row_idx, path in enumerate(export_df["图片文件"], start=1):
        worksheet.set_row(row_idx, 78)
        if path and Path(path).exists():
            worksheet.insert_image(row_idx, preview_col, path, {"x_scale": 0.52, "y_scale": 0.52, "x_offset": 4, "y_offset": 4})
        else:
            worksheet.write(row_idx, preview_col, "无图")
        worksheet.write(row_idx, helper_col, "")


def main() -> None:
    selected = pd.read_excel(SELECTED_XLS, sheet_name=0).iloc[0]
    image_paths = [
        Path(r"C:\Users\Administrator\Desktop\Codex Project\task-codex\step5_output\exact_images\621_4.webp"),
        Path(r"C:\Users\Administrator\Desktop\Codex Project\task-codex\step5_output\exact_images\621_3.webp"),
        Path(r"C:\Users\Administrator\Desktop\Codex Project\task-codex\step5_output\exact_images\692_1.jpg"),
    ]

    tasks = [
        SearchTask(
            search_id=f"exact_img_{idx+1}",
            prompt="帮我找这件白色短袖T恤的同款或高度接近同版型同面料的1688供应商。只要白色、纯色、短袖、圆领、厚实/重磅、基础款，排除长袖、卫衣、印花、文化衫、班服、非T恤。",
            image_path=path,
        )
        for idx, path in enumerate(image_paths)
        if path.exists()
    ]

    search_df = run_searches(tasks)
    step4_df = pd.read_csv(RAW_STEP4_CSV)
    history_df = step4_df[step4_df["sku_id"] == "T01"].copy()
    history_df["search_id"] = "step4_history"
    history_df["prompt"] = "第4步历史 T01 候选"
    history_df["search_image"] = ""
    history_df["search_session_url"] = ""

    combined = pd.concat([history_df, search_df], ignore_index=True, sort=False)
    combined["candidate_key"] = combined.apply(
        lambda row: str(row.get("product_link", "")).strip() if "detail.1688.com/offer/" in str(row.get("product_link", "")) else f"{row.get('supplier','')}||{row.get('title','')}",
        axis=1,
    )
    combined["同款初筛分"], combined["同款初筛说明"] = zip(*combined["title"].apply(same_style_score))
    filtered = (
        combined.sort_values(["同款初筛分"], ascending=[False])
        .drop_duplicates(subset=["candidate_key"], keep="first")
        .query("同款初筛分 >= 4")
        .copy()
    )

    filtered["图片文件"] = [
        ensure_local_preview(url, f"cand_{idx+1:03d}") for idx, url in enumerate(filtered["image_url"])
    ]

    export_df = filtered[
        [
            "search_id",
            "alpha_rank",
            "title",
            "supplier",
            "price",
            "sales_signal",
            "moq",
            "origin",
            "highlights",
            "同款初筛分",
            "同款初筛说明",
            "image_url",
            "product_link",
            "search_session_url",
            "图片文件",
        ]
    ].rename(
        columns={
            "search_id": "搜索来源",
            "alpha_rank": "AlphaShop排名",
            "title": "标题",
            "supplier": "供应商",
            "price": "价格",
            "sales_signal": "销量信号",
            "moq": "起批量",
            "origin": "发货地",
            "highlights": "AI亮点",
            "image_url": "图片链接",
            "product_link": "商品链接",
            "search_session_url": "搜索会话链接",
        }
    )

    summary_df = pd.DataFrame(
        [
            {"项目": "SKU", "值": selected["SKU编号"]},
            {"项目": "产品名称", "值": selected["产品名称"]},
            {"项目": "当前目标", "值": "先确认同款候选，再进入素材阶段"},
            {"项目": "候选数量", "值": len(export_df)},
            {"项目": "说明", "值": "这版只做同款初筛，保留图片让你人工确认是否真同款。"},
        ]
    )

    with pd.ExcelWriter(OUT_XLSX, engine="xlsxwriter") as writer:
        summary_df.to_excel(writer, sheet_name="概览", index=False)
        write_image_sheet(writer, export_df)
        combined.to_excel(writer, sheet_name="原始合并结果", index=False)

    export_df.to_csv(OUT_CSV, index=False, encoding="utf-8-sig")
    print(OUT_XLSX)
    print(OUT_CSV)


if __name__ == "__main__":
    main()
