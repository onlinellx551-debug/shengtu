from __future__ import annotations

import random
import re
import time
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

import pandas as pd
import requests
from PIL import Image
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


EXPORT_DATE = "2026-03-20"
CDP_URL = "http://127.0.0.1:9222"
ROOT = Path(__file__).resolve().parent
STEP3_DIR = ROOT / "step3_output"
OUT_DIR = ROOT / "step4_output"
IMG_DIR = OUT_DIR / "alphashop_images"
OUT_DIR.mkdir(exist_ok=True)
IMG_DIR.mkdir(exist_ok=True)

OUTPUT_XLSX = OUT_DIR / f"日本男装第四步_AlphaShop找商版_{EXPORT_DATE}.xlsx"
OUTPUT_MD = OUT_DIR / "第四步_AlphaShop找商版结论.md"
OUTPUT_RAW = OUT_DIR / f"alphashop_supplier_raw_{EXPORT_DATE}.csv"

HOME_URL = "https://create.alphashop.cn/"
REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/146.0.0.0 Safari/537.36"
    ),
    "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
}


@dataclass(frozen=True)
class SkuPrompt:
    sku_id: str
    product_name: str
    priority: str
    style_line: str
    scenario: str
    prompt_en: str
    preferred_words: tuple[str, ...]
    negative_words: tuple[str, ...]


SKU_PROMPTS: list[SkuPrompt] = [
    SkuPrompt(
        sku_id="O01",
        product_name="可机洗弹力海军蓝单西 / 套装上衣",
        priority="A",
        style_line="办公室休闲",
        scenario="通勤 / 轻正式",
        prompt_en=(
            "Find 1688 suppliers for mens navy lightweight blazer or suit jacket for Japan April sales. "
            "Requirements: office casual, washable or easy care, slight stretch, clean structure, not stiff, "
            "not wedding formal, supports small MOQ and sampling. Prefer factories or strong merchants."
        ),
        preferred_words=("藏青", "海军蓝", "轻薄", "弹力", "免烫", "商务休闲", "单西", "西装外套"),
        negative_words=("新郎", "结婚", "礼服", "oversize", "短袖", "垫肩", "主持", "影楼"),
    ),
    SkuPrompt(
        sku_id="S01",
        product_name="免烫常规领白衬衫",
        priority="A",
        style_line="办公室休闲",
        scenario="通勤 / 面试 / 入职",
        prompt_en=(
            "Find 1688 suppliers for mens non-iron white long-sleeve shirt for Japan April sales. "
            "Requirements: low sheen, not see-through, office casual, regular collar, easy care, "
            "supports small MOQ and sampling. Prefer factories or strong merchants."
        ),
        preferred_words=("免烫", "抗皱", "白色", "长袖", "商务", "通勤", "DP", "纯棉"),
        negative_words=("短袖", "半袖", "黑色", "修身", "新郎", "结婚", "工装", "制服"),
    ),
    SkuPrompt(
        sku_id="P01",
        product_name="同面料轻弹套装裤",
        priority="A",
        style_line="办公室休闲",
        scenario="通勤 / 套装搭配",
        prompt_en=(
            "Find 1688 suppliers for mens lightweight suit trousers for Japan April sales. "
            "Requirements: office casual, straight fit, slight stretch, wrinkle resistant, suitable to match a navy blazer, "
            "not slim, not cropped, supports small MOQ and sampling. Prefer factories or strong merchants."
        ),
        preferred_words=("西裤", "西装裤", "弹力", "抗皱", "垂感", "直筒", "通勤"),
        negative_words=("加绒", "加厚", "九分", "修身", "小脚", "执勤", "制服", "冰丝", "速干"),
    ),
    SkuPrompt(
        sku_id="P02",
        product_name="一褶宽直筒西裤",
        priority="A",
        style_line="办公室休闲",
        scenario="通勤 / 日常",
        prompt_en=(
            "Find 1688 suppliers for mens one-pleat clean straight slacks for Japan April sales. "
            "Requirements: office casual, relaxed straight fit, clean drape, not too wide, not slim, "
            "not cropped, supports small MOQ and sampling. Prefer factories or strong merchants."
        ),
        preferred_words=("一褶", "直筒", "宽松", "垂感", "西裤", "西装裤"),
        negative_words=("阔腿", "拖地", "修身", "九分", "韩版", "喇叭"),
    ),
    SkuPrompt(
        sku_id="D01",
        product_name="原色直筒牛仔裤",
        priority="A",
        style_line="轻美式",
        scenario="日常 / 通勤",
        prompt_en=(
            "Find 1688 suppliers for mens raw or dark indigo straight jeans for Japan April sales. "
            "Requirements: clean wash, regular straight fit, not ripped, not cargo, not overly wide, "
            "supports small MOQ and sampling. Prefer factories or strong merchants."
        ),
        preferred_words=("原色", "直筒", "牛仔裤", "深蓝", "赤耳", "丹宁"),
        negative_words=("破洞", "做旧", "工装", "喇叭", "拖地", "阔腿"),
    ),
    SkuPrompt(
        sku_id="S02",
        product_name="蓝白条牛津扣领衬衫",
        priority="A",
        style_line="轻学院",
        scenario="通勤 / 日常",
        prompt_en=(
            "Find 1688 suppliers for mens blue white striped oxford button-down shirt for Japan April sales. "
            "Requirements: office casual, light preppy feel, long sleeve, clean collar, not uniform-like, "
            "supports small MOQ and sampling. Prefer factories or strong merchants."
        ),
        preferred_words=("条纹", "蓝白", "牛津纺", "扣领", "长袖", "衬衫"),
        negative_words=("短袖", "校服", "制服", "修身", "韩版"),
    ),
    SkuPrompt(
        sku_id="K01",
        product_name="V领轻学院开衫",
        priority="B",
        style_line="轻学院",
        scenario="叠穿 / 日常",
        prompt_en=(
            "Find 1688 suppliers for mens light V-neck cardigan for Japan April sales. "
            "Requirements: plain color, lightweight knit, easy layering, mild preppy feel, not loud logo, "
            "supports small MOQ and sampling. Prefer factories or strong merchants."
        ),
        preferred_words=("V领", "开衫", "针织", "纯色", "薄款"),
        negative_words=("条纹", "校服", "刺绣", "logo", "背心", "童装"),
    ),
    SkuPrompt(
        sku_id="D02",
        product_name="黑色轻宽直筒牛仔裤",
        priority="B",
        style_line="轻街头",
        scenario="日常 / 通勤",
        prompt_en=(
            "Find 1688 suppliers for mens black straight jeans for Japan April sales. "
            "Requirements: clean silhouette, slightly loose straight fit, not ripped, not cargo, not low rise, "
            "supports small MOQ and sampling. Prefer factories or strong merchants."
        ),
        preferred_words=("黑色", "直筒", "牛仔裤", "宽松"),
        negative_words=("破洞", "工装", "低腰", "拖地", "阔腿"),
    ),
    SkuPrompt(
        sku_id="T01",
        product_name="厚实不透白T",
        priority="B",
        style_line="基础通勤",
        scenario="打底 / 单穿",
        prompt_en=(
            "Find 1688 suppliers for mens thick white t-shirt for Japan April sales. "
            "Requirements: not see-through, clean basic style, no print, good collar shape, office casual friendly, "
            "supports small MOQ and sampling. Prefer factories or strong merchants."
        ),
        preferred_words=("白色", "T恤", "重磅", "纯棉", "不透"),
        negative_words=("印花", "图案", "字母", "背心", "修身"),
    ),
    SkuPrompt(
        sku_id="J01",
        product_name="短款牛仔夹克",
        priority="C",
        style_line="轻美式",
        scenario="春季外套",
        prompt_en=(
            "Find 1688 suppliers for mens cropped or short denim jacket for Japan April sales. "
            "Requirements: clean wash, light American casual feel, not heavily distressed, not oversized, "
            "supports small MOQ and sampling. Prefer factories or strong merchants."
        ),
        preferred_words=("短款", "短版", "牛仔夹克", "牛仔外套", "春季"),
        negative_words=("破洞", "重做旧", "oversize", "印花", "拼接"),
    ),
]


def latest_workbook(directory: Path) -> Path:
    files = sorted(p for p in directory.glob("*.xlsx") if not p.name.startswith("~$"))
    if not files:
        raise FileNotFoundError(f"Missing xlsx file in {directory}")
    return files[-1]


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


def click_supplier_tab(page) -> None:
    page.goto(HOME_URL, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(6000)
    ok = page.evaluate(
        """
        () => {
          const items = Array.from(document.querySelectorAll('[class*="navItem"]'));
          const el = items.find(node => (node.innerText || '').includes('找商'));
          if (el) {
            el.click();
            return true;
          }
          return false;
        }
        """
    )
    if not ok:
        page.mouse.click(500, 255)
    page.wait_for_timeout(2500)


def submit_supplier_prompt(page, prompt: str) -> str:
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

    send_button = None
    buttons = page.locator("button")
    for i in range(buttons.count()):
        box = buttons.nth(i).bounding_box()
        if not box:
            continue
        if box["width"] <= 40 and box["height"] <= 40 and box["y"] < 500:
            send_button = buttons.nth(i)
            break
    if send_button is None:
        raise RuntimeError("Unable to find AlphaShop send button.")
    send_button.click()

    page.wait_for_url("**/sourcing**", timeout=60000)
    wait_for_results(page)
    return page.url


def wait_for_results(page) -> None:
    for _ in range(45):
        body_text = page.evaluate("() => document.body.innerText")
        if "供应商推荐" in body_text and page.locator("tr").count() >= 4:
            return
        page.wait_for_timeout(2000)
    raise TimeoutError("AlphaShop supplier results did not finish loading in time.")


def find_after(label: str, text: str) -> str:
    match = re.search(rf"{re.escape(label)}\s*[:：]\s*([^\n]+)", text)
    return match.group(1).strip() if match else ""


def find_metric(label: str, text: str) -> str:
    match = re.search(rf"{re.escape(label)}\s*([0-9.]+%?)", text)
    return match.group(1).strip() if match else ""


def parse_row_text(row_text: str) -> dict[str, str]:
    lines = [line.strip() for line in row_text.splitlines() if line.strip()]
    price = ""
    price_match = re.search(r"¥\s*([0-9.]+)", row_text)
    if price_match:
        price = price_match.group(1)

    rank = lines[0] if lines and lines[0].isdigit() else ""
    title = lines[1] if len(lines) > 1 else ""
    supplier = ""
    if price and f"¥{price}" in lines:
        price_idx = lines.index(f"¥{price}")
        supplier = lines[price_idx + 1] if len(lines) > price_idx + 1 else ""
    elif len(lines) > 2:
        supplier = lines[2]

    highlights = [line for line in lines if line.startswith("✨")][:2]
    supplier_tags: list[str] = []
    if supplier:
        supplier_idx = lines.index(supplier)
        for line in lines[supplier_idx + 1 :]:
            if line.startswith("✨") or line.startswith("男士") or line.startswith("近一年"):
                break
            supplier_tags.append(line)

    return {
        "alpha_rank": rank,
        "title": title,
        "price": price,
        "supplier": supplier,
        "supplier_tags": " / ".join(supplier_tags[:4]),
        "highlights": " / ".join(highlights),
        "sales_signal": find_after("近一年全网销量", row_text) or find_after("近一年销量", row_text),
        "moq": find_after("起批量", row_text),
        "origin": find_after("发货地", row_text),
        "service_score": find_metric("综合服务分", row_text),
        "response_rate": find_metric("客服响应率", row_text),
        "repeat_rate": find_metric("90天回头率", row_text),
        "refund_rate": find_metric("品质退款率", row_text),
        "fulfillment_rate": find_after("发货履约率", row_text),
        "pickup_rate": find_after("48小时揽收率", row_text),
        "raw_text": row_text,
    }


def to_float(value: str) -> float:
    try:
        return float(value)
    except Exception:
        return 0.0


def local_fit_score(candidate: dict[str, str], sku: SkuPrompt) -> float:
    title = candidate["title"]
    score = 100.0
    rank = int(candidate["alpha_rank"]) if candidate["alpha_rank"].isdigit() else 99
    score -= (rank - 1) * 8
    score += to_float(candidate["service_score"]) * 2.5
    response = candidate["response_rate"].replace("%", "")
    repeat = candidate["repeat_rate"].replace("%", "")
    score += to_float(response) / 10
    score += to_float(repeat) / 12
    if "源头工厂" in candidate["supplier_tags"]:
        score += 6
    if "实力商家" in candidate["supplier_tags"]:
        score += 4
    if "超级工厂" in candidate["supplier_tags"]:
        score += 4
    if candidate["moq"].startswith("1件") or candidate["moq"].startswith("2件"):
        score += 4
    for word in sku.preferred_words:
        if word in title:
            score += 6
    for word in sku.negative_words:
        if word.lower() in title.lower():
            score -= 28
    return round(score, 2)


def extract_product_link(page, row) -> str:
    image = row.locator("img").first
    if image.count() == 0:
        return ""
    try:
        with page.expect_popup(timeout=10000) as popup_info:
            image.click()
        popup = popup_info.value
        popup.wait_for_timeout(2000)
        url = popup.url
        popup.close()
        page.wait_for_timeout(1000)
        return url
    except PlaywrightTimeoutError:
        return ""


def query_one_sku(ctx, sku: SkuPrompt, max_rows: int | None = None) -> list[dict[str, object]]:
    page = ctx.new_page()
    try:
        click_supplier_tab(page)
        session_url = submit_supplier_prompt(page, sku.prompt_en)
        rows = page.locator("tr")
        collected: list[dict[str, object]] = []
        stop_at = rows.count() if max_rows is None else min(rows.count(), 2 + max_rows)
        for idx in range(2, stop_at):
            row = rows.nth(idx)
            row_text = row.inner_text()
            parsed = parse_row_text(row_text)
            image = row.locator("img").first
            image_src = image.get_attribute("src") or ""
            product_link = extract_product_link(page, row)
            parsed.update(
                {
                    "sku_id": sku.sku_id,
                    "product_name": sku.product_name,
                    "priority": sku.priority,
                    "style_line": sku.style_line,
                    "scenario": sku.scenario,
                    "prompt_en": sku.prompt_en,
                    "session_url": session_url,
                    "image_url": image_src,
                    "product_link": product_link,
                    "local_fit_score": 0.0,
                }
            )
            parsed["local_fit_score"] = local_fit_score(parsed, sku)
            collected.append(parsed)
        return collected
    finally:
        page.close()


def download_image(url: str, dest_prefix: str) -> str:
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


def choose_final_rows(raw_df: pd.DataFrame, prior_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    final_rows: list[dict[str, object]] = []
    flat_rows: list[dict[str, object]] = []
    for _, prior in prior_df.iterrows():
        sku_id = prior["SKU编号"]
        pool = raw_df.loc[raw_df["sku_id"] == sku_id].copy()
        if pool.empty:
            continue
        pool = pool.sort_values(["local_fit_score", "alpha_rank"], ascending=[False, True]).reset_index(drop=True)
        main = pool.iloc[0]
        backup = pool.iloc[1 if len(pool) > 1 else 0]
        final_rows.append(
            {
                "SKU编号": sku_id,
                "产品名称": prior["产品名称"],
                "优先级": prior["优先级"],
                "风格线": prior["风格线"],
                "场景": prior["场景"],
                "AlphaShop会话链接": main["session_url"],
                "主采标题": main["title"],
                "主采供应商": main["supplier"],
                "主采价格": main["price"],
                "主采起批量": main["moq"],
                "主采发货地": main["origin"],
                "主采销量信号": main["sales_signal"],
                "主采服务分": main["service_score"],
                "主采客服响应率": main["response_rate"],
                "主采回头率": main["repeat_rate"],
                "主采亮点": main["highlights"],
                "主采图片链接": main["image_url"],
                "主采商品链接": main["product_link"],
                "主采选择理由": (
                    f"AlphaShop 排名 {main['alpha_rank']}；"
                    f"供应商为 {main['supplier_tags'] or '常规商家'}；"
                    f"起批量 {main['moq'] or '未显示'}；"
                    f"客服响应率 {main['response_rate'] or '未显示'}；"
                    f"综合服务分 {main['service_score'] or '未显示'}。"
                ),
                "备采标题": backup["title"],
                "备采供应商": backup["supplier"],
                "备采价格": backup["price"],
                "备采起批量": backup["moq"],
                "备采发货地": backup["origin"],
                "备采销量信号": backup["sales_signal"],
                "备采服务分": backup["service_score"],
                "备采客服响应率": backup["response_rate"],
                "备采回头率": backup["repeat_rate"],
                "备采亮点": backup["highlights"],
                "备采图片链接": backup["image_url"],
                "备采商品链接": backup["product_link"],
                "备采选择理由": (
                    f"AlphaShop 排名 {backup['alpha_rank']}；"
                    f"供应商为 {backup['supplier_tags'] or '常规商家'}；"
                    f"起批量 {backup['moq'] or '未显示'}；"
                    f"客服响应率 {backup['response_rate'] or '未显示'}；"
                    f"综合服务分 {backup['service_score'] or '未显示'}。"
                ),
            }
        )
        for role, row in [("主采", main), ("备采", backup)]:
            flat_rows.append(
                {
                    "角色": role,
                    "SKU编号": sku_id,
                    "产品名称": prior["产品名称"],
                    "供应商": row["supplier"],
                    "商品标题": row["title"],
                    "价格": row["price"],
                    "起批量": row["moq"],
                    "发货地": row["origin"],
                    "销量信号": row["sales_signal"],
                    "服务分": row["service_score"],
                    "客服响应率": row["response_rate"],
                    "回头率": row["repeat_rate"],
                    "供应商标签": row["supplier_tags"],
                    "AI亮点": row["highlights"],
                    "图片链接": row["image_url"],
                    "商品链接": row["product_link"],
                    "AlphaShop会话链接": row["session_url"],
                    "本地适配分": row["local_fit_score"],
                    "原始文本": row["raw_text"],
                }
            )
    return pd.DataFrame(final_rows), pd.DataFrame(flat_rows)


def attach_images(final_df: pd.DataFrame) -> pd.DataFrame:
    df = final_df.copy()
    df["主采图片文件"] = [download_image(str(url), f"{sku}_main") for sku, url in zip(df["SKU编号"], df["主采图片链接"])]
    df["备采图片文件"] = [download_image(str(url), f"{sku}_backup") for sku, url in zip(df["SKU编号"], df["备采图片链接"])]
    return df


def autosize(ws, df: pd.DataFrame, limit: int = 48) -> None:
    for idx, col in enumerate(df.columns):
        vals = df[col].astype(str).tolist()
        width = min(limit, max([len(str(col))] + [len(v) for v in vals]) + 2)
        ws.set_column(idx, idx, width)
    ws.freeze_panes(1, 0)


def write_df_sheet(writer: pd.ExcelWriter, name: str, df: pd.DataFrame, limit: int = 48) -> None:
    df.to_excel(writer, sheet_name=name, index=False)
    ws = writer.sheets[name]
    autosize(ws, df, limit=limit)


def write_final_sheet(writer: pd.ExcelWriter, final_df: pd.DataFrame) -> None:
    workbook = writer.book
    ws = workbook.add_worksheet("最终选商清单")
    writer.sheets["最终选商清单"] = ws

    header_fmt = workbook.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1, "valign": "top", "text_wrap": True})
    text_fmt = workbook.add_format({"border": 1, "valign": "top", "text_wrap": True})
    link_fmt = workbook.add_format({"border": 1, "font_color": "blue", "underline": 1, "valign": "top", "text_wrap": True})

    columns = [
        "SKU编号", "产品名称", "优先级", "风格线", "场景", "AlphaShop会话链接",
        "主采标题", "主采供应商", "主采价格", "主采起批量", "主采发货地", "主采销量信号",
        "主采服务分", "主采客服响应率", "主采回头率", "主采亮点", "主采图片预览", "主采图片链接",
        "主采商品链接", "主采选择理由", "备采标题", "备采供应商", "备采价格", "备采起批量",
        "备采发货地", "备采销量信号", "备采服务分", "备采客服响应率", "备采回头率", "备采亮点",
        "备采图片预览", "备采图片链接", "备采商品链接", "备采选择理由",
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
            if col in {"AlphaShop会话链接", "主采图片链接", "主采商品链接", "备采图片链接", "备采商品链接"}:
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
        "SKU编号": 10, "产品名称": 22, "优先级": 8, "风格线": 14, "场景": 18, "AlphaShop会话链接": 12,
        "主采标题": 36, "主采供应商": 22, "主采价格": 10, "主采起批量": 12, "主采发货地": 14, "主采销量信号": 14,
        "主采服务分": 10, "主采客服响应率": 12, "主采回头率": 10, "主采亮点": 28, "主采图片预览": 18, "主采图片链接": 12,
        "主采商品链接": 12, "主采选择理由": 34, "备采标题": 36, "备采供应商": 22, "备采价格": 10, "备采起批量": 12,
        "备采发货地": 14, "备采销量信号": 14, "备采服务分": 10, "备采客服响应率": 12, "备采回头率": 10,
        "备采亮点": 28, "备采图片预览": 18, "备采图片链接": 12, "备采商品链接": 12, "备采选择理由": 34,
    }
    for col, width in width_map.items():
        ws.set_column(col_index[col], col_index[col], width)
    ws.freeze_panes(1, 0)


def build_overview(raw_df: pd.DataFrame, final_df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"项目": "第四步目标", "内容": "用 AlphaShop 官方找商能力替代 1688 搜索页批量搜索。"},
            {"项目": "运行日期", "内容": EXPORT_DATE},
            {"项目": "找商 SKU 数", "内容": f"{final_df.shape[0]} 款"},
            {"项目": "原始候选数", "内容": f"{raw_df.shape[0]} 条"},
            {"项目": "每个 SKU 抽取数", "内容": "AlphaShop 当前可见的全部候选结果"},
            {"项目": "风控策略", "内容": "仅在 AlphaShop 找商页串行发送指令，不再批量扫描 1688 搜索页。"},
            {"项目": "来源保留", "内容": "保留 AlphaShop 会话链接和商品直链，方便后续复核。"},
        ]
    )


def build_rules() -> pd.DataFrame:
    return pd.DataFrame(
        [
            ("输入策略", "中文自动输入会乱码，因此对 AlphaShop 使用英文提示词，输出结果仍可抓到中文商品信息。"),
            ("访问策略", "一页一 SKU，串行执行；每个 SKU 结束后等待 8-12 秒。"),
            ("筛选策略", "AlphaShop 原始排名优先，再叠加本地适配分过滤明显跑偏的款。"),
            ("链路策略", "商品链接只从 AlphaShop 结果卡片点击获取，不回到 1688 搜索页。"),
            ("四月适配", "继续排除加绒、九分、拖地、重做旧、短袖、制服感和过度正式款。"),
        ],
        columns=["规则项", "执行方式"],
    )


def write_markdown(final_df: pd.DataFrame) -> None:
    lines = [
        "# 日本男装第四步：AlphaShop 找商版",
        "",
        f"- 日期：{EXPORT_DATE}",
        f"- 最终 SKU：{len(final_df)}",
        "",
        "## 主采结论",
        "",
    ]
    for _, row in final_df.iterrows():
        lines.append(
            f"- {row['SKU编号']} {row['产品名称']}：主采《{row['主采标题']}》 / {row['主采供应商']}；"
            f"备采《{row['备采标题']}》 / {row['备采供应商']}。"
        )
    OUTPUT_MD.write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    prior_df = build_prior_reference()
    all_rows: list[dict[str, object]] = []
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(CDP_URL)
        if not browser.contexts:
            raise RuntimeError("No logged-in browser context found.")
        ctx = browser.contexts[0]
        for sku in SKU_PROMPTS:
            print(f"[AlphaShop] running {sku.sku_id} {sku.product_name}", flush=True)
            all_rows.extend(query_one_sku(ctx, sku, max_rows=None))
            time.sleep(random.uniform(8.0, 12.0))

    raw_df = pd.DataFrame(all_rows)
    raw_df.to_csv(OUTPUT_RAW, index=False, encoding="utf-8-sig")

    final_df, flat_df = choose_final_rows(raw_df, prior_df)
    final_df = attach_images(final_df)

    with pd.ExcelWriter(OUTPUT_XLSX, engine="xlsxwriter") as writer:
        write_df_sheet(writer, "概览", build_overview(raw_df, final_df), limit=60)
        write_df_sheet(writer, "规则与方法", build_rules(), limit=60)
        write_df_sheet(writer, "前序结论_低优先", prior_df, limit=42)
        write_final_sheet(writer, final_df)
        write_df_sheet(writer, "AlphaShop原始结果", raw_df, limit=42)
        write_df_sheet(writer, "最终平面表", flat_df, limit=42)

    write_markdown(final_df)


if __name__ == "__main__":
    main()
