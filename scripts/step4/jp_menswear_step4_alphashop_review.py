from __future__ import annotations

import hashlib
import math
import re
from io import BytesIO
from pathlib import Path

import pandas as pd
import requests
from PIL import Image


EXPORT_DATE = "2026-03-20"
ROOT = Path(__file__).resolve().parent
STEP3_DIR = ROOT / "step3_output"
OUT_DIR = ROOT / "step4_output"
IMG_DIR = OUT_DIR / "alphashop_review_images"
RAW_CSV = OUT_DIR / f"alphashop_supplier_raw_{EXPORT_DATE}.csv"
OUTPUT_XLSX = OUT_DIR / f"日本男装第四步_AlphaShop复选版_{EXPORT_DATE}_修复版.xlsx"
OUTPUT_MD = OUT_DIR / f"第四步_AlphaShop复选版结论_{EXPORT_DATE}_修复版.md"

IMG_DIR.mkdir(exist_ok=True)
OUT_DIR.mkdir(exist_ok=True)

REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/146.0.0.0 Safari/537.36"
    ),
    "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
}


SKU_RULES: dict[str, dict[str, list[str]]] = {
    "O01": {
        "must": ["单西", "西装外套", "西装", "夹克"],
        "prefer": ["海军蓝", "藏青", "轻薄", "弹力", "免烫", "通勤", "商务休闲", "可机洗"],
        "avoid": ["新郎", "结婚", "礼服", "修身", "伴郎", "燕尾", "爸爸", "中年", "韩版"],
    },
    "S01": {
        "must": ["白", "白色", "免烫", "抗皱", "长袖"],
        "prefer": ["白", "白色", "免烫", "抗皱", "长袖", "商务", "通勤", "DP", "纯棉"],
        "avoid": ["米黄", "米黄色", "字母", "印花", "黑", "黑白", "立领", "中山领", "韩版", "制服", "DK"],
    },
    "P01": {
        "must": ["西裤", "西装裤", "直筒"],
        "prefer": ["西裤", "西装裤", "直筒", "弹力", "抗皱", "垂感", "商务", "通勤", "藏青", "海军蓝"],
        "avoid": ["修身", "微喇", "喇叭", "工作服", "职业装", "中老年", "桑蚕丝", "高腰正装"],
    },
    "P02": {
        "must": ["西裤", "西装裤", "直筒", "单褶", "有褶"],
        "prefer": ["单褶", "直筒", "垂感", "西裤", "西装裤", "通勤", "宽松"],
        "avoid": ["阔腿", "萝卜裤", "毛呢", "秋冬", "剑道", "韩版", "中年", "拖地"],
    },
    "D01": {
        "must": ["原色", "赤耳", "靛蓝", "丹宁", "直筒"],
        "prefer": ["原色", "赤耳", "靛蓝", "丹宁", "直筒", "基础", "脱浆"],
        "avoid": ["破洞", "工装", "阔腿", "磨烂", "高街", "天丝", "做旧", "乞丐", "深档"],
    },
    "S02": {
        "must": ["条纹", "牛津", "衬衫"],
        "prefer": ["条纹", "蓝", "白", "牛津", "扣领", "长袖", "衬衫", "职场", "商务休闲"],
        "avoid": ["纯色", "修身", "校服", "制服", "紫条", "灰条"],
    },
    "K01": {
        "must": ["V领", "开衫", "针织"],
        "prefer": ["V领", "开衫", "针织", "纯色", "薄款", "基础", "春秋"],
        "avoid": ["字母", "提花", "背心", "无袖", "校服", "童装"],
    },
    "D02": {
        "must": ["黑", "黑色", "直筒", "牛仔"],
        "prefer": ["黑", "黑色", "直筒", "宽松", "牛仔", "轻商务", "基础"],
        "avoid": ["低腰", "喇叭", "迷彩", "工装", "阔腿", "拖地", "做旧", "弯刀"],
    },
    "T01": {
        "must": ["白", "白色", "T恤", "短袖", "厚实", "重磅"],
        "prefer": ["白", "白色", "T恤", "短袖", "重磅", "厚实", "不透", "圆领", "纯棉"],
        "avoid": ["长袖", "德绒", "卫衣", "立领", "印花", "加厚保暖", "拉链", "秋冬", "中式"],
    },
    "J01": {
        "must": ["牛仔夹克", "牛仔外套", "短款", "短宽", "boxy"],
        "prefer": ["短款", "boxy", "原色", "牛仔夹克", "牛仔外套", "正肩", "复古"],
        "avoid": ["做旧", "垫肩", "工装", "宽松", "衬衣", "植绒", "机车"],
    },
}


def latest_non_lock_xlsx(directory: Path) -> Path:
    files = [path for path in directory.glob("*.xlsx") if not path.name.startswith("~$")]
    if not files:
        raise FileNotFoundError(f"未找到可读取的 Excel 文件: {directory}")
    return max(files, key=lambda path: path.stat().st_mtime)


def clean_lines(text: str) -> list[str]:
    return [line.strip() for line in re.split(r"[\r\n]+", str(text)) if line.strip()]


def value_before_label(text: str, label: str) -> str:
    lines = clean_lines(text)
    for index, line in enumerate(lines):
        if line == label and index > 0:
            return lines[index - 1]
    return ""


def pct_to_num(value: str) -> float:
    text = str(value or "").replace("%", "").strip()
    if not text or text == "-":
        return 0.0
    if text.startswith("<"):
        numbers = re.findall(r"[0-9.]+", text)
        return float(numbers[0]) / 2 if numbers else 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def sales_to_num(value: str) -> float:
    text = str(value or "").strip()
    if not text or text == "-":
        return 0.0
    if text.startswith("<"):
        numbers = re.findall(r"[0-9.]+", text)
        return float(numbers[0]) / 2 if numbers else 0.0
    multiplier = 1.0
    if "万" in text:
        multiplier = 10000.0
    elif "千" in text:
        multiplier = 1000.0
    numbers = re.findall(r"[0-9.]+", text)
    if not numbers:
        return 0.0
    return float(numbers[0]) * multiplier


def num_to_float(value: str) -> float:
    text = str(value or "").strip()
    if not text or text == "-":
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def moq_to_num(value: str) -> float:
    numbers = re.findall(r"[0-9.]+", str(value or ""))
    return float(numbers[0]) if numbers else 0.0


def price_to_num(value: object) -> float:
    numbers = re.findall(r"[0-9.]+", str(value or ""))
    return float(numbers[0]) if numbers else 0.0


def has_valid_link(url: str) -> bool:
    return "detail.1688.com/offer/" in str(url or "")


def keyword_hits(title: str, words: list[str]) -> list[str]:
    title_lower = str(title).lower()
    hits: list[str] = []
    for word in words:
        if word.lower() in title_lower:
            hits.append(word)
    return hits


def must_have_score(hits: list[str]) -> float:
    if len(hits) >= 2:
        return 16.0
    if len(hits) == 1:
        return 6.0
    return -26.0


def build_review_score(row: pd.Series) -> float:
    rules = SKU_RULES[row["sku_id"]]
    must_hits = keyword_hits(row["title"], rules["must"])
    prefer_hits = keyword_hits(row["title"], rules["prefer"])
    avoid_hits = keyword_hits(row["title"], rules["avoid"])

    score = 100.0
    score += (11.0 - float(row["alpha_rank"])) * 4.0
    score += min(20.0, math.log10(float(row["销量数"]) + 1.0) * 5.0)
    score += min(10.0, float(row["综合服务分数"]) * 2.0)
    score += float(row["客服响应率数"]) / 10.0
    score += float(row["90天回头率数"]) / 18.0
    score -= float(row["品质退款率数"]) * 2.0
    score += min(8.0, math.log10(float(row["30天订单数"]) + 1.0) * 2.0)
    score += min(8.0, math.log10(float(row["180天买家数"]) + 1.0) * 2.0)
    score += must_have_score(must_hits)
    score += min(len(prefer_hits), 5) * 6.0
    score -= min(len(avoid_hits), 5) * 18.0

    supplier_tags = str(row.get("supplier_tags", ""))
    if "源头工厂" in supplier_tags:
        score += 6.0
    if "实力商家" in supplier_tags:
        score += 5.0
    if "超级工厂" in supplier_tags:
        score += 4.0

    moq_num = float(row["起批量数"])
    if 0 < moq_num <= 2:
        score += 5.0
    elif 0 < moq_num <= 5:
        score += 3.0
    elif moq_num >= 10:
        score -= 3.0

    if not bool(row["直链有效"]):
        score -= 60.0

    return round(score, 2)


def build_reason(row: pd.Series) -> str:
    ai_highlight = row.get("AI亮点", row.get("highlights", ""))
    parts = [
        f"AlphaShop 排名 {int(row['alpha_rank'])}",
        f"复选分 {row['复选分']}",
        f"起批量 {row['moq'] or '-'}",
        f"销量 {row['sales_signal'] or '-'}",
        f"客服响应率 {row['客服响应率_修正'] or '-'}",
        f"90天回头率 {row['90天回头率_修正'] or '-'}",
        f"品质退款率 {row['品质退款率_修正'] or '-'}",
    ]
    if row["匹配词"]:
        parts.append(f"匹配词 {row['匹配词']}")
    if ai_highlight:
        parts.append(f"AI亮点 {ai_highlight}")
    return "；".join(parts)


def build_risk(row: pd.Series) -> str:
    risks: list[str] = []
    if row["风险词"]:
        risks.append(f"标题含 {row['风险词']}")
    if not bool(row["直链有效"]):
        risks.append("商品直链异常")
    if float(row["客服响应率数"]) and float(row["客服响应率数"]) < 50.0:
        risks.append("客服响应率偏低")
    if float(row["品质退款率数"]) > 2.0:
        risks.append("品质退款率偏高")
    if float(row["起批量数"]) >= 10.0:
        risks.append("起批量偏高")
    if not risks:
        risks.append("主要风险在版型和面料，需要打样复核")
    return "；".join(risks)


def enrich_raw(raw_df: pd.DataFrame, step3_df: pd.DataFrame) -> pd.DataFrame:
    df = raw_df.copy()
    df["品质退款率_修正"] = df["raw_text"].apply(lambda text: value_before_label(text, "品质退款率"))
    df["客服响应率_修正"] = df["raw_text"].apply(lambda text: value_before_label(text, "客服响应率"))
    df["90天回头率_修正"] = df["raw_text"].apply(lambda text: value_before_label(text, "90天回头率"))
    df["综合服务分_修正"] = df["raw_text"].apply(lambda text: value_before_label(text, "综合服务分"))
    df["30天订单_修正"] = df["raw_text"].apply(lambda text: value_before_label(text, "30天订单"))
    df["180天买家_修正"] = df["raw_text"].apply(lambda text: value_before_label(text, "180天买家"))
    df["全部商品_修正"] = df["raw_text"].apply(lambda text: value_before_label(text, "全部商品"))

    df["销量数"] = df["sales_signal"].apply(sales_to_num)
    df["客服响应率数"] = df["客服响应率_修正"].apply(pct_to_num)
    df["90天回头率数"] = df["90天回头率_修正"].apply(pct_to_num)
    df["品质退款率数"] = df["品质退款率_修正"].apply(pct_to_num)
    df["综合服务分数"] = df["综合服务分_修正"].apply(num_to_float)
    df["30天订单数"] = df["30天订单_修正"].apply(num_to_float)
    df["180天买家数"] = df["180天买家_修正"].apply(num_to_float)
    df["起批量数"] = df["moq"].apply(moq_to_num)
    df["价格数"] = df["price"].apply(price_to_num)
    df["直链有效"] = df["product_link"].apply(has_valid_link)

    df["匹配词"] = df.apply(
        lambda row: "、".join(keyword_hits(str(row["title"]), SKU_RULES[row["sku_id"]]["prefer"])),
        axis=1,
    )
    df["风险词"] = df.apply(
        lambda row: "、".join(keyword_hits(str(row["title"]), SKU_RULES[row["sku_id"]]["avoid"])),
        axis=1,
    )
    df["必须词命中"] = df.apply(
        lambda row: "、".join(keyword_hits(str(row["title"]), SKU_RULES[row["sku_id"]]["must"])),
        axis=1,
    )
    df["复选分"] = df.apply(build_review_score, axis=1)
    df["评分理由"] = df.apply(build_reason, axis=1)
    df["主要风险"] = df.apply(build_risk, axis=1)
    df["AI亮点"] = df["highlights"].fillna("")

    step3_meta = step3_df.rename(columns={"SKU编号": "sku_id"}).copy()
    return df.merge(step3_meta, on="sku_id", how="left")


def shortlist_by_sku(df: pd.DataFrame) -> pd.DataFrame:
    rows: list[pd.DataFrame] = []
    for sku_id, group in df.groupby("sku_id", sort=False):
        unique_group = (
            group.sort_values(["复选分", "alpha_rank"], ascending=[False, True])
            .drop_duplicates(subset=["supplier", "title"], keep="first")
            .head(3)
            .copy()
        )
        unique_group["入围位次"] = range(1, len(unique_group) + 1)
        rows.append(unique_group)
    return pd.concat(rows, ignore_index=True)


def pick_backup(group: pd.DataFrame, main_row: pd.Series) -> pd.Series:
    for _, row in group.iterrows():
        if row["product_link"] != main_row["product_link"] and row["supplier"] != main_row["supplier"]:
            return row
    if len(group) > 1:
        return group.iloc[1]
    return main_row


def build_final_table(shortlist_df: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    order = shortlist_df["sku_id"].drop_duplicates().tolist()
    for sku_id in order:
        group = shortlist_df[shortlist_df["sku_id"] == sku_id].sort_values(
            ["复选分", "alpha_rank"], ascending=[False, True]
        )
        main_row = group.iloc[0]
        backup_row = pick_backup(group, main_row)
        rows.append(
            {
                "SKU编号": main_row["sku_id"],
                "产品名称": main_row["产品名称"],
                "优先级": main_row["优先级"],
                "风格线": main_row["风格线"],
                "场景": main_row["场景"],
                "推荐颜色": main_row["推荐颜色"],
                "推荐面料": main_row["推荐面料"],
                "推荐版型": main_row["推荐版型"],
                "建议价格带": main_row["建议价格带"],
                "核心卖点": main_row["核心卖点"],
                "关键风险": main_row["关键风险"],
                "日文搜索词": main_row["日文搜索词"],
                "主采图片链接": main_row["image_url"],
                "主采商品标题": main_row["title"],
                "主采供应商": main_row["supplier"],
                "主采价格": main_row["price"],
                "主采起批量": main_row["moq"],
                "主采发货地": main_row["origin"],
                "主采销量信号": main_row["sales_signal"],
                "主采复选分": main_row["复选分"],
                "主采商品直链": main_row["product_link"],
                "主采选择理由": main_row["评分理由"],
                "主采主要风险": main_row["主要风险"],
                "备采图片链接": backup_row["image_url"],
                "备采商品标题": backup_row["title"],
                "备采供应商": backup_row["supplier"],
                "备采价格": backup_row["price"],
                "备采起批量": backup_row["moq"],
                "备采发货地": backup_row["origin"],
                "备采销量信号": backup_row["sales_signal"],
                "备采复选分": backup_row["复选分"],
                "备采商品直链": backup_row["product_link"],
                "备采选择理由": backup_row["评分理由"],
                "备采主要风险": backup_row["主要风险"],
            }
        )
    return pd.DataFrame(rows)


def download_image(url: str) -> str:
    url = str(url or "").strip()
    if not url.startswith("http"):
        return ""

    digest = hashlib.md5(url.encode("utf-8")).hexdigest()[:16]
    out_path = IMG_DIR / f"{digest}.png"
    if out_path.exists():
        return str(out_path)

    try:
        response = requests.get(url, headers=REQUEST_HEADERS, timeout=30)
        response.raise_for_status()
        image = Image.open(BytesIO(response.content)).convert("RGB")
        image.thumbnail((220, 220))
        image.save(out_path, format="PNG")
        return str(out_path)
    except Exception:
        return ""


def autosize(worksheet, df: pd.DataFrame, extra: dict[str, int] | None = None) -> None:
    extra = extra or {}
    for index, column in enumerate(df.columns):
        width = max(len(str(column)), 12)
        if column in extra:
            width = extra[column]
        else:
            sample = df[column].astype(str).head(40).tolist()
            if sample:
                width = min(max(width, max(len(text) for text in sample) + 2), 42)
        worksheet.set_column(index, index, width)


def apply_links(workbook, worksheet, df: pd.DataFrame, columns: list[str]) -> None:
    url_format = workbook.add_format({"font_color": "blue", "underline": 1})
    for column in columns:
        if column not in df.columns:
            continue
        col_index = df.columns.get_loc(column)
        for row_index, value in enumerate(df[column], start=1):
            text = str(value or "").strip()
            if text.startswith("http"):
                worksheet.write_url(row_index, col_index, text, url_format, string=text)


def write_plain_sheet(writer: pd.ExcelWriter, name: str, df: pd.DataFrame, width_map: dict[str, int] | None = None) -> None:
    df.to_excel(writer, sheet_name=name, index=False)
    workbook = writer.book
    worksheet = writer.sheets[name]
    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, len(df), max(len(df.columns) - 1, 0))
    autosize(worksheet, df, extra=width_map)
    apply_links(workbook, worksheet, df, [column for column in df.columns if "链接" in column or "直链" in column])


def write_image_sheet(
    writer: pd.ExcelWriter,
    name: str,
    df: pd.DataFrame,
    image_file_column: str,
    width_map: dict[str, int] | None = None,
) -> None:
    export_df = df.copy()
    export_df.insert(0, "图片预览", "")
    export_df.to_excel(writer, sheet_name=name, index=False)
    workbook = writer.book
    worksheet = writer.sheets[name]
    worksheet.freeze_panes(1, 1)
    worksheet.autofilter(0, 0, len(export_df), max(len(export_df.columns) - 1, 0))
    autosize(worksheet, export_df, extra=width_map)
    worksheet.set_column(0, 0, 14)

    url_columns = [column for column in export_df.columns if "链接" in column or "直链" in column]
    apply_links(workbook, worksheet, export_df, url_columns)

    image_col_index = export_df.columns.get_loc(image_file_column)
    worksheet.set_column(image_col_index, image_col_index, 2, None, {"hidden": True})
    for row_index, image_path in enumerate(export_df[image_file_column], start=1):
        worksheet.set_row(row_index, 76)
        if image_path:
            worksheet.insert_image(
                row_index,
                0,
                image_path,
                {"x_scale": 0.5, "y_scale": 0.5, "x_offset": 4, "y_offset": 4},
            )
        else:
            worksheet.write(row_index, 0, "无图")
        # Clear the helper path if you do not want to expose the local file path.
        worksheet.write(row_index, image_col_index, "")


def write_final_sheet(writer: pd.ExcelWriter, final_df: pd.DataFrame) -> None:
    export_df = final_df.copy()
    export_df.insert(0, "主采图片预览", "")
    export_df.insert(export_df.columns.get_loc("备采图片链接"), "备采图片预览", "")
    export_df.to_excel(writer, sheet_name="最终主备采", index=False)
    workbook = writer.book
    worksheet = writer.sheets["最终主备采"]
    worksheet.freeze_panes(1, 2)
    worksheet.autofilter(0, 0, len(export_df), len(export_df.columns) - 1)

    main_preview_col = export_df.columns.get_loc("主采图片预览")
    backup_preview_col = export_df.columns.get_loc("备采图片预览")
    main_file_col = export_df.columns.get_loc("主采图片文件")
    backup_file_col = export_df.columns.get_loc("备采图片文件")
    worksheet.set_column(main_file_col, main_file_col, 2, None, {"hidden": True})
    worksheet.set_column(backup_file_col, backup_file_col, 2, None, {"hidden": True})

    for row_index, (_, row) in enumerate(export_df.iterrows(), start=1):
        worksheet.set_row(row_index, 82)
        main_file = row["主采图片文件"]
        backup_file = row["备采图片文件"]
        if main_file:
            worksheet.insert_image(
                row_index,
                main_preview_col,
                main_file,
                {"x_scale": 0.52, "y_scale": 0.52, "x_offset": 4, "y_offset": 4},
            )
        else:
            worksheet.write(row_index, main_preview_col, "无图")

        if backup_file:
            worksheet.insert_image(
                row_index,
                backup_preview_col,
                backup_file,
                {"x_scale": 0.52, "y_scale": 0.52, "x_offset": 4, "y_offset": 4},
            )
        else:
            worksheet.write(row_index, backup_preview_col, "无图")

        worksheet.write(row_index, main_file_col, "")
        worksheet.write(row_index, backup_file_col, "")

    autosize(
        worksheet,
        export_df,
        extra={
            "主采图片预览": 14,
            "备采图片预览": 14,
            "产品名称": 24,
            "推荐面料": 24,
            "推荐版型": 24,
            "核心卖点": 22,
            "关键风险": 22,
            "主采商品标题": 40,
            "备采商品标题": 40,
            "主采供应商": 24,
            "备采供应商": 24,
            "主采选择理由": 48,
            "备采选择理由": 48,
            "主采主要风险": 30,
            "备采主要风险": 30,
        },
    )
    apply_links(
        workbook,
        worksheet,
        export_df,
        ["主采图片链接", "主采商品直链", "备采图片链接", "备采商品直链"],
    )


def build_overview(raw_df: pd.DataFrame, shortlist_df: pd.DataFrame, final_df: pd.DataFrame) -> pd.DataFrame:
    valid_links = int(raw_df["直链有效"].sum())
    invalid_links = int((~raw_df["直链有效"]).sum())
    return pd.DataFrame(
        [
            {"项目": "导出日期", "值": EXPORT_DATE},
            {"项目": "原始候选总数", "值": len(raw_df)},
            {"项目": "SKU 数量", "值": raw_df["sku_id"].nunique()},
            {"项目": "有效 1688 直链", "值": valid_links},
            {"项目": "异常直链", "值": invalid_links},
            {"项目": "复选入围数量", "值": len(shortlist_df)},
            {"项目": "最终主备采行数", "值": len(final_df)},
            {"项目": "说明", "值": "这版为修复版，重建自 AlphaShop 原始 CSV，不依赖之前写坏的空文件。"},
        ]
    )


def write_markdown(raw_df: pd.DataFrame, shortlist_df: pd.DataFrame, final_df: pd.DataFrame) -> None:
    lines = [
        "# 日本男装第四步：AlphaShop 复选版（修复版）",
        "",
        f"- 日期：{EXPORT_DATE}",
        f"- 原始候选：{len(raw_df)} 条",
        f"- 有效 1688 直链：{int(raw_df['直链有效'].sum())} 条",
        f"- 异常直链：{int((~raw_df['直链有效']).sum())} 条",
        f"- 复选入围：{len(shortlist_df)} 条",
        "",
        "## 最终主推",
        "",
    ]

    for _, row in final_df.iterrows():
        lines.extend(
            [
                f"### {row['SKU编号']} {row['产品名称']}",
                f"- 主采：{row['主采供应商']} | {row['主采商品标题']} | {row['主采价格']}",
                f"- 备采：{row['备采供应商']} | {row['备采商品标题']} | {row['备采价格']}",
                f"- 选择理由：{row['主采选择理由']}",
                f"- 主要风险：{row['主采主要风险']}",
                "",
            ]
        )

    OUTPUT_MD.write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    if not RAW_CSV.exists():
        raise FileNotFoundError(f"缺少 AlphaShop 原始 CSV：{RAW_CSV}")

    step3_workbook = latest_non_lock_xlsx(STEP3_DIR)
    step3_df = pd.read_excel(step3_workbook, sheet_name=4)
    raw_df = pd.read_csv(RAW_CSV)

    review_df = enrich_raw(raw_df, step3_df)
    shortlist_df = shortlist_by_sku(review_df)
    final_df = build_final_table(shortlist_df)
    bad_link_df = review_df[~review_df["直链有效"]].copy()
    overview_df = build_overview(review_df, shortlist_df, final_df)

    review_df["图片文件"] = review_df["image_url"].apply(download_image)
    shortlist_df["图片文件"] = shortlist_df["image_url"].apply(download_image)
    final_df["主采图片文件"] = final_df["主采图片链接"].apply(download_image)
    final_df["备采图片文件"] = final_df["备采图片链接"].apply(download_image)

    full_export = review_df[
        [
            "sku_id",
            "产品名称",
            "优先级",
            "风格线",
            "场景",
            "alpha_rank",
            "复选分",
            "title",
            "supplier",
            "price",
            "moq",
            "origin",
            "sales_signal",
            "综合服务分_修正",
            "客服响应率_修正",
            "90天回头率_修正",
            "品质退款率_修正",
            "30天订单_修正",
            "180天买家_修正",
            "supplier_tags",
            "AI亮点",
            "必须词命中",
            "匹配词",
            "风险词",
            "评分理由",
            "主要风险",
            "image_url",
            "product_link",
            "直链有效",
            "session_url",
            "图片文件",
        ]
    ].rename(
        columns={
            "sku_id": "SKU编号",
            "alpha_rank": "AlphaShop排名",
            "title": "商品标题",
            "supplier": "供应商",
            "price": "价格",
            "moq": "起批量",
            "origin": "发货地",
            "sales_signal": "销量信号",
            "supplier_tags": "供应商标签",
            "image_url": "图片链接",
            "product_link": "商品直链",
            "session_url": "AlphaShop会话链接",
        }
    )

    shortlist_export = shortlist_df[
        [
            "sku_id",
            "产品名称",
            "优先级",
            "风格线",
            "场景",
            "入围位次",
            "alpha_rank",
            "复选分",
            "title",
            "supplier",
            "price",
            "moq",
            "origin",
            "sales_signal",
            "综合服务分_修正",
            "客服响应率_修正",
            "90天回头率_修正",
            "品质退款率_修正",
            "必须词命中",
            "匹配词",
            "风险词",
            "评分理由",
            "主要风险",
            "image_url",
            "product_link",
            "图片文件",
        ]
    ].rename(
        columns={
            "sku_id": "SKU编号",
            "alpha_rank": "AlphaShop排名",
            "title": "商品标题",
            "supplier": "供应商",
            "price": "价格",
            "moq": "起批量",
            "origin": "发货地",
            "sales_signal": "销量信号",
            "image_url": "图片链接",
            "product_link": "商品直链",
        }
    )

    bad_link_export = bad_link_df[
        [
            "sku_id",
            "产品名称",
            "alpha_rank",
            "title",
            "supplier",
            "image_url",
            "product_link",
            "评分理由",
            "主要风险",
        ]
    ].rename(
        columns={
            "sku_id": "SKU编号",
            "alpha_rank": "AlphaShop排名",
            "title": "商品标题",
            "supplier": "供应商",
            "image_url": "图片链接",
            "product_link": "商品直链",
        }
    )

    with pd.ExcelWriter(OUTPUT_XLSX, engine="xlsxwriter") as writer:
        write_plain_sheet(writer, "概览", overview_df, width_map={"项目": 18, "值": 96})
        write_plain_sheet(writer, "第三步参考_低优先", step3_df, width_map={"产品名称": 24, "核心卖点": 24, "关键风险": 24})
        write_final_sheet(writer, final_df)
        write_image_sheet(
            writer,
            "复选入围30条",
            shortlist_export,
            "图片文件",
            width_map={
                "产品名称": 24,
                "商品标题": 40,
                "供应商": 24,
                "评分理由": 46,
                "主要风险": 28,
                "图片链接": 18,
                "商品直链": 18,
            },
        )
        write_image_sheet(
            writer,
            "AlphaShop全量100条",
            full_export,
            "图片文件",
            width_map={
                "产品名称": 24,
                "商品标题": 40,
                "供应商": 24,
                "评分理由": 46,
                "主要风险": 28,
                "图片链接": 18,
                "商品直链": 18,
                "AlphaShop会话链接": 18,
            },
        )
        write_plain_sheet(writer, "异常链接", bad_link_export, width_map={"商品标题": 40, "供应商": 22, "评分理由": 42, "主要风险": 24})

    write_markdown(review_df, shortlist_df, final_df)
    print(f"已生成: {OUTPUT_XLSX}")
    print(f"已生成: {OUTPUT_MD}")


if __name__ == "__main__":
    main()
