from __future__ import annotations

import json
from pathlib import Path

import pandas as pd


BASE_DIR = Path("trend_output")
SOURCE_PATH = BASE_DIR / "japan_menswear_google_trends_analysis_2026-03-18_cn_v2.xlsx"
MANIFEST_PATH = BASE_DIR / "manifest.json"
OUTPUT_PATH = BASE_DIR / "japan_menswear_google_trends_analysis_2026-03-18_zh.xlsx"

SHEET_NAME_MAP = {
    "Overview": "概览",
    "Keyword Map": "关键词对照",
    "Header Map": "字段对照",
    "Related Query Map": "关联词对照",
    "Style Universe": "风格词库",
    "Trend Plan": "分析计划",
    "Search Modifiers": "搜索后缀",
    "Style Summary": "风格汇总",
    "Style Raw": "风格原始数据",
    "Style Related": "风格关联词",
    "Style Chart": "风格图表",
    "Item Summary": "单品汇总",
    "Item Raw": "单品原始数据",
    "Item Related": "单品关联词",
    "Item Chart": "单品图表",
    "Shop Summary": "购物汇总",
    "Shop Raw": "购物原始数据",
    "Shop Related": "购物关联词",
    "Shop Chart": "购物图表",
}

LAYER_MAP = {
    "Core style": "核心风格",
    "Sub-style": "细分风格",
    "Silhouette": "版型",
    "Scene": "场景",
    "Age segment": "年龄段",
    "Function": "功能",
}

GROUP_MAP = {
    "Mainstream style": "主流风格",
    "Regional trend": "区域趋势",
    "Urban clean": "都市清爽",
    "Street youth": "年轻街头",
    "American heritage": "美式传承",
    "Design-led": "设计导向",
    "Outdoor tech": "户外机能",
    "Scene style": "场景风格",
    "Fit and shape": "版型与轮廓",
    "Usage occasion": "使用场景",
    "Age demand": "年龄需求",
    "Functional demand": "功能需求",
}

BATCH_MAP = {
    "Style baseline": "风格基线",
    "Urban clean": "都市清爽",
    "Street youth": "年轻街头",
    "American heritage": "美式传承",
    "Outdoor and tech": "户外与机能",
    "Silhouette shift": "版型变化",
    "Scene demand": "场景需求",
    "Age and function": "年龄与功能",
}

MODIFIER_TYPE_MAP = {
    "Gender or target": "人群限定",
    "Outfit intent": "穿搭意图",
    "Recommendation": "推荐意图",
    "Popularity": "热度意图",
    "Brand search": "品牌搜索",
    "Season": "季节",
    "Scene": "场景",
    "Pain point": "痛点",
}

SUMMARY_LABELS = {
    "last_52w_avg": "近52周均值",
    "prev_52w_avg": "前52周均值",
    "yoy_pct": "同比变化（%）",
    "peak_week": "峰值周",
    "peak_value": "峰值指数",
    "latest_value": "最新值",
}


def autosize(worksheet, dataframe: pd.DataFrame) -> None:
    for idx, col in enumerate(dataframe.columns):
        series = dataframe[col].astype(str)
        max_len = max([len(str(col))] + series.map(len).tolist())
        worksheet.set_column(idx, idx, min(max_len + 2, 42))


def write_df(writer, sheet_name: str, dataframe: pd.DataFrame, header_fmt) -> None:
    dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    for col_num, value in enumerate(dataframe.columns.values):
        worksheet.write(0, col_num, value, header_fmt)
    autosize(worksheet, dataframe)
    worksheet.freeze_panes(1, 0)


def combine_zh_en(zh: str, en: str) -> str:
    zh = "" if pd.isna(zh) else str(zh)
    en = "" if pd.isna(en) else str(en)
    return f"{zh}（{en}）" if zh and en else zh or en


def build_lookup(keyword_map: pd.DataFrame, style_universe: pd.DataFrame) -> dict[str, str]:
    lookup: dict[str, str] = {}
    for _, row in keyword_map.iterrows():
        lookup[str(row["keyword_ja"])] = combine_zh_en(row["keyword_zh_cn"], row["keyword_en"])
    for _, row in style_universe.iterrows():
        lookup[str(row["keyword_ja"])] = combine_zh_en(row["keyword_zh_cn"], row["keyword_en"])
    return lookup


def translate_compare_terms(text: str, lookup: dict[str, str]) -> str:
    if pd.isna(text):
        return ""
    parts = [part.strip() for part in str(text).split("/") if part.strip()]
    translated = [lookup.get(part, part) for part in parts]
    return " / ".join(translated)


def transform_keyword_map(df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "日文关键词": df["keyword_ja"],
            "中文关键词（English）": [combine_zh_en(zh, en) for zh, en in zip(df["keyword_zh_cn"], df["keyword_en"])],
        }
    )


def transform_header_map(df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "数据组": df["dataset"],
            "原字段": df["header_original"],
            "中文说明（English）": [combine_zh_en(zh, en) for zh, en in zip(df["header_zh_cn"], df["header_en"])],
        }
    )


def transform_related_query_map(df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame({"关联词（日文）": df["query_ja"], "关联词中文": df["query_zh_cn"]})


def transform_style_universe(df: pd.DataFrame, lookup: dict[str, str]) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "层级": df["layer"].map(LAYER_MAP).fillna(df["layer"]),
            "分组": df["group_name"].map(GROUP_MAP).fillna(df["group_name"]),
            "日文关键词": df["keyword_ja"],
            "中文关键词（English）": [combine_zh_en(zh, en) for zh, en in zip(df["keyword_zh_cn"], df["keyword_en"])],
            "为什么看这个词": df["why_track"],
            "建议对比词": df["compare_with"].map(lambda value: translate_compare_terms(value, lookup)),
        }
    )


def transform_trend_plan(df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "步骤": df["step"],
            "批次名称": df["batch_name"].map(BATCH_MAP).fillna(df["batch_name"]),
            "目标": df["goal_zh_cn"],
            "关键词组（日文）": df["keywords_ja"],
            "关键词组（中文）": df["keywords_zh_cn"],
            "怎么看": df["how_to_read"],
        }
    )


def transform_modifiers(df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "后缀类型": df["modifier_type"].map(MODIFIER_TYPE_MAP).fillna(df["modifier_type"]),
            "后缀（日文）": df["append_ja"],
            "后缀（中文）": df["append_zh_cn"],
            "用途": df["use_case"],
        }
    )


def transform_summary(df: pd.DataFrame) -> pd.DataFrame:
    result = pd.DataFrame(
        {
            "日文关键词": df["keyword_ja"],
            "中文关键词（English）": [combine_zh_en(zh, en) for zh, en in zip(df["keyword_zh_cn"], df["keyword_en"])],
        }
    )
    for old, new in SUMMARY_LABELS.items():
        result[new] = df[old]
    return result


def transform_raw(df: pd.DataFrame, lookup: dict[str, str]) -> pd.DataFrame:
    rename_map = {"date": "日期", "isPartial": "是否为未完结周数据"}
    for column in df.columns:
        if column not in rename_map:
            rename_map[column] = lookup.get(column, column)
    return df.rename(columns=rename_map)


def transform_related(df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "日文关键词": df["keyword_ja"],
            "中文关键词（English）": [combine_zh_en(zh, en) for zh, en in zip(df["keyword_zh_cn"], df["keyword_en"])],
            "关联词（日文）": df["query_ja"],
            "关联词中文": df["query_zh_cn"],
            "数值": df["value"],
        }
    )


def write_overview(writer, manifest: list[dict[str, str]], title_fmt, header_fmt, wrap_fmt, link_fmt) -> None:
    workbook = writer.book
    worksheet = workbook.add_worksheet("概览")
    writer.sheets["概览"] = worksheet

    worksheet.write("A1", "日本男装 Google 趋势（Google Trends）分析", title_fmt)
    worksheet.write("A3", "导出日期")
    worksheet.write("B3", "2026-03-18")
    worksheet.write("A4", "地区")
    worksheet.write("B4", "日本")
    worksheet.write("A5", "说明")
    worksheet.write("B5", "本工作簿使用中文表达，必要英文放在中文后面的括号里。")

    worksheet.write("A7", "数据组", header_fmt)
    worksheet.write_row("A8", ["来源标识", "汇总表", "原始数据", "关联词", "图表", "Google 趋势链接"], header_fmt)

    row = 8
    row_map = {
        "style_web_5y": ("风格组（style_web_5y）", "风格汇总", "风格原始数据", "风格关联词", "风格图表"),
        "item_web_5y": ("单品组（item_web_5y）", "单品汇总", "单品原始数据", "单品关联词", "单品图表"),
        "item_shopping_12m": ("购物组（item_shopping_12m）", "购物汇总", "购物原始数据", "购物关联词", "购物图表"),
    }
    for item in manifest:
        label_row = row_map[item["slug"]]
        worksheet.write_row(row, 0, label_row)
        worksheet.write_url(row, 5, item["link"], link_fmt, "打开 Google 趋势")
        row += 1

    worksheet.write("A14", "使用提醒", header_fmt)
    notes = [
        "每张图里的指数都是 0-100 的相对值，不同图之间不要直接横比。",
        "单看风格词不够，需要把风格、版型、场景、年龄和功能一起看。",
        "购物搜索数据更稀疏，适合作为辅助判断，不适合单独下结论。",
    ]
    for idx, note in enumerate(notes, start=15):
        worksheet.write(f"A{idx}", note, wrap_fmt)

    worksheet.set_column("A:A", 24)
    worksheet.set_column("B:E", 20)
    worksheet.set_column("F:F", 28)


def main() -> None:
    source_sheets = pd.read_excel(SOURCE_PATH, sheet_name=None)
    manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))

    keyword_map = source_sheets["Keyword Map"].copy()
    style_universe = source_sheets["Style Universe"].copy()
    lookup = build_lookup(keyword_map, style_universe)

    with pd.ExcelWriter(OUTPUT_PATH, engine="xlsxwriter") as writer:
        workbook = writer.book
        title_fmt = workbook.add_format({"bold": True, "font_size": 16})
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
        wrap_fmt = workbook.add_format({"text_wrap": True, "valign": "top"})
        link_fmt = workbook.add_format({"font_color": "blue", "underline": 1})

        write_overview(writer, manifest, title_fmt, header_fmt, wrap_fmt, link_fmt)
        write_df(writer, "关键词对照", transform_keyword_map(keyword_map), header_fmt)
        write_df(writer, "字段对照", transform_header_map(source_sheets["Header Map"]), header_fmt)
        write_df(writer, "关联词对照", transform_related_query_map(source_sheets["Related Query Map"]), header_fmt)
        write_df(writer, "风格词库", transform_style_universe(style_universe, lookup), header_fmt)
        write_df(writer, "分析计划", transform_trend_plan(source_sheets["Trend Plan"]), header_fmt)
        write_df(writer, "搜索后缀", transform_modifiers(source_sheets["Search Modifiers"]), header_fmt)

        data_sheet_map = {
            "风格汇总": transform_summary(source_sheets["Style Summary"]),
            "风格原始数据": transform_raw(source_sheets["Style Raw"], lookup),
            "风格关联词": transform_related(source_sheets["Style Related"]),
            "单品汇总": transform_summary(source_sheets["Item Summary"]),
            "单品原始数据": transform_raw(source_sheets["Item Raw"], lookup),
            "单品关联词": transform_related(source_sheets["Item Related"]),
            "购物汇总": transform_summary(source_sheets["Shop Summary"]),
            "购物原始数据": transform_raw(source_sheets["Shop Raw"], lookup),
            "购物关联词": transform_related(source_sheets["Shop Related"]),
        }
        for sheet_name, dataframe in data_sheet_map.items():
            write_df(writer, sheet_name, dataframe, header_fmt)

        chart_map = {
            "风格图表": ("style_web_5y", "图表关键词请对照“关键词对照”和“风格词库”工作表。"),
            "单品图表": ("item_web_5y", "图表关键词请对照“关键词对照”工作表。"),
            "购物图表": ("item_shopping_12m", "购物搜索更稀疏，建议结合前两组一起判断。"),
        }
        manifest_lookup = {item["slug"]: item["link"] for item in manifest}
        for sheet_name, (slug, note) in chart_map.items():
            worksheet = workbook.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = worksheet
            worksheet.write("A1", sheet_name, title_fmt)
            worksheet.write("A3", note)
            worksheet.write_url("A4", manifest_lookup[slug], link_fmt, "打开 Google 趋势")
            worksheet.insert_image("A6", str(BASE_DIR / f"{slug}.png"), {"x_scale": 0.75, "y_scale": 0.75})
            worksheet.set_column("A:A", 52)

    print(OUTPUT_PATH)


if __name__ == "__main__":
    main()
