from __future__ import annotations

import json
import math
import time
from pathlib import Path
from urllib.parse import quote

import matplotlib.pyplot as plt
import pandas as pd
import pytrends.request as pytrends_request
from pytrends.request import TrendReq
from urllib3.util.retry import Retry as _Retry


class RetryCompat(_Retry):
    def __init__(self, *args, method_whitelist=None, **kwargs):
        if method_whitelist is not None and "allowed_methods" not in kwargs:
            kwargs["allowed_methods"] = method_whitelist
        super().__init__(*args, **kwargs)


pytrends_request.Retry = RetryCompat


plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Noto Sans CJK JP", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False


OUT_DIR = Path("market_map_output")
OUT_DIR.mkdir(exist_ok=True)

GEO = "JP"
HL = "ja-JP"
TZ = 540
EXPORT_DATE = "2026-03-18"
BASE_SLEEP_SECONDS = 8


KEYWORD_META = {
    "きれいめ メンズ": {"zh": "干净利落男装", "en": "Clean", "layer": "核心风格"},
    "ストリート メンズ": {"zh": "街头风男装", "en": "Street", "layer": "核心风格"},
    "アメカジ メンズ": {"zh": "美式休闲男装", "en": "Americana", "layer": "核心风格"},
    "古着 メンズ": {"zh": "古着男装", "en": "Vintage", "layer": "核心风格"},
    "モード メンズ": {"zh": "时装感男装", "en": "Mode", "layer": "核心风格"},
    "大人カジュアル メンズ": {"zh": "成熟休闲男装", "en": "Adult casual", "layer": "细分风格"},
    "オフィスカジュアル メンズ": {"zh": "办公室休闲男装", "en": "Office casual", "layer": "细分风格"},
    "シンプル メンズ": {"zh": "简约男装", "en": "Simple", "layer": "细分风格"},
    "ミニマル メンズ": {"zh": "极简男装", "en": "Minimal", "layer": "细分风格"},
    "韓国ストリート メンズ": {"zh": "韩式街头男装", "en": "Korean street", "layer": "细分风格"},
    "Y2K メンズ": {"zh": "Y2K男装", "en": "Y2K", "layer": "细分风格"},
    "スケーター メンズ": {"zh": "滑板风男装", "en": "Skater", "layer": "细分风格"},
    "古着ミックス メンズ": {"zh": "古着混搭男装", "en": "Vintage mix", "layer": "细分风格"},
    "ワークウェア メンズ": {"zh": "工装风男装", "en": "Workwear", "layer": "细分风格"},
    "ミリタリー メンズ": {"zh": "军事风男装", "en": "Military", "layer": "细分风格"},
    "アイビー メンズ": {"zh": "常春藤男装", "en": "Ivy", "layer": "细分风格"},
    "プレッピー メンズ": {"zh": "学院派男装", "en": "Preppy", "layer": "细分风格"},
    "モノトーン メンズ": {"zh": "黑白灰男装", "en": "Monotone", "layer": "细分风格"},
    "ドレス メンズ": {"zh": "偏正式男装", "en": "Dressy", "layer": "细分风格"},
    "テック系 メンズ": {"zh": "机能风男装", "en": "Techwear", "layer": "细分风格"},
    "ゴープコア メンズ": {"zh": "山系机能男装", "en": "Gorpcore", "layer": "细分风格"},
    "アウトドア メンズ": {"zh": "户外男装", "en": "Outdoor", "layer": "细分风格"},
    "スポーツミックス メンズ": {"zh": "运动混搭男装", "en": "Sports mix", "layer": "细分风格"},
    "アスレジャー メンズ": {"zh": "运动休闲男装", "en": "Athleisure", "layer": "细分风格"},
    "ゴルフウェア メンズ": {"zh": "高尔夫男装", "en": "Golf wear", "layer": "细分风格"},
    "ワイドパンツ メンズ": {"zh": "男士阔腿裤", "en": "Wide pants", "layer": "单品"},
    "カーゴパンツ メンズ": {"zh": "男士工装裤", "en": "Cargo pants", "layer": "单品"},
    "スラックス メンズ": {"zh": "男士西裤", "en": "Slacks", "layer": "单品"},
    "セットアップ メンズ": {"zh": "男士套装", "en": "Setup", "layer": "单品"},
    "カーディガン メンズ": {"zh": "男士开衫", "en": "Cardigan", "layer": "单品"},
    "ワイドシルエット メンズ": {"zh": "宽松廓形男装", "en": "Wide silhouette", "layer": "版型"},
    "オーバーサイズ メンズ": {"zh": "超大版男装", "en": "Oversized", "layer": "版型"},
    "ジャストサイズ メンズ": {"zh": "合体版男装", "en": "True-to-size", "layer": "版型"},
    "バギー メンズ": {"zh": "宽松垮版男装", "en": "Baggy", "layer": "版型"},
    "通勤 メンズ": {"zh": "通勤男装", "en": "Commuting", "layer": "场景"},
    "休日コーデ メンズ": {"zh": "休闲日穿搭男装", "en": "Weekend outfit", "layer": "场景"},
    "洗える メンズ": {"zh": "可机洗男装", "en": "Washable", "layer": "功能"},
    "ストレッチ メンズ": {"zh": "弹力男装", "en": "Stretch", "layer": "功能"},
    "撥水 メンズ": {"zh": "防泼水男装", "en": "Water repellent", "layer": "功能"},
    "接触冷感 メンズ": {"zh": "凉感男装", "en": "Cool touch", "layer": "功能"},
    "軽量 アウター メンズ": {"zh": "轻量外套男装", "en": "Lightweight outerwear", "layer": "功能"},
}


GROUPS = [
    {
        "slug": "style_core_5y",
        "sheet_label": "核心风格",
        "title": "日本男装核心风格，过去5年",
        "timeframe": "today 5-y",
        "anchor": "きれいめ メンズ",
        "keywords": ["きれいめ メンズ", "ストリート メンズ", "アメカジ メンズ", "古着 メンズ", "モード メンズ"],
        "type": "style",
    },
    {
        "slug": "style_urban_5y",
        "sheet_label": "都市清爽",
        "title": "日本男装都市清爽风格，过去5年",
        "timeframe": "today 5-y",
        "anchor": "きれいめ メンズ",
        "keywords": ["きれいめ メンズ", "大人カジュアル メンズ", "オフィスカジュアル メンズ", "シンプル メンズ", "ミニマル メンズ"],
        "type": "style",
    },
    {
        "slug": "style_youth_5y",
        "sheet_label": "年轻街头",
        "title": "日本男装年轻街头风格，过去5年",
        "timeframe": "today 5-y",
        "anchor": "きれいめ メンズ",
        "keywords": ["きれいめ メンズ", "韓国ストリート メンズ", "Y2K メンズ", "スケーター メンズ", "古着ミックス メンズ"],
        "type": "style",
    },
    {
        "slug": "style_heritage_5y",
        "sheet_label": "美式传承",
        "title": "日本男装美式传承风格，过去5年",
        "timeframe": "today 5-y",
        "anchor": "きれいめ メンズ",
        "keywords": ["きれいめ メンズ", "ワークウェア メンズ", "ミリタリー メンズ", "アイビー メンズ", "プレッピー メンズ"],
        "type": "style",
    },
    {
        "slug": "style_design_5y",
        "sheet_label": "设计与机能",
        "title": "日本男装设计与机能风格，过去5年",
        "timeframe": "today 5-y",
        "anchor": "きれいめ メンズ",
        "keywords": ["きれいめ メンズ", "モノトーン メンズ", "ドレス メンズ", "テック系 メンズ", "ゴープコア メンズ"],
        "type": "style",
    },
    {
        "slug": "style_outdoor_5y",
        "sheet_label": "户外运动",
        "title": "日本男装户外运动风格，过去5年",
        "timeframe": "today 5-y",
        "anchor": "きれいめ メンズ",
        "keywords": ["きれいめ メンズ", "アウトドア メンズ", "スポーツミックス メンズ", "アスレジャー メンズ", "ゴルフウェア メンズ"],
        "type": "style",
    },
    {
        "slug": "item_core_5y",
        "sheet_label": "核心单品",
        "title": "日本男装核心单品，过去5年",
        "timeframe": "today 5-y",
        "anchor": "セットアップ メンズ",
        "keywords": ["セットアップ メンズ", "スラックス メンズ", "ワイドパンツ メンズ", "カーゴパンツ メンズ", "カーディガン メンズ"],
        "type": "item",
    },
    {
        "slug": "silhouette_5y",
        "sheet_label": "版型方向",
        "title": "日本男装版型方向，过去5年",
        "timeframe": "today 5-y",
        "anchor": "ワイドパンツ メンズ",
        "keywords": ["ワイドパンツ メンズ", "ワイドシルエット メンズ", "オーバーサイズ メンズ", "ジャストサイズ メンズ", "バギー メンズ"],
        "type": "silhouette",
    },
    {
        "slug": "scene_12m",
        "sheet_label": "场景需求",
        "title": "日本男装场景需求，过去12个月",
        "timeframe": "today 12-m",
        "anchor": "セットアップ メンズ",
        "keywords": ["セットアップ メンズ", "オフィスカジュアル メンズ", "通勤 メンズ", "休日コーデ メンズ", "ゴルフウェア メンズ"],
        "type": "scene",
    },
    {
        "slug": "function_12m",
        "sheet_label": "功能卖点",
        "title": "日本男装功能卖点，过去12个月",
        "timeframe": "today 12-m",
        "anchor": "洗える メンズ",
        "keywords": ["洗える メンズ", "ストレッチ メンズ", "撥水 メンズ", "接触冷感 メンズ", "軽量 アウター メンズ"],
        "type": "function",
    },
]


def zh_en_label(keyword: str) -> str:
    meta = KEYWORD_META.get(keyword, {})
    zh = meta.get("zh", keyword)
    en = meta.get("en", "")
    return f"{zh}（{en}）" if en else zh


def build_client() -> TrendReq:
    return TrendReq(
        hl=HL,
        tz=TZ,
        timeout=(10, 30),
        retries=0,
        backoff_factor=0,
    )


def trends_link(keywords: list[str], timeframe: str, gprop: str = "") -> str:
    parts = [
        "https://trends.google.com/trends/explore",
        f"?date={quote(timeframe, safe='')}",
        f"&geo={GEO}",
    ]
    if gprop:
        parts.append(f"&gprop={quote(gprop, safe='')}")
    encoded_keywords = ",".join(quote(keyword, safe="") for keyword in keywords)
    parts.append(f"&q={encoded_keywords}")
    return "".join(parts)


def safe_sleep(seconds: int) -> None:
    if seconds > 0:
        time.sleep(seconds)


def fetch_interest_over_time(group: dict[str, object]) -> pd.DataFrame:
    max_attempts = 4
    for attempt in range(1, max_attempts + 1):
        try:
            client = build_client()
            client.build_payload(
                kw_list=group["keywords"],
                cat=0,
                timeframe=group["timeframe"],
                geo=GEO,
                gprop="",
            )
            df = client.interest_over_time()
            if df.empty:
                raise RuntimeError(f"{group['slug']} returned no data")
            return df
        except Exception as exc:
            message = str(exc)
            if attempt == max_attempts:
                raise
            sleep_seconds = 20 if "429" not in message else 75 * attempt
            print(f"[retry] {group['slug']} attempt {attempt} failed: {message[:120]}")
            print(f"[retry] sleep {sleep_seconds}s")
            safe_sleep(sleep_seconds)
    raise RuntimeError(f"Failed to fetch {group['slug']}")


def summarize_group(df: pd.DataFrame, timeframe: str, anchor: str) -> pd.DataFrame:
    data = df.drop(columns=["isPartial"], errors="ignore").copy()
    if timeframe == "today 5-y":
        long_window = min(52, len(data))
        prev_window = data.iloc[max(0, len(data) - 104) : max(0, len(data) - 52)]
        recent_window = min(12, len(data))
        prev_recent = data.iloc[max(0, len(data) - 24) : max(0, len(data) - 12)]
        long_label = "52"
        recent_label = "12"
    else:
        long_window = min(13, len(data))
        prev_window = data.iloc[max(0, len(data) - 26) : max(0, len(data) - 13)]
        recent_window = min(4, len(data))
        prev_recent = data.iloc[max(0, len(data) - 8) : max(0, len(data) - 4)]
        long_label = "13"
        recent_label = "4"

    current_long = data.tail(long_window)
    current_recent = data.tail(recent_window)
    anchor_mean = float(current_long[anchor].mean()) if anchor in current_long.columns else math.nan

    rows: list[dict[str, object]] = []
    for column in data.columns:
        series = data[column]
        current_long_mean = round(float(current_long[column].mean()), 2)
        prev_long_mean = round(float(prev_window[column].mean()), 2) if len(prev_window) else None
        current_recent_mean = round(float(current_recent[column].mean()), 2)
        prev_recent_mean = round(float(prev_recent[column].mean()), 2) if len(prev_recent) else None
        yoy = None
        if prev_long_mean not in (None, 0):
            yoy = round(((current_long_mean - prev_long_mean) / prev_long_mean) * 100, 2)
        recent_change = None
        if prev_recent_mean not in (None, 0):
            recent_change = round(((current_recent_mean - prev_recent_mean) / prev_recent_mean) * 100, 2)
        relative_index = None
        if not math.isnan(anchor_mean) and anchor_mean != 0:
            relative_index = round((current_long_mean / anchor_mean) * 100, 2)
        peak_week = series.idxmax().date().isoformat()
        rows.append(
            {
                "keyword_ja": column,
                "keyword_zh": KEYWORD_META.get(column, {}).get("zh", column),
                "keyword_en": KEYWORD_META.get(column, {}).get("en", ""),
                f"近{long_label}周均值": current_long_mean,
                f"前{long_label}周均值": prev_long_mean,
                "同比变化（%）": yoy,
                f"近{recent_label}周均值": current_recent_mean,
                f"前{recent_label}周均值": prev_recent_mean,
                "近期变化（%）": recent_change,
                "峰值周": peak_week,
                "峰值指数": int(series.max()),
                "最新值": int(series.iloc[-1]),
                "相对锚点指数": relative_index,
            }
        )
    return pd.DataFrame(rows)


def plot_group(df: pd.DataFrame, group: dict[str, object], output_path: Path) -> None:
    plot_df = df.drop(columns=["isPartial"], errors="ignore").copy()
    renamed = {keyword: zh_en_label(keyword) for keyword in plot_df.columns}
    plot_df = plot_df.rename(columns=renamed)
    plt.figure(figsize=(12, 6))
    for column in plot_df.columns:
        plt.plot(plot_df.index, plot_df[column], linewidth=2, label=column)
    plt.title(group["title"])
    plt.xlabel("周")
    plt.ylabel("趋势指数（0-100）")
    plt.grid(alpha=0.25)
    plt.legend(fontsize=9)
    plt.tight_layout()
    plt.savefig(output_path, dpi=180)
    plt.close()


def classify_style_opportunity(relative_index: float | None, yoy: float | None, recent: float | None) -> str:
    rel = -999 if relative_index is None else relative_index
    yoy_val = -999 if yoy is None else yoy
    recent_val = -999 if recent is None else recent
    if rel >= 140 and yoy_val >= -5:
        return "高需求主赛道"
    if rel >= 80 and yoy_val >= 8:
        return "成长机会"
    if rel >= 80 and yoy_val < -8:
        return "大盘仍大但在回落"
    if rel >= 35 and yoy_val >= 10:
        return "细分潜力"
    if recent_val >= 10 and rel >= 20:
        return "近期抬头"
    return "观察项"


def build_style_market_table(group_results: list[dict[str, object]]) -> pd.DataFrame:
    frames = []
    for result in group_results:
        if result["type"] != "style":
            continue
        summary = result["summary_df"].copy()
        summary = summary[summary["keyword_ja"] != result["anchor"]].copy()
        summary["风格分组"] = result["sheet_label"]
        summary["机会判断"] = summary.apply(
            lambda row: classify_style_opportunity(
                row["相对锚点指数"], row["同比变化（%）"], row["近期变化（%）"]
            ),
            axis=1,
        )
        frames.append(summary)
    market = pd.concat(frames, ignore_index=True)
    market = market[
        [
            "风格分组",
            "keyword_ja",
            "keyword_zh",
            "keyword_en",
            "相对锚点指数",
            "近52周均值",
            "同比变化（%）",
            "近期变化（%）",
            "最新值",
            "机会判断",
        ]
    ]
    market = market.sort_values(["相对锚点指数", "同比变化（%）"], ascending=[False, False])
    market = market.rename(
        columns={
            "keyword_ja": "日文关键词",
            "keyword_zh": "中文关键词",
            "keyword_en": "英文关键词",
        }
    )
    return market


def build_item_market_table(group_results: list[dict[str, object]]) -> pd.DataFrame:
    target_slugs = {"item_core_5y", "silhouette_5y", "scene_12m", "function_12m"}
    frames = []
    for result in group_results:
        if result["slug"] not in target_slugs:
            continue
        summary = result["summary_df"].copy()
        summary["数据组"] = result["sheet_label"]
        if result["slug"] in {"item_core_5y", "silhouette_5y"}:
            summary = summary[summary["keyword_ja"] != result["anchor"]].copy()
        frames.append(summary)
    market = pd.concat(frames, ignore_index=True)
    market = market.rename(
        columns={
            "keyword_ja": "日文关键词",
            "keyword_zh": "中文关键词",
            "keyword_en": "英文关键词",
        }
    )
    return market[
        [
            "数据组",
            "日文关键词",
            "中文关键词",
            "英文关键词",
            "相对锚点指数",
            next(col for col in market.columns if col.startswith("近") and "周均值" in col),
        ]
    ]


def write_df(writer, sheet_name: str, df: pd.DataFrame, header_fmt) -> None:
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_fmt)
    worksheet.freeze_panes(1, 0)
    for idx, column in enumerate(df.columns):
        lengths = [len(str(column))] + df[column].astype(str).map(len).tolist()
        worksheet.set_column(idx, idx, min(max(lengths) + 2, 42))


def conclusion_lines(style_market: pd.DataFrame, item_core: pd.DataFrame, silhouette: pd.DataFrame, scene: pd.DataFrame, function: pd.DataFrame) -> list[str]:
    top_style = style_market.iloc[0]
    growth_styles = style_market[style_market["机会判断"].isin(["成长机会", "细分潜力"])].head(5)
    top_items = item_core.sort_values("近52周均值", ascending=False).head(3)
    silhouette_top = silhouette.sort_values("相对锚点指数", ascending=False).head(2)
    scene_top = scene.sort_values("近13周均值", ascending=False).head(2)
    function_top = function.sort_values("近13周均值", ascending=False).head(2)

    lines = [
        f"本次数据抓取时间为 {EXPORT_DATE}，最新周数据截至 2026-03-15 所在周。",
        f"风格层面，当前最大盘的方向是 {top_style['中文关键词']}，相对锚点指数为 {top_style['相对锚点指数']:.2f}。",
        "适合重点关注的成长风格包括：" + "、".join(growth_styles["中文关键词"].head(4).tolist()) + "。",
        "核心单品上，最值得优先开发的是：" + "、".join(top_items["keyword_zh"].tolist()) + "。",
        "版型层面，当前更偏向：" + "、".join(silhouette_top["keyword_zh"].tolist()) + "。",
        "消费场景层面，近12个月最强的是：" + "、".join(scene_top["keyword_zh"].tolist()) + "。",
        "功能卖点层面，近12个月最有感知的是：" + "、".join(function_top["keyword_zh"].tolist()) + "。",
    ]
    return lines


def export_workbook(group_results: list[dict[str, object]], style_market: pd.DataFrame) -> Path:
    workbook_path = OUT_DIR / f"日本男装市场机会图谱_{EXPORT_DATE}.xlsx"
    with pd.ExcelWriter(workbook_path, engine="xlsxwriter") as writer:
        workbook = writer.book
        title_fmt = workbook.add_format({"bold": True, "font_size": 16})
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
        wrap_fmt = workbook.add_format({"text_wrap": True, "valign": "top"})
        link_fmt = workbook.add_format({"font_color": "blue", "underline": 1})

        overview = workbook.add_worksheet("概览")
        writer.sheets["概览"] = overview
        overview.write("A1", "日本男装市场机会图谱", title_fmt)
        overview.write("A3", "导出日期")
        overview.write("B3", EXPORT_DATE)
        overview.write("A4", "地区")
        overview.write("B4", "日本")
        overview.write("A5", "数据源")
        overview.write("B5", "Google 趋势（Google Trends）网页搜索")
        overview.write("A7", "数据组", header_fmt)
        overview.write_row("A8", ["数据组名称", "时间范围", "用途", "Google 趋势链接"], header_fmt)
        row = 8
        for result in group_results:
            usage = {
                "style": "判断风格方向",
                "item": "判断核心单品",
                "silhouette": "判断版型方向",
                "scene": "判断消费场景",
                "function": "判断功能卖点",
            }[result["type"]]
            overview.write_row(row, 0, [result["sheet_label"], result["timeframe"], usage])
            overview.write_url(row, 3, result["link"], link_fmt, "打开 Google 趋势")
            row += 1
        overview.write("A20", "提醒", header_fmt)
        notes = [
            "同一张图内的 0-100 指数可以横向比较，不同图之间的绝对值不能直接横比。",
            "风格机会表中的“相对锚点指数”使用固定锚点做近似归一化，用于跨批次观察强弱。",
            "本次没有使用 Google Trends 服饰分类，而是使用日本区网页搜索，以保证关键词覆盖更全。",
        ]
        for idx, note in enumerate(notes, start=21):
            overview.write(f"A{idx}", note, wrap_fmt)
        overview.set_column("A:A", 20)
        overview.set_column("B:C", 18)
        overview.set_column("D:D", 26)

        keyword_rows = []
        for keyword, meta in KEYWORD_META.items():
            keyword_rows.append(
                {
                    "日文关键词": keyword,
                    "中文关键词": meta["zh"],
                    "英文关键词": meta["en"],
                    "层级": meta["layer"],
                }
            )
        write_df(writer, "关键词对照", pd.DataFrame(keyword_rows).sort_values(["层级", "日文关键词"]), header_fmt)
        write_df(writer, "风格机会排序", style_market, header_fmt)

        for result in group_results:
            raw_df = result["df"].copy().rename(columns={keyword: KEYWORD_META[keyword]["zh"] for keyword in result["df"].columns if keyword in KEYWORD_META})
            write_df(writer, f"{result['sheet_label']}汇总", result["summary_df"], header_fmt)
            write_df(writer, f"{result['sheet_label']}原始", raw_df.reset_index(), header_fmt)
            chart_ws = workbook.add_worksheet(f"{result['sheet_label']}图表")
            writer.sheets[f"{result['sheet_label']}图表"] = chart_ws
            chart_ws.write("A1", result["title"], title_fmt)
            chart_ws.write_url("A3", result["link"], link_fmt, "打开 Google 趋势")
            chart_ws.insert_image("A5", str(result["chart_path"]), {"x_scale": 0.75, "y_scale": 0.75})
            chart_ws.set_column("A:A", 42)

        item_core = next(result["summary_df"] for result in group_results if result["slug"] == "item_core_5y")
        silhouette = next(result["summary_df"] for result in group_results if result["slug"] == "silhouette_5y")
        scene = next(result["summary_df"] for result in group_results if result["slug"] == "scene_12m")
        function = next(result["summary_df"] for result in group_results if result["slug"] == "function_12m")

        conclusion = workbook.add_worksheet("结论摘要")
        writer.sheets["结论摘要"] = conclusion
        conclusion.write("A1", "结论摘要", title_fmt)
        for idx, line in enumerate(conclusion_lines(style_market, item_core, silhouette, scene, function), start=3):
            conclusion.write(f"A{idx}", line, wrap_fmt)
        conclusion.set_column("A:A", 80)

    return workbook_path


def build_report(group_results: list[dict[str, object]], style_market: pd.DataFrame) -> Path:
    report_path = OUT_DIR / "商业分析摘要.md"
    item_core = next(result["summary_df"] for result in group_results if result["slug"] == "item_core_5y")
    silhouette = next(result["summary_df"] for result in group_results if result["slug"] == "silhouette_5y")
    scene = next(result["summary_df"] for result in group_results if result["slug"] == "scene_12m")
    function = next(result["summary_df"] for result in group_results if result["slug"] == "function_12m")

    high_track = style_market[style_market["机会判断"] == "高需求主赛道"].head(5)
    growth_track = style_market[style_market["机会判断"].isin(["成长机会", "细分潜力"])].head(8)
    falling_track = style_market[style_market["机会判断"] == "大盘仍大但在回落"].head(5)

    lines = [
        f"# 日本男装市场机会分析",
        "",
        f"- 抓取日期：{EXPORT_DATE}",
        f"- 最新周数据：截至 2026-03-15 所在周",
        "",
        "## 风格判断",
        "- 高需求主赛道：" + "、".join(high_track["中文关键词"].tolist()) if not high_track.empty else "- 高需求主赛道：暂无明显项",
        "- 成长机会：" + "、".join(growth_track["中文关键词"].tolist()) if not growth_track.empty else "- 成长机会：暂无明显项",
        "- 大盘回落风格：" + "、".join(falling_track["中文关键词"].tolist()) if not falling_track.empty else "- 大盘回落风格：暂无明显项",
        "",
        "## 单品判断",
        "- 近52周均值前3：" + "、".join(item_core.sort_values("近52周均值", ascending=False)["keyword_zh"].head(3).tolist()),
        "- 近52周同比增长前2：" + "、".join(item_core.sort_values("同比变化（%）", ascending=False)["keyword_zh"].head(2).tolist()),
        "",
        "## 版型判断",
        "- 当前更强的版型方向：" + "、".join(silhouette.sort_values("相对锚点指数", ascending=False)["keyword_zh"].head(3).tolist()),
        "",
        "## 场景判断",
        "- 近13周均值前3：" + "、".join(scene.sort_values("近13周均值", ascending=False)["keyword_zh"].head(3).tolist()),
        "",
        "## 功能判断",
        "- 近13周均值前3：" + "、".join(function.sort_values("近13周均值", ascending=False)["keyword_zh"].head(3).tolist()),
    ]
    report_path.write_text("\n".join(lines), encoding="utf-8")
    return report_path


def main() -> None:
    results: list[dict[str, object]] = []
    manifest_rows = []
    for index, group in enumerate(GROUPS, start=1):
        print(f"[{index}/{len(GROUPS)}] fetching {group['slug']}")
        df = fetch_interest_over_time(group)
        summary_df = summarize_group(df, group["timeframe"], group["anchor"])

        csv_path = OUT_DIR / f"{group['slug']}.csv"
        summary_path = OUT_DIR / f"{group['slug']}_summary.csv"
        chart_path = OUT_DIR / f"{group['slug']}.png"
        df.to_csv(csv_path, encoding="utf-8-sig")
        summary_df.to_csv(summary_path, index=False, encoding="utf-8-sig")
        plot_group(df, group, chart_path)

        result = {
            "slug": group["slug"],
            "sheet_label": group["sheet_label"],
            "title": group["title"],
            "timeframe": group["timeframe"],
            "type": group["type"],
            "anchor": group["anchor"],
            "link": trends_link(group["keywords"], group["timeframe"]),
            "csv_path": csv_path,
            "summary_path": summary_path,
            "chart_path": chart_path,
            "df": df,
            "summary_df": summary_df,
        }
        results.append(result)
        manifest_rows.append(
            {
                "slug": group["slug"],
                "sheet_label": group["sheet_label"],
                "timeframe": group["timeframe"],
                "link": result["link"],
                "csv": str(csv_path),
                "summary": str(summary_path),
                "chart": str(chart_path),
            }
        )
        if index < len(GROUPS):
            safe_sleep(BASE_SLEEP_SECONDS)

    style_market = build_style_market_table(results)
    style_market_path = OUT_DIR / "style_market_table.csv"
    style_market.to_csv(style_market_path, index=False, encoding="utf-8-sig")

    manifest_path = OUT_DIR / "manifest.json"
    manifest_path.write_text(json.dumps(manifest_rows, ensure_ascii=False, indent=2), encoding="utf-8")

    workbook_path = export_workbook(results, style_market)
    report_path = build_report(results, style_market)

    print(f"manifest={manifest_path}")
    print(f"style_market={style_market_path}")
    print(f"workbook={workbook_path}")
    print(f"report={report_path}")


if __name__ == "__main__":
    main()
