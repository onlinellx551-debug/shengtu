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


OUT_DIR = Path("april_sell_output")
OUT_DIR.mkdir(exist_ok=True)

GEO = "JP"
HL = "ja-JP"
TZ = 540
EXPORT_DATE = "2026-03-19"
LATEST_WEEK = "2026-03-15"
BASE_SLEEP_SECONDS = 8


KEYWORD_META = {
    "きれいめ メンズ": {"zh": "干净利落男装", "en": "Clean"},
    "オフィスカジュアル メンズ": {"zh": "办公室休闲男装", "en": "Office casual"},
    "アメカジ メンズ": {"zh": "美式休闲男装", "en": "Americana"},
    "ストリート メンズ": {"zh": "街头风男装", "en": "Street"},
    "古着 メンズ": {"zh": "古着男装", "en": "Vintage"},
    "シンプル メンズ": {"zh": "简约男装", "en": "Simple"},
    "モード メンズ": {"zh": "时装感男装", "en": "Mode"},
    "プレッピー メンズ": {"zh": "学院派男装", "en": "Preppy"},
    "テック系 メンズ": {"zh": "机能风男装", "en": "Techwear"},
    "セットアップ メンズ": {"zh": "男士套装", "en": "Setup"},
    "スラックス メンズ": {"zh": "男士西裤", "en": "Slacks"},
    "ワイドパンツ メンズ": {"zh": "男士阔腿裤", "en": "Wide pants"},
    "デニム メンズ": {"zh": "男士牛仔裤", "en": "Denim"},
    "カーゴパンツ メンズ": {"zh": "男士工装裤", "en": "Cargo pants"},
    "ライトアウター メンズ": {"zh": "轻外套男装", "en": "Light outerwear"},
    "ブルゾン メンズ": {"zh": "男士夹克", "en": "Blouson"},
    "シャツジャケット メンズ": {"zh": "男士衬衫夹克", "en": "Shirt jacket"},
    "カーディガン メンズ": {"zh": "男士开衫", "en": "Cardigan"},
    "ナイロンジャケット メンズ": {"zh": "男士尼龙夹克", "en": "Nylon jacket"},
    "シャツ メンズ": {"zh": "男士衬衫", "en": "Shirt"},
    "白シャツ メンズ": {"zh": "男士白衬衫", "en": "White shirt"},
    "カットソー メンズ": {"zh": "男士打底衫", "en": "Cut and sew"},
    "ロンt メンズ": {"zh": "男士长袖 T 恤", "en": "Long sleeve tee"},
    "ポロシャツ メンズ": {"zh": "男士 POLO 衫", "en": "Polo shirt"},
    "ワイドシルエット メンズ": {"zh": "宽松廓形男装", "en": "Wide silhouette"},
    "オーバーサイズ メンズ": {"zh": "超大版男装", "en": "Oversized"},
    "バギー メンズ": {"zh": "宽松垮版男装", "en": "Baggy"},
    "ジャストサイズ メンズ": {"zh": "合体版男装", "en": "True-to-size"},
    "春服 メンズ": {"zh": "春装男装", "en": "Spring menswear"},
    "通勤 メンズ": {"zh": "通勤男装", "en": "Commuting"},
    "休日コーデ メンズ": {"zh": "休闲日穿搭男装", "en": "Weekend outfit"},
    "ゴルフウェア メンズ": {"zh": "高尔夫男装", "en": "Golf wear"},
    "洗える メンズ": {"zh": "可机洗男装", "en": "Washable"},
    "ストレッチ メンズ": {"zh": "弹力男装", "en": "Stretch"},
    "撥水 メンズ": {"zh": "防泼水男装", "en": "Water repellent"},
    "接触冷感 メンズ": {"zh": "凉感男装", "en": "Cool touch"},
    "軽量 アウター メンズ": {"zh": "轻量外套男装", "en": "Lightweight outerwear"},
}


GROUPS = [
    {
        "slug": "style_main_5y",
        "label": "主流风格",
        "title": "日本男装主流风格，过去5年",
        "keywords": ["きれいめ メンズ", "オフィスカジュアル メンズ", "アメカジ メンズ", "ストリート メンズ", "古着 メンズ"],
        "anchor": "きれいめ メンズ",
        "kind": "style",
    },
    {
        "slug": "style_sub_5y",
        "label": "补充风格",
        "title": "日本男装补充风格，过去5年",
        "keywords": ["きれいめ メンズ", "シンプル メンズ", "モード メンズ", "プレッピー メンズ", "テック系 メンズ"],
        "anchor": "きれいめ メンズ",
        "kind": "style",
    },
    {
        "slug": "bottoms_5y",
        "label": "裤装与套装",
        "title": "日本男装裤装与套装，过去5年",
        "keywords": ["セットアップ メンズ", "スラックス メンズ", "ワイドパンツ メンズ", "デニム メンズ", "カーゴパンツ メンズ"],
        "anchor": "セットアップ メンズ",
        "kind": "item",
    },
    {
        "slug": "outerwear_5y",
        "label": "春季外套",
        "title": "日本男装春季外套，过去5年",
        "keywords": ["ライトアウター メンズ", "ブルゾン メンズ", "シャツジャケット メンズ", "カーディガン メンズ", "ナイロンジャケット メンズ"],
        "anchor": "ライトアウター メンズ",
        "kind": "item",
    },
    {
        "slug": "tops_5y",
        "label": "春季上装",
        "title": "日本男装春季上装，过去5年",
        "keywords": ["シャツ メンズ", "白シャツ メンズ", "カットソー メンズ", "ロンt メンズ", "ポロシャツ メンズ"],
        "anchor": "シャツ メンズ",
        "kind": "item",
    },
    {
        "slug": "silhouette_5y",
        "label": "版型方向",
        "title": "日本男装版型方向，过去5年",
        "keywords": ["ワイドパンツ メンズ", "ワイドシルエット メンズ", "オーバーサイズ メンズ", "バギー メンズ", "ジャストサイズ メンズ"],
        "anchor": "ワイドパンツ メンズ",
        "kind": "silhouette",
    },
    {
        "slug": "scene_5y",
        "label": "四月场景",
        "title": "日本男装四月消费场景，过去5年",
        "keywords": ["春服 メンズ", "オフィスカジュアル メンズ", "通勤 メンズ", "休日コーデ メンズ", "ゴルフウェア メンズ"],
        "anchor": "春服 メンズ",
        "kind": "scene",
    },
    {
        "slug": "function_5y",
        "label": "功能卖点",
        "title": "日本男装功能卖点，过去5年",
        "keywords": ["洗える メンズ", "ストレッチ メンズ", "撥水 メンズ", "接触冷感 メンズ", "軽量 アウター メンズ"],
        "anchor": "洗える メンズ",
        "kind": "function",
    },
]


def zh(keyword: str) -> str:
    return KEYWORD_META[keyword]["zh"]


def zh_en(keyword: str) -> str:
    meta = KEYWORD_META[keyword]
    return f"{meta['zh']}（{meta['en']}）"


def build_client() -> TrendReq:
    return TrendReq(
        hl=HL,
        tz=TZ,
        timeout=(10, 30),
        retries=0,
        backoff_factor=0,
    )


def trends_link(keywords: list[str]) -> str:
    q = ",".join(quote(keyword, safe="") for keyword in keywords)
    return f"https://trends.google.com/trends/explore?date=today%205-y&geo={GEO}&q={q}"


def safe_sleep(seconds: int) -> None:
    if seconds > 0:
        time.sleep(seconds)


def fetch_group(group: dict[str, object]) -> pd.DataFrame:
    csv_path = OUT_DIR / f"{group['slug']}.csv"
    if csv_path.exists():
        return pd.read_csv(csv_path, encoding="utf-8-sig", index_col=0, parse_dates=True)

    max_attempts = 4
    for attempt in range(1, max_attempts + 1):
        try:
            client = build_client()
            client.build_payload(
                kw_list=group["keywords"],
                cat=0,
                timeframe="today 5-y",
                geo=GEO,
                gprop="",
            )
            df = client.interest_over_time()
            if df.empty:
                raise RuntimeError(f"{group['slug']} returned empty data")
            df.to_csv(csv_path, encoding="utf-8-sig")
            return df
        except Exception as exc:
            if attempt == max_attempts:
                raise
            message = str(exc)
            sleep_seconds = 20 if "429" not in message else 75 * attempt
            print(f"[retry] {group['slug']} attempt {attempt} failed: {message[:120]}")
            print(f"[retry] sleep {sleep_seconds}s")
            safe_sleep(sleep_seconds)
    raise RuntimeError(f"failed to fetch {group['slug']}")


def best_month(series: pd.Series) -> tuple[int, float]:
    monthly = series.groupby(series.index.month).mean()
    month = int(monthly.idxmax())
    value = round(float(monthly.max()), 2)
    return month, value


def month_name(month: int) -> str:
    names = {
        1: "1月",
        2: "2月",
        3: "3月",
        4: "4月",
        5: "5月",
        6: "6月",
        7: "7月",
        8: "8月",
        9: "9月",
        10: "10月",
        11: "11月",
        12: "12月",
    }
    return names[month]


def summarize_april(df: pd.DataFrame, anchor: str) -> pd.DataFrame:
    data = df.drop(columns=["isPartial"], errors="ignore").copy()
    april_mask = data.index.month == 4
    march_mask = data.index.month == 3
    current_52 = data.tail(min(52, len(data)))
    prev_52 = data.iloc[max(0, len(data) - 104) : max(0, len(data) - 52)]
    current_12 = data.tail(min(12, len(data)))
    prev_12 = data.iloc[max(0, len(data) - 24) : max(0, len(data) - 12)]
    anchor_year_mean = float(current_52[anchor].mean()) if anchor in current_52.columns else math.nan

    rows: list[dict[str, object]] = []
    for column in data.columns:
        series = data[column]
        annual_avg = round(float(current_52[column].mean()), 2)
        prev_annual = round(float(prev_52[column].mean()), 2) if len(prev_52) else None
        yoy = None
        if prev_annual not in (None, 0):
            yoy = round(((annual_avg - prev_annual) / prev_annual) * 100, 2)

        recent_avg = round(float(current_12[column].mean()), 2)
        prev_recent = round(float(prev_12[column].mean()), 2) if len(prev_12) else None
        recent_change = None
        if prev_recent not in (None, 0):
            recent_change = round(((recent_avg - prev_recent) / prev_recent) * 100, 2)

        april_avg = round(float(series[april_mask].mean()), 2) if april_mask.any() else 0.0
        march_avg = round(float(series[march_mask].mean()), 2) if march_mask.any() else 0.0
        april_index = None
        if annual_avg != 0:
            april_index = round((april_avg / annual_avg) * 100, 2)
        april_vs_march = None
        if march_avg != 0:
            april_vs_march = round(((april_avg - march_avg) / march_avg) * 100, 2)

        peak_month, peak_month_avg = best_month(series)
        relative_to_anchor = None
        if not math.isnan(anchor_year_mean) and anchor_year_mean != 0:
            relative_to_anchor = round((annual_avg / anchor_year_mean) * 100, 2)

        rows.append(
            {
                "日文关键词": column,
                "中文关键词": zh(column),
                "英文关键词": KEYWORD_META[column]["en"],
                "近52周均值": annual_avg,
                "前52周均值": prev_annual,
                "同比变化（%）": yoy,
                "近12周均值": recent_avg,
                "前12周均值": prev_recent,
                "近期变化（%）": recent_change,
                "历史4月均值": april_avg,
                "历史3月均值": march_avg,
                "4月季节指数": april_index,
                "4月较3月变化（%）": april_vs_march,
                "峰值月份": month_name(peak_month),
                "峰值月均值": peak_month_avg,
                "最新值": int(series.iloc[-1]),
                "相对锚点指数": relative_to_anchor,
            }
        )
    return pd.DataFrame(rows)


def classify_style(row: pd.Series) -> str:
    if row["近52周均值"] < 2 and row["历史4月均值"] < 2:
        return "低搜索量，不单独决策"
    if row["近52周均值"] >= 18 and row["同比变化（%）"] is not None and row["同比变化（%）"] >= 0:
        return "四月主线风格"
    if row["近52周均值"] >= 18 and row["同比变化（%）"] is not None and row["同比变化（%）"] <= -8:
        return "大盘仍大但在回落"
    if row["4月季节指数"] is not None and row["4月季节指数"] >= 110 and row["近52周均值"] >= 5:
        return "四月加分风格"
    if row["同比变化（%）"] is not None and row["同比变化（%）"] >= 10:
        return "成长风格"
    return "观察项"


def classify_product(row: pd.Series) -> str:
    if row["近52周均值"] < 2 and row["历史4月均值"] < 2:
        return "低搜索量，不单独决策"
    if row["历史4月均值"] >= 20 and row["4月季节指数"] is not None and row["4月季节指数"] >= 105 and (row["近期变化（%）"] is None or row["近期变化（%）"] >= -15):
        return "四月主推"
    if row["近52周均值"] >= 18 and row["4月季节指数"] is not None and 90 <= row["4月季节指数"] < 105:
        return "四月稳卖"
    if row["历史4月均值"] >= 8 and row["4月季节指数"] is not None and row["4月季节指数"] >= 115:
        return "四月季节机会"
    if row["近期变化（%）"] is not None and row["近期变化（%）"] >= 10 and row["历史4月均值"] >= 5:
        return "四月可试"
    if (row["4月季节指数"] is not None and row["4月季节指数"] < 85) or (row["近期变化（%）"] is not None and row["近期变化（%）"] <= -20):
        return "四月谨慎"
    return "观察项"


def plot_group(df: pd.DataFrame, group: dict[str, object], output_path: Path) -> None:
    plot_df = df.drop(columns=["isPartial"], errors="ignore").copy()
    plot_df = plot_df.rename(columns={keyword: zh_en(keyword) for keyword in plot_df.columns})
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


def write_df(writer, sheet_name: str, df: pd.DataFrame, header_fmt) -> None:
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_fmt)
    worksheet.freeze_panes(1, 0)
    for idx, column in enumerate(df.columns):
        lengths = [len(str(column))] + df[column].astype(str).map(len).tolist()
        worksheet.set_column(idx, idx, min(max(lengths) + 2, 42))


def build_conclusion(style_table: pd.DataFrame, product_table: pd.DataFrame, function_table: pd.DataFrame) -> list[str]:
    style_main = style_table[style_table["判断"] == "四月主线风格"]["中文关键词"].tolist()
    style_growth = style_table[style_table["判断"].isin(["成长风格", "四月加分风格"])]["中文关键词"].tolist()
    style_decline = style_table[style_table["判断"] == "大盘仍大但在回落"]["中文关键词"].tolist()

    main_products = product_table[product_table["判断"].isin(["四月主推", "四月稳卖"])]["中文关键词"].tolist()
    seasonal_products = product_table[product_table["判断"] == "四月季节机会"]["中文关键词"].tolist()
    cautious_products = product_table[product_table["判断"] == "四月谨慎"]["中文关键词"].tolist()

    top_functions = function_table.sort_values("历史4月均值", ascending=False)["中文关键词"].head(3).tolist()

    return [
        f"本分析基于 Google 趋势（Google Trends）日本区网页搜索数据，抓取日期为 {EXPORT_DATE}，最新周数据截至 {LATEST_WEEK} 所在周。",
        "四月主线风格建议优先围绕：" + "、".join(style_main[:3]) + "。" if style_main else "四月主线风格暂无明显单一答案。",
        "成长或加分风格包括：" + "、".join(style_growth[:4]) + "。" if style_growth else "成长风格信号不明显。",
        "虽然流量仍大但不建议重押的风格包括：" + "、".join(style_decline[:3]) + "。" if style_decline else "大盘回落风格不明显。",
        "四月商品主推建议优先布局：" + "、".join(main_products[:6]) + "。" if main_products else "四月主推商品不明显。",
        "适合做季节补充或胶囊的方向包括：" + "、".join(seasonal_products[:5]) + "。" if seasonal_products else "季节补充方向不明显。",
        "不建议四月重押的方向包括：" + "、".join(cautious_products[:5]) + "。" if cautious_products else "暂未发现明显需要回避的方向。",
        "功能卖点层面，四月更适合强调：" + "、".join(top_functions) + "。",
    ]


def main() -> None:
    results = []
    for index, group in enumerate(GROUPS, start=1):
        print(f"[{index}/{len(GROUPS)}] {group['slug']}")
        df = fetch_group(group)
        summary = summarize_april(df, group["anchor"])
        if group["kind"] == "style":
            summary["判断"] = summary.apply(classify_style, axis=1)
        else:
            summary["判断"] = summary.apply(classify_product, axis=1)
        chart_path = OUT_DIR / f"{group['slug']}.png"
        plot_group(df, group, chart_path)
        summary_path = OUT_DIR / f"{group['slug']}_summary.csv"
        summary.to_csv(summary_path, index=False, encoding="utf-8-sig")
        results.append(
            {
                "group": group,
                "df": df,
                "summary": summary,
                "chart_path": chart_path,
                "summary_path": summary_path,
                "link": trends_link(group["keywords"]),
            }
        )
        if index < len(GROUPS):
            safe_sleep(BASE_SLEEP_SECONDS)

    style_frames = []
    product_frames = []
    function_table = None
    scene_table = None
    silhouette_table = None
    for result in results:
        group = result["group"]
        summary = result["summary"].copy()
        summary["数据组"] = group["label"]
        if group["kind"] == "style":
            summary = summary[summary["日文关键词"] != group["anchor"]].copy()
            style_frames.append(summary)
        elif group["kind"] == "function":
            function_table = summary.copy()
            function_table["数据组"] = group["label"]
        elif group["kind"] == "scene":
            scene_table = summary.copy()
            scene_table["数据组"] = group["label"]
        elif group["kind"] == "silhouette":
            silhouette_table = summary.copy()
            silhouette_table["数据组"] = group["label"]
            product_frames.append(summary)
        else:
            product_frames.append(summary)

    style_table = pd.concat(style_frames, ignore_index=True)
    style_table = style_table.sort_values(["判断", "近52周均值", "历史4月均值"], ascending=[True, False, False])
    product_table = pd.concat(product_frames, ignore_index=True)
    product_table = product_table.sort_values(["判断", "历史4月均值", "近52周均值"], ascending=[True, False, False])
    if function_table is None or scene_table is None or silhouette_table is None:
        raise RuntimeError("Missing function or scene or silhouette table")

    style_table_path = OUT_DIR / "四月风格机会表.csv"
    product_table_path = OUT_DIR / "四月商品机会表.csv"
    function_table_path = OUT_DIR / "四月功能机会表.csv"
    scene_table_path = OUT_DIR / "四月场景机会表.csv"
    style_table.to_csv(style_table_path, index=False, encoding="utf-8-sig")
    product_table.to_csv(product_table_path, index=False, encoding="utf-8-sig")
    function_table.to_csv(function_table_path, index=False, encoding="utf-8-sig")
    scene_table.to_csv(scene_table_path, index=False, encoding="utf-8-sig")

    manifest_rows = []
    for result in results:
        group = result["group"]
        manifest_rows.append(
            {
                "slug": group["slug"],
                "label": group["label"],
                "link": result["link"],
                "summary": str(result["summary_path"]),
                "chart": str(result["chart_path"]),
            }
        )
    manifest_path = OUT_DIR / "manifest.json"
    manifest_path.write_text(json.dumps(manifest_rows, ensure_ascii=False, indent=2), encoding="utf-8")

    workbook_path = OUT_DIR / f"日本男装四月选品分析_{EXPORT_DATE}.xlsx"
    with pd.ExcelWriter(workbook_path, engine="xlsxwriter") as writer:
        workbook = writer.book
        title_fmt = workbook.add_format({"bold": True, "font_size": 16})
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
        wrap_fmt = workbook.add_format({"text_wrap": True, "valign": "top"})
        link_fmt = workbook.add_format({"font_color": "blue", "underline": 1})

        overview = workbook.add_worksheet("概览")
        writer.sheets["概览"] = overview
        overview.write("A1", "日本男装四月选品分析", title_fmt)
        overview.write("A3", "抓取日期")
        overview.write("B3", EXPORT_DATE)
        overview.write("A4", "最新周数据")
        overview.write("B4", LATEST_WEEK)
        overview.write("A5", "说明")
        overview.write("B5", "这份分析专门面向 4 月卖货，重点看历史 4 月表现、4 月相对 3 月变化和最近 12 周动量。")
        overview.write("A7", "数据组", header_fmt)
        overview.write_row("A8", ["数据组", "用途", "Google 趋势链接"], header_fmt)
        row = 8
        usage_map = {
            "style": "判断风格主线",
            "item": "判断主推商品",
            "silhouette": "判断版型",
            "scene": "判断四月场景",
            "function": "判断卖点功能",
        }
        for result in results:
            overview.write_row(row, 0, [result["group"]["label"], usage_map[result["group"]["kind"]]])
            overview.write_url(row, 2, result["link"], link_fmt, "打开 Google 趋势")
            row += 1
        overview.write("A19", "提醒", header_fmt)
        notes = [
            "4 月相关判断是基于过去 5 年历史 4 月表现和当前动量推断出来的，不是对未来销量的直接预测。",
            "同一张图内可以比较强弱，不同图之间不能直接把 0-100 指数横比。",
            "低搜索量词只作为补充观察，不能单独决定选品。",
        ]
        for idx, note in enumerate(notes, start=20):
            overview.write(f"A{idx}", note, wrap_fmt)
        overview.set_column("A:A", 18)
        overview.set_column("B:B", 22)
        overview.set_column("C:C", 28)

        write_df(writer, "四月风格机会", style_table, header_fmt)
        write_df(writer, "四月商品机会", product_table, header_fmt)
        write_df(writer, "四月场景机会", scene_table, header_fmt)
        write_df(writer, "四月功能机会", function_table, header_fmt)

        keyword_rows = []
        for keyword, meta in KEYWORD_META.items():
            keyword_rows.append({"日文关键词": keyword, "中文关键词": meta["zh"], "英文关键词": meta["en"]})
        write_df(writer, "关键词对照", pd.DataFrame(keyword_rows).sort_values("日文关键词"), header_fmt)

        for result in results:
            label = result["group"]["label"]
            write_df(writer, f"{label}汇总", result["summary"], header_fmt)
            raw_df = result["df"].reset_index().rename(columns={"index": "日期"})
            raw_df = raw_df.rename(columns={keyword: zh(keyword) for keyword in result["group"]["keywords"]})
            write_df(writer, f"{label}原始", raw_df, header_fmt)
            chart_ws = workbook.add_worksheet(f"{label}图表")
            writer.sheets[f"{label}图表"] = chart_ws
            chart_ws.write("A1", result["group"]["title"], title_fmt)
            chart_ws.write_url("A3", result["link"], link_fmt, "打开 Google 趋势")
            chart_ws.insert_image("A5", str(result["chart_path"]), {"x_scale": 0.75, "y_scale": 0.75})
            chart_ws.set_column("A:A", 42)

        conclusion_ws = workbook.add_worksheet("结论")
        writer.sheets["结论"] = conclusion_ws
        conclusion_ws.write("A1", "四月商业结论", title_fmt)
        for idx, line in enumerate(build_conclusion(style_table, product_table, function_table), start=3):
            conclusion_ws.write(f"A{idx}", line, wrap_fmt)
        conclusion_ws.set_column("A:A", 90)

    report_lines = [
        "# 日本男装四月选品分析",
        "",
        f"- 抓取日期：{EXPORT_DATE}",
        f"- 最新周数据：截至 {LATEST_WEEK} 所在周",
        "",
        "## 风格机会",
        style_table.to_csv(index=False),
        "",
        "## 商品机会",
        product_table.to_csv(index=False),
        "",
        "## 场景机会",
        scene_table.to_csv(index=False),
        "",
        "## 功能机会",
        function_table.to_csv(index=False),
    ]
    report_path = OUT_DIR / "四月商业分析.md"
    report_path.write_text("\n".join(report_lines), encoding="utf-8")

    print(f"workbook={workbook_path}")
    print(f"style_table={style_table_path}")
    print(f"product_table={product_table_path}")
    print(f"scene_table={scene_table_path}")
    print(f"function_table={function_table_path}")
    print(f"manifest={manifest_path}")


if __name__ == "__main__":
    main()
