from __future__ import annotations

import json
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


OUT_DIR = Path("trend_output")
OUT_DIR.mkdir(exist_ok=True)

JP_TZ = 540
HL = "ja-JP"
GEO = "JP"


def build_client() -> TrendReq:
    return TrendReq(
        hl=HL,
        tz=JP_TZ,
        timeout=(10, 30),
        retries=2,
        backoff_factor=0.3,
    )


GROUPS = [
    {
        "slug": "style_web_5y",
        "title": "Japan menswear style terms, Web Search, past 5 years",
        "timeframe": "today 5-y",
        "gprop": "",
        "keywords": [
            ("\u304d\u308c\u3044\u3081 \u30e1\u30f3\u30ba", "Clean"),
            ("\u30b9\u30c8\u30ea\u30fc\u30c8 \u30e1\u30f3\u30ba", "Street"),
            ("\u30a2\u30e1\u30ab\u30b8 \u30e1\u30f3\u30ba", "Americana"),
            ("\u53e4\u7740 \u30e1\u30f3\u30ba", "Vintage"),
            ("\u97d3\u56fd\u30d5\u30a1\u30c3\u30b7\u30e7\u30f3 \u30e1\u30f3\u30ba", "K-fashion"),
        ],
    },
    {
        "slug": "item_web_5y",
        "title": "Japan menswear item terms, Web Search, past 5 years",
        "timeframe": "today 5-y",
        "gprop": "",
        "keywords": [
            ("\u30ef\u30a4\u30c9\u30d1\u30f3\u30c4 \u30e1\u30f3\u30ba", "Wide pants"),
            ("\u30ab\u30fc\u30b4\u30d1\u30f3\u30c4 \u30e1\u30f3\u30ba", "Cargo pants"),
            ("\u30b9\u30e9\u30c3\u30af\u30b9 \u30e1\u30f3\u30ba", "Slacks"),
            ("\u30bb\u30c3\u30c8\u30a2\u30c3\u30d7 \u30e1\u30f3\u30ba", "Setup"),
            ("\u30ab\u30fc\u30c7\u30a3\u30ac\u30f3 \u30e1\u30f3\u30ba", "Cardigan"),
        ],
    },
    {
        "slug": "item_shopping_12m",
        "title": "Japan menswear item terms, Shopping Search, past 12 months",
        "timeframe": "today 12-m",
        "gprop": "froogle",
        "keywords": [
            ("\u30ef\u30a4\u30c9\u30d1\u30f3\u30c4 \u30e1\u30f3\u30ba", "Wide pants"),
            ("\u30ab\u30fc\u30b4\u30d1\u30f3\u30c4 \u30e1\u30f3\u30ba", "Cargo pants"),
            ("\u30b9\u30e9\u30c3\u30af\u30b9 \u30e1\u30f3\u30ba", "Slacks"),
            ("\u30bb\u30c3\u30c8\u30a2\u30c3\u30d7 \u30e1\u30f3\u30ba", "Setup"),
            ("\u30ab\u30fc\u30c7\u30a3\u30ac\u30f3 \u30e1\u30f3\u30ba", "Cardigan"),
        ],
    },
]


def trends_link(keywords: list[str], timeframe: str, gprop: str) -> str:
    comparison = [
        {"keyword": keyword, "geo": GEO, "time": timeframe} for keyword in keywords
    ]
    req = {"comparisonItem": comparison, "category": 0, "property": gprop}
    return "https://trends.google.com/trends/explore?geo=JP&q=" + quote(
        json.dumps(req, ensure_ascii=False, separators=(",", ":"))
    )


def month_label(ts: pd.Timestamp) -> str:
    return f"{ts.year}-{ts.month:02d}"


def summarize(df: pd.DataFrame) -> pd.DataFrame:
    if "isPartial" in df.columns:
        df = df.drop(columns=["isPartial"])
    last_52 = df.tail(min(52, len(df)))
    prev_52 = df.iloc[max(0, len(df) - 104) : max(0, len(df) - 52)]

    rows: list[dict[str, object]] = []
    for column in df.columns:
        series = df[column]
        last_mean = round(float(last_52[column].mean()), 2)
        prev_mean = round(float(prev_52[column].mean()), 2) if len(prev_52) else None
        yoy = None
        if prev_mean not in (None, 0):
            yoy = round(((last_mean - prev_mean) / prev_mean) * 100, 2)
        peak_date = series.idxmax()
        peak_value = int(series.max())
        rows.append(
            {
                "keyword": column,
                "last_52w_avg": last_mean,
                "prev_52w_avg": prev_mean,
                "yoy_pct": yoy,
                "peak_week": peak_date.date().isoformat(),
                "peak_value": peak_value,
                "latest_value": int(series.iloc[-1]),
            }
        )
    return pd.DataFrame(rows).sort_values("last_52w_avg", ascending=False)


def fetch_related_queries(client: TrendReq, keywords: list[str]) -> dict[str, list[dict[str, object]]]:
    related = client.related_queries()
    result: dict[str, list[dict[str, object]]] = {}
    for keyword in keywords:
        bucket = related.get(keyword, {})
        rising = bucket.get("rising")
        if rising is None or rising.empty:
            result[keyword] = []
            continue
        top_rows = []
        for _, row in rising.head(5).iterrows():
            top_rows.append({"query": row["query"], "value": row["value"]})
        result[keyword] = top_rows
    return result


def plot_group(df: pd.DataFrame, labels: dict[str, str], title: str, output: Path) -> None:
    plot_df = df.drop(columns=["isPartial"], errors="ignore").rename(columns=labels)
    plt.figure(figsize=(12, 6))
    for column in plot_df.columns:
        plt.plot(plot_df.index, plot_df[column], linewidth=2, label=column)
    plt.title(title)
    plt.ylabel("Trend index (0-100)")
    plt.xlabel("Week")
    plt.grid(alpha=0.25)
    plt.legend()
    plt.tight_layout()
    plt.savefig(output, dpi=180)
    plt.close()


def run_group(group: dict[str, object]) -> dict[str, object]:
    client = build_client()
    keyword_pairs = group["keywords"]
    keywords = [keyword for keyword, _ in keyword_pairs]
    labels = {keyword: label for keyword, label in keyword_pairs}
    client.build_payload(
        kw_list=keywords,
        cat=0,
        timeframe=group["timeframe"],
        geo=GEO,
        gprop=group["gprop"],
    )

    df = client.interest_over_time()
    csv_path = OUT_DIR / f"{group['slug']}.csv"
    chart_path = OUT_DIR / f"{group['slug']}.png"
    summary_path = OUT_DIR / f"{group['slug']}_summary.csv"
    related_path = OUT_DIR / f"{group['slug']}_related.json"

    df.to_csv(csv_path, encoding="utf-8-sig")
    summarize(df).to_csv(summary_path, index=False, encoding="utf-8-sig")
    plot_group(df, labels, group["title"], chart_path)

    related = fetch_related_queries(client, keywords)
    related_path.write_text(
        json.dumps(related, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    return {
        "slug": group["slug"],
        "csv": str(csv_path),
        "summary": str(summary_path),
        "chart": str(chart_path),
        "related": str(related_path),
        "link": trends_link(keywords, group["timeframe"], group["gprop"]),
    }


def main() -> None:
    outputs = [run_group(group) for group in GROUPS]
    manifest = OUT_DIR / "manifest.json"
    manifest.write_text(json.dumps(outputs, ensure_ascii=False, indent=2), encoding="utf-8")
    print(manifest)


if __name__ == "__main__":
    main()
