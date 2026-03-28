from __future__ import annotations

import json
from pathlib import Path

import pandas as pd


BASE_DIR = Path("trend_output")
MANIFEST_PATH = BASE_DIR / "manifest.json"
OUTPUT_PATH = BASE_DIR / "japan_menswear_google_trends_analysis_2026-03-18_cn_v2.xlsx"

KEYWORD_TRANSLATIONS: dict[str, tuple[str, str]] = {
    "きれいめ メンズ": ("Clean", "干净利落男装"),
    "ストリート メンズ": ("Street", "街头风男装"),
    "アメカジ メンズ": ("Americana", "美式休闲男装"),
    "古着 メンズ": ("Vintage", "古着男装"),
    "韓国ファッション メンズ": ("K-fashion", "韩系男装"),
    "ワイドパンツ メンズ": ("Wide pants", "男士阔腿裤"),
    "カーゴパンツ メンズ": ("Cargo pants", "男士工装裤"),
    "スラックス メンズ": ("Slacks", "男士西裤"),
    "セットアップ メンズ": ("Setup", "男士套装"),
    "カーディガン メンズ": ("Cardigan", "男士开衫"),
}

RELATED_QUERY_TRANSLATIONS: dict[str, str] = {
    "韓国 ストリート メンズ": "韩式街头男装",
    "アメカジ メンズ 40 代": "40代美式休闲男装",
    "パラシュート パンツ": "伞兵裤",
    "バギー スラックス メンズ": "男士宽松西裤",
    "ワークマン スラックス メンズ": "Workman男士西裤",
    "フレア スラックス メンズ": "男士喇叭西裤",
    "gu スラックス メンズ": "GU男士西裤",
    "ジーユー スラックス メンズ": "GU男士西裤",
    "ギャップ セットアップ メンズ": "GAP男士套装",
    "gap セットアップ メンズ": "GAP男士套装",
    "ゴルフ セットアップ メンズ": "高尔夫男士套装",
    "ニューバランス セットアップ メンズ": "New Balance男士套装",
    "nike セットアップ メンズ": "Nike男士套装",
    "ワークマン カーディガン メンズ": "Workman男士开衫",
    "マルジェラ カーディガン メンズ": "Margiela男士开衫",
    "オーバー サイズ カーディガン メンズ": "男士超大版开衫",
    "モヘア カーディガン メンズ": "男士马海毛开衫",
    "オフィス カジュアル カーディガン メンズ": "男士办公室休闲开衫",
    "クロノス セットアップ メンズ": "CRONOS男士套装",
    "スエット セットアップ メンズ": "男士卫衣套装",
    "アンダー アーマー セットアップ メンズ": "Under Armour男士套装",
    "プーマ セットアップ メンズ": "Puma男士套装",
}

STYLE_UNIVERSE = [
    {
        "layer": "Core style",
        "group_name": "Mainstream style",
        "keyword_ja": "きれいめ メンズ",
        "keyword_en": "Clean",
        "keyword_zh_cn": "干净利落男装",
        "why_track": "判断都会感、通勤友好型审美是否在走强。",
        "compare_with": "大人カジュアル メンズ / オフィスカジュアル メンズ / ミニマル メンズ",
    },
    {
        "layer": "Core style",
        "group_name": "Mainstream style",
        "keyword_ja": "ストリート メンズ",
        "keyword_en": "Street",
        "keyword_zh_cn": "街头风男装",
        "why_track": "判断年轻客群和强趋势感风格的总流量池。",
        "compare_with": "韓国ストリート メンズ / Y2K メンズ / スケーター メンズ",
    },
    {
        "layer": "Core style",
        "group_name": "Mainstream style",
        "keyword_ja": "アメカジ メンズ",
        "keyword_en": "Americana",
        "keyword_zh_cn": "美式休闲男装",
        "why_track": "判断基础款、复古美式和成熟 casual 方向。",
        "compare_with": "ワークウェア メンズ / ミリタリー メンズ / アイビー メンズ",
    },
    {
        "layer": "Core style",
        "group_name": "Mainstream style",
        "keyword_ja": "古着 メンズ",
        "keyword_en": "Vintage",
        "keyword_zh_cn": "古着男装",
        "why_track": "判断复古和二手审美的整体热度。",
        "compare_with": "古着ミックス メンズ / ストリート メンズ / Y2K メンズ",
    },
    {
        "layer": "Core style",
        "group_name": "Mainstream style",
        "keyword_ja": "モード メンズ",
        "keyword_en": "Mode",
        "keyword_zh_cn": "时装感男装",
        "why_track": "判断偏设计感、单色、廓形导向风格。",
        "compare_with": "モノトーン メンズ / ドレス メンズ / ミニマル メンズ",
    },
    {
        "layer": "Core style",
        "group_name": "Regional trend",
        "keyword_ja": "韓国ファッション メンズ",
        "keyword_en": "K-fashion",
        "keyword_zh_cn": "韩系男装",
        "why_track": "判断韩系大方向是否仍有独立热度。",
        "compare_with": "韓国ストリート メンズ / きれいめ メンズ / ストリート メンズ",
    },
    {
        "layer": "Sub-style",
        "group_name": "Urban clean",
        "keyword_ja": "大人カジュアル メンズ",
        "keyword_en": "Adult casual",
        "keyword_zh_cn": "成熟休闲男装",
        "why_track": "对应更高年龄层和更稳定消费力。",
        "compare_with": "きれいめ メンズ / オフィスカジュアル メンズ / シンプル メンズ",
    },
    {
        "layer": "Sub-style",
        "group_name": "Urban clean",
        "keyword_ja": "オフィスカジュアル メンズ",
        "keyword_en": "Office casual",
        "keyword_zh_cn": "办公室休闲男装",
        "why_track": "判断通勤、商务休闲场景的稳定需求。",
        "compare_with": "きれいめ メンズ / 大人カジュアル メンズ / セットアップ メンズ",
    },
    {
        "layer": "Sub-style",
        "group_name": "Urban clean",
        "keyword_ja": "シンプル メンズ",
        "keyword_en": "Simple",
        "keyword_zh_cn": "简约男装",
        "why_track": "判断基础、低风险审美是否更强。",
        "compare_with": "ミニマル メンズ / きれいめ メンズ / モノトーン メンズ",
    },
    {
        "layer": "Sub-style",
        "group_name": "Urban clean",
        "keyword_ja": "ミニマル メンズ",
        "keyword_en": "Minimal",
        "keyword_zh_cn": "极简男装",
        "why_track": "判断低 logo、低装饰、强版型方向。",
        "compare_with": "シンプル メンズ / モード メンズ / モノトーン メンズ",
    },
    {
        "layer": "Sub-style",
        "group_name": "Street youth",
        "keyword_ja": "韓国ストリート メンズ",
        "keyword_en": "Korean street",
        "keyword_zh_cn": "韩式街头男装",
        "why_track": "判断韩系和街头是否在年轻层融合。",
        "compare_with": "ストリート メンズ / 韓国ファッション メンズ / Y2K メンズ",
    },
    {
        "layer": "Sub-style",
        "group_name": "Street youth",
        "keyword_ja": "Y2K メンズ",
        "keyword_en": "Y2K",
        "keyword_zh_cn": "Y2K男装",
        "why_track": "判断短周期潮流和年轻用户审美波动。",
        "compare_with": "ストリート メンズ / 古着 メンズ / スケーター メンズ",
    },
    {
        "layer": "Sub-style",
        "group_name": "Street youth",
        "keyword_ja": "スケーター メンズ",
        "keyword_en": "Skater",
        "keyword_zh_cn": "滑板风男装",
        "why_track": "判断更硬核街头亚文化是否有搜索量。",
        "compare_with": "ストリート メンズ / ワークウェア メンズ / Y2K メンズ",
    },
    {
        "layer": "Sub-style",
        "group_name": "Street youth",
        "keyword_ja": "古着ミックス メンズ",
        "keyword_en": "Vintage mix",
        "keyword_zh_cn": "古着混搭男装",
        "why_track": "判断古着不是单独概念而是混搭方式时的热度。",
        "compare_with": "古着 メンズ / アメカジ メンズ / ストリート メンズ",
    },
    {
        "layer": "Sub-style",
        "group_name": "American heritage",
        "keyword_ja": "ワークウェア メンズ",
        "keyword_en": "Workwear",
        "keyword_zh_cn": "工装风男装",
        "why_track": "判断美式工装、工具感细分方向。",
        "compare_with": "アメカジ メンズ / ミリタリー メンズ / カーゴパンツ メンズ",
    },
    {
        "layer": "Sub-style",
        "group_name": "American heritage",
        "keyword_ja": "ミリタリー メンズ",
        "keyword_en": "Military",
        "keyword_zh_cn": "军事风男装",
        "why_track": "判断机能、军装元素是否在回升。",
        "compare_with": "ワークウェア メンズ / アメカジ メンズ / カーゴパンツ メンズ",
    },
    {
        "layer": "Sub-style",
        "group_name": "American heritage",
        "keyword_ja": "アイビー メンズ",
        "keyword_en": "Ivy",
        "keyword_zh_cn": "常春藤男装",
        "why_track": "判断学院感和传统美式是否回暖。",
        "compare_with": "プレッピー メンズ / アメカジ メンズ / きれいめ メンズ",
    },
    {
        "layer": "Sub-style",
        "group_name": "American heritage",
        "keyword_ja": "プレッピー メンズ",
        "keyword_en": "Preppy",
        "keyword_zh_cn": "学院派男装",
        "why_track": "判断学院感、干净感、年轻通勤方向。",
        "compare_with": "アイビー メンズ / きれいめ メンズ / カーディガン メンズ",
    },
    {
        "layer": "Sub-style",
        "group_name": "Design-led",
        "keyword_ja": "モノトーン メンズ",
        "keyword_en": "Monotone",
        "keyword_zh_cn": "黑白灰男装",
        "why_track": "判断颜色审美是否偏低饱和和安全化。",
        "compare_with": "モード メンズ / シンプル メンズ / ミニマル メンズ",
    },
    {
        "layer": "Sub-style",
        "group_name": "Design-led",
        "keyword_ja": "ドレス メンズ",
        "keyword_en": "Dressy",
        "keyword_zh_cn": "偏正式男装",
        "why_track": "判断更讲究、正式、利落的服装取向。",
        "compare_with": "モード メンズ / きれいめ メンズ / スラックス メンズ",
    },
    {
        "layer": "Sub-style",
        "group_name": "Outdoor tech",
        "keyword_ja": "テック系 メンズ",
        "keyword_en": "Techwear",
        "keyword_zh_cn": "机能风男装",
        "why_track": "判断科技面料、机能细节和城市户外方向。",
        "compare_with": "ゴープコア メンズ / アウトドア メンズ / 撥水 メンズ",
    },
    {
        "layer": "Sub-style",
        "group_name": "Outdoor tech",
        "keyword_ja": "ゴープコア メンズ",
        "keyword_en": "Gorpcore",
        "keyword_zh_cn": "山系机能男装",
        "why_track": "判断近年户外潮流是否有持续搜索基础。",
        "compare_with": "テック系 メンズ / アウトドア メンズ / スポーツミックス メンズ",
    },
    {
        "layer": "Sub-style",
        "group_name": "Outdoor tech",
        "keyword_ja": "アウトドア メンズ",
        "keyword_en": "Outdoor",
        "keyword_zh_cn": "户外男装",
        "why_track": "判断功能和休闲交叉的大盘。",
        "compare_with": "ゴープコア メンズ / テック系 メンズ / スポーツミックス メンズ",
    },
    {
        "layer": "Sub-style",
        "group_name": "Outdoor tech",
        "keyword_ja": "スポーツミックス メンズ",
        "keyword_en": "Sports mix",
        "keyword_zh_cn": "运动混搭男装",
        "why_track": "判断运动品牌化和日常穿搭融合程度。",
        "compare_with": "アスレジャー メンズ / ゴルフウェア メンズ / セットアップ メンズ",
    },
    {
        "layer": "Sub-style",
        "group_name": "Outdoor tech",
        "keyword_ja": "アスレジャー メンズ",
        "keyword_en": "Athleisure",
        "keyword_zh_cn": "运动休闲男装",
        "why_track": "判断舒适、轻运动、日常化需求。",
        "compare_with": "スポーツミックス メンズ / セットアップ メンズ / ゴルフウェア メンズ",
    },
    {
        "layer": "Sub-style",
        "group_name": "Scene style",
        "keyword_ja": "ゴルフウェア メンズ",
        "keyword_en": "Golf wear",
        "keyword_zh_cn": "高尔夫男装",
        "why_track": "判断细分场景带来的中高客单机会。",
        "compare_with": "オフィスカジュアル メンズ / セットアップ メンズ / アスレジャー メンズ",
    },
    {
        "layer": "Silhouette",
        "group_name": "Fit and shape",
        "keyword_ja": "ワイドシルエット メンズ",
        "keyword_en": "Wide silhouette",
        "keyword_zh_cn": "宽松廓形男装",
        "why_track": "判断整体廓形是不是继续偏宽。",
        "compare_with": "オーバーサイズ メンズ / ジャストサイズ メンズ / バギー メンズ",
    },
    {
        "layer": "Silhouette",
        "group_name": "Fit and shape",
        "keyword_ja": "オーバーサイズ メンズ",
        "keyword_en": "Oversized",
        "keyword_zh_cn": "超大版男装",
        "why_track": "判断大廓形是否仍是主流。",
        "compare_with": "ワイドシルエット メンズ / ジャストサイズ メンズ / ストリート メンズ",
    },
    {
        "layer": "Silhouette",
        "group_name": "Fit and shape",
        "keyword_ja": "ジャストサイズ メンズ",
        "keyword_en": "True-to-size",
        "keyword_zh_cn": "合体版男装",
        "why_track": "判断审美是否从大廓形回归合身。",
        "compare_with": "オーバーサイズ メンズ / きれいめ メンズ / スラックス メンズ",
    },
    {
        "layer": "Silhouette",
        "group_name": "Fit and shape",
        "keyword_ja": "バギー メンズ",
        "keyword_en": "Baggy",
        "keyword_zh_cn": "宽松垮版男装",
        "why_track": "判断街头和年轻用户对裤型的偏好。",
        "compare_with": "ワイドシルエット メンズ / ワイドパンツ メンズ / フレアパンツ メンズ",
    },
    {
        "layer": "Silhouette",
        "group_name": "Fit and shape",
        "keyword_ja": "フレアパンツ メンズ",
        "keyword_en": "Flared pants",
        "keyword_zh_cn": "男士喇叭裤",
        "why_track": "判断设计感裤型是否有扩张空间。",
        "compare_with": "バギー メンズ / スラックス メンズ / モード メンズ",
    },
    {
        "layer": "Scene",
        "group_name": "Usage occasion",
        "keyword_ja": "通勤 メンズ",
        "keyword_en": "Commuting",
        "keyword_zh_cn": "通勤男装",
        "why_track": "判断买衣服的实际穿着场景。",
        "compare_with": "オフィスカジュアル メンズ / 休日コーデ メンズ / セットアップ メンズ",
    },
    {
        "layer": "Scene",
        "group_name": "Usage occasion",
        "keyword_ja": "休日コーデ メンズ",
        "keyword_en": "Weekend outfit",
        "keyword_zh_cn": "休闲日穿搭男装",
        "why_track": "判断周末 casual 场景偏好。",
        "compare_with": "通勤 メンズ / ストリート メンズ / アメカジ メンズ",
    },
    {
        "layer": "Age segment",
        "group_name": "Age demand",
        "keyword_ja": "30代 メンズ ファッション",
        "keyword_en": "Menswear in 30s",
        "keyword_zh_cn": "30代男装",
        "why_track": "判断主力消费年龄层需求。",
        "compare_with": "40代 メンズ ファッション / 大人カジュアル メンズ / きれいめ メンズ",
    },
    {
        "layer": "Age segment",
        "group_name": "Age demand",
        "keyword_ja": "40代 メンズ ファッション",
        "keyword_en": "Menswear in 40s",
        "keyword_zh_cn": "40代男装",
        "why_track": "判断成熟市场的稳定性和高客单潜力。",
        "compare_with": "30代 メンズ ファッション / 50代 メンズ ファッション / オフィスカジュアル メンズ",
    },
    {
        "layer": "Age segment",
        "group_name": "Age demand",
        "keyword_ja": "50代 メンズ ファッション",
        "keyword_en": "Menswear in 50s",
        "keyword_zh_cn": "50代男装",
        "why_track": "判断更高年龄段的常青需求。",
        "compare_with": "40代 メンズ ファッション / 大人カジュアル メンズ / シンプル メンズ",
    },
    {
        "layer": "Function",
        "group_name": "Functional demand",
        "keyword_ja": "接触冷感 メンズ",
        "keyword_en": "Cool touch",
        "keyword_zh_cn": "凉感男装",
        "why_track": "判断日本夏季功能诉求和季节爆点。",
        "compare_with": "リネン メンズ / 軽量 アウター メンズ / 洗える メンズ",
    },
    {
        "layer": "Function",
        "group_name": "Functional demand",
        "keyword_ja": "撥水 メンズ",
        "keyword_en": "Water repellent",
        "keyword_zh_cn": "防泼水男装",
        "why_track": "判断机能、通勤和户外融合机会。",
        "compare_with": "テック系 メンズ / アウトドア メンズ / 軽量 アウター メンズ",
    },
    {
        "layer": "Function",
        "group_name": "Functional demand",
        "keyword_ja": "軽量 アウター メンズ",
        "keyword_en": "Lightweight outerwear",
        "keyword_zh_cn": "轻量外套男装",
        "why_track": "判断春秋过渡季需求。",
        "compare_with": "撥水 メンズ / テック系 メンズ / オフィスカジュアル メンズ",
    },
    {
        "layer": "Function",
        "group_name": "Functional demand",
        "keyword_ja": "洗える メンズ",
        "keyword_en": "Washable",
        "keyword_zh_cn": "可机洗男装",
        "why_track": "判断实用主义购买诉求。",
        "compare_with": "オフィスカジュアル メンズ / セットアップ メンズ / 接触冷感 メンズ",
    },
    {
        "layer": "Function",
        "group_name": "Functional demand",
        "keyword_ja": "ストレッチ メンズ",
        "keyword_en": "Stretch",
        "keyword_zh_cn": "弹力男装",
        "why_track": "判断舒适性功能的普遍程度。",
        "compare_with": "洗える メンズ / セットアップ メンズ / スラックス メンズ",
    },
]

TREND_PLAN = [
    {
        "step": 1,
        "batch_name": "Style baseline",
        "goal_zh_cn": "先看日本男装主要风格大盘，确认强流量和下降风格。",
        "keywords_ja": "きれいめ メンズ | ストリート メンズ | アメカジ メンズ | 古着 メンズ | モード メンズ",
        "keywords_zh_cn": "干净利落 | 街头风 | 美式休闲 | 古着 | 时装感",
        "how_to_read": "看过去 5 年和过去 12 个月，判断是上涨、稳定还是回落。",
    },
    {
        "step": 2,
        "batch_name": "Urban clean",
        "goal_zh_cn": "判断通勤友好、成熟客群是否值得做主线。",
        "keywords_ja": "きれいめ メンズ | 大人カジュアル メンズ | オフィスカジュアル メンズ | シンプル メンズ | ミニマル メンズ",
        "keywords_zh_cn": "干净利落 | 成熟休闲 | 办公室休闲 | 简约 | 极简",
        "how_to_read": "看是否有稳定全年需求，以及 30-40 代相关词是否同步走强。",
    },
    {
        "step": 3,
        "batch_name": "Street youth",
        "goal_zh_cn": "判断年轻用户的强趋势风格是否还值得重仓。",
        "keywords_ja": "ストリート メンズ | 韓国ストリート メンズ | Y2K メンズ | スケーター メンズ | 古着ミックス メンズ",
        "keywords_zh_cn": "街头风 | 韩式街头 | Y2K | 滑板风 | 古着混搭",
        "how_to_read": "如果高位回落明显，就更适合做点缀，不适合做主系列。",
    },
    {
        "step": 4,
        "batch_name": "American heritage",
        "goal_zh_cn": "判断美式、工装、学院感这条线有没有长期机会。",
        "keywords_ja": "アメカジ メンズ | ワークウェア メンズ | ミリタリー メンズ | アイビー メンズ | プレッピー メンズ",
        "keywords_zh_cn": "美式休闲 | 工装风 | 军事风 | 常春藤 | 学院派",
        "how_to_read": "看近 52 周均值和同比变化，确认是不是稳步抬升。",
    },
    {
        "step": 5,
        "batch_name": "Outdoor and tech",
        "goal_zh_cn": "判断机能、户外、运动混搭是不是新的增量来源。",
        "keywords_ja": "テック系 メンズ | ゴープコア メンズ | アウトドア メンズ | スポーツミックス メンズ | アスレジャー メンズ",
        "keywords_zh_cn": "机能风 | 山系机能 | 户外 | 运动混搭 | 运动休闲",
        "how_to_read": "看功能词和场景词是否一起上升，而不是单独一两个概念词上涨。",
    },
    {
        "step": 6,
        "batch_name": "Silhouette shift",
        "goal_zh_cn": "判断审美是继续宽松，还是回归合身。",
        "keywords_ja": "ワイドシルエット メンズ | オーバーサイズ メンズ | ジャストサイズ メンズ | バギー メンズ | フレアパンツ メンズ",
        "keywords_zh_cn": "宽松廓形 | 超大版 | 合体版 | 垮版 | 喇叭裤",
        "how_to_read": "这一步决定裤装、外套和上衣的版型开发方向。",
    },
    {
        "step": 7,
        "batch_name": "Scene demand",
        "goal_zh_cn": "判断用户是为通勤买，还是为周末和社交买。",
        "keywords_ja": "通勤 メンズ | オフィスカジュアル メンズ | 休日コーデ メンズ | ゴルフウェア メンズ | セットアップ メンズ",
        "keywords_zh_cn": "通勤 | 办公室休闲 | 周末穿搭 | 高尔夫男装 | 套装",
        "how_to_read": "场景词更接近消费需求，适合判断实际卖点。",
    },
    {
        "step": 8,
        "batch_name": "Age and function",
        "goal_zh_cn": "判断成熟用户和功能诉求是否能带来更稳的销量。",
        "keywords_ja": "30代 メンズ ファッション | 40代 メンズ ファッション | 50代 メンズ ファッション | 接触冷感 メンズ | 撥水 メンズ",
        "keywords_zh_cn": "30代男装 | 40代男装 | 50代男装 | 凉感男装 | 防泼水男装",
        "how_to_read": "如果年龄词稳定、功能词季节性强，说明可以用常青款加季节款组合。",
    },
]

SEARCH_MODIFIERS = [
    {
        "modifier_type": "Gender or target",
        "append_ja": "メンズ",
        "append_zh_cn": "男装",
        "use_case": "作为基础后缀，统一限定目标人群。",
    },
    {
        "modifier_type": "Outfit intent",
        "append_ja": "コーデ",
        "append_zh_cn": "穿搭",
        "use_case": "判断内容型需求和搭配关注度。",
    },
    {
        "modifier_type": "Recommendation",
        "append_ja": "おすすめ",
        "append_zh_cn": "推荐",
        "use_case": "判断用户是否已经进入选择阶段。",
    },
    {
        "modifier_type": "Popularity",
        "append_ja": "人気",
        "append_zh_cn": "热门",
        "use_case": "看大众化消费和高热产品方向。",
    },
    {
        "modifier_type": "Brand search",
        "append_ja": "ブランド",
        "append_zh_cn": "品牌",
        "use_case": "判断用户是否在找品牌而不是泛需求。",
    },
    {
        "modifier_type": "Season",
        "append_ja": "春",
        "append_zh_cn": "春季",
        "use_case": "看春季需求和换季节点。",
    },
    {
        "modifier_type": "Season",
        "append_ja": "夏",
        "append_zh_cn": "夏季",
        "use_case": "看夏季需求和凉感、轻薄卖点。",
    },
    {
        "modifier_type": "Season",
        "append_ja": "秋",
        "append_zh_cn": "秋季",
        "use_case": "看秋季过渡款和针织、轻外套需求。",
    },
    {
        "modifier_type": "Season",
        "append_ja": "冬",
        "append_zh_cn": "冬季",
        "use_case": "看保暖、层搭和外套需求。",
    },
    {
        "modifier_type": "Scene",
        "append_ja": "通勤",
        "append_zh_cn": "通勤",
        "use_case": "看办公室和城市日常穿着需求。",
    },
    {
        "modifier_type": "Scene",
        "append_ja": "休日",
        "append_zh_cn": "休闲日",
        "use_case": "看周末 casual 和社交穿着需求。",
    },
    {
        "modifier_type": "Pain point",
        "append_ja": "着回し",
        "append_zh_cn": "一衣多搭",
        "use_case": "判断实用主义和单品复用需求。",
    },
]

SHEET_MAP = {
    "style_web_5y": ("Style Summary", "Style Raw", "Style Related", "Style Chart"),
    "item_web_5y": ("Item Summary", "Item Raw", "Item Related", "Item Chart"),
    "item_shopping_12m": ("Shop Summary", "Shop Raw", "Shop Related", "Shop Chart"),
}


def autosize(worksheet, dataframe: pd.DataFrame) -> None:
    for idx, col in enumerate(dataframe.columns):
        series = dataframe[col].astype(str)
        max_len = max([len(str(col))] + series.map(len).tolist())
        worksheet.set_column(idx, idx, min(max_len + 2, 42))


def keyword_map_frame() -> pd.DataFrame:
    rows = [
        {"keyword_ja": keyword, "keyword_en": en, "keyword_zh_cn": zh}
        for keyword, (en, zh) in KEYWORD_TRANSLATIONS.items()
    ]
    return pd.DataFrame(rows).sort_values("keyword_ja")


def header_map_frame(slug: str, raw_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for column in raw_df.columns:
        if column == "date":
            rows.append(
                {
                    "dataset": slug,
                    "header_original": "date",
                    "header_en": "Date",
                    "header_zh_cn": "日期",
                }
            )
            continue
        if column == "isPartial":
            rows.append(
                {
                    "dataset": slug,
                    "header_original": "isPartial",
                    "header_en": "Partial data flag",
                    "header_zh_cn": "是否为未完结周数据",
                }
            )
            continue
        en, zh = KEYWORD_TRANSLATIONS.get(column, ("", ""))
        rows.append(
            {
                "dataset": slug,
                "header_original": column,
                "header_en": en,
                "header_zh_cn": zh,
            }
        )
    return pd.DataFrame(rows)


def related_frame(related_json: dict[str, list[dict[str, object]]]) -> pd.DataFrame:
    rows = []
    for keyword, values in related_json.items():
        keyword_en, keyword_zh = KEYWORD_TRANSLATIONS.get(keyword, ("", ""))
        if not values:
            rows.append(
                {
                    "keyword_ja": keyword,
                    "keyword_en": keyword_en,
                    "keyword_zh_cn": keyword_zh,
                    "query_ja": "",
                    "query_zh_cn": "",
                    "value": "",
                }
            )
            continue
        for entry in values:
            query = entry.get("query", "")
            rows.append(
                {
                    "keyword_ja": keyword,
                    "keyword_en": keyword_en,
                    "keyword_zh_cn": keyword_zh,
                    "query_ja": query,
                    "query_zh_cn": RELATED_QUERY_TRANSLATIONS.get(query, ""),
                    "value": entry.get("value", ""),
                }
            )
    return pd.DataFrame(rows)


def summary_frame(summary_df: pd.DataFrame) -> pd.DataFrame:
    enriched = summary_df.rename(columns={"keyword": "keyword_ja"}).copy()
    enriched.insert(
        1,
        "keyword_en",
        enriched["keyword_ja"].map(lambda value: KEYWORD_TRANSLATIONS.get(value, ("", ""))[0]),
    )
    enriched.insert(
        2,
        "keyword_zh_cn",
        enriched["keyword_ja"].map(lambda value: KEYWORD_TRANSLATIONS.get(value, ("", ""))[1]),
    )
    return enriched


def write_df(writer, sheet_name: str, dataframe: pd.DataFrame, header_fmt) -> None:
    dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    for col_num, value in enumerate(dataframe.columns.values):
        worksheet.write(0, col_num, value, header_fmt)
    autosize(worksheet, dataframe)
    worksheet.freeze_panes(1, 0)


def main() -> None:
    manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))

    with pd.ExcelWriter(OUTPUT_PATH, engine="xlsxwriter") as writer:
        workbook = writer.book
        title_fmt = workbook.add_format({"bold": True, "font_size": 16})
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
        wrap_fmt = workbook.add_format({"text_wrap": True, "valign": "top"})
        link_fmt = workbook.add_format({"font_color": "blue", "underline": 1})

        overview = workbook.add_worksheet("Overview")
        writer.sheets["Overview"] = overview
        overview.write("A1", "Japan Menswear Google Trends Analysis", title_fmt)
        overview.write("A3", "Export date")
        overview.write("B3", "2026-03-18")
        overview.write("A4", "Region")
        overview.write("B4", "Japan")
        overview.write("A5", "Workbook scope")
        overview.write("B5", "Existing data plus expanded style keyword framework")
        overview.write("A7", "Included datasets", header_fmt)
        overview.write_row(
            "A8",
            ["Slug", "Summary sheet", "Raw sheet", "Related sheet", "Chart sheet", "Google Trends link"],
            header_fmt,
        )

        row = 8
        header_rows = []
        for item in manifest:
            slug = item["slug"]
            summary_sheet, raw_sheet, related_sheet, chart_sheet = SHEET_MAP[slug]
            overview.write_row(row, 0, [slug, summary_sheet, raw_sheet, related_sheet, chart_sheet])
            overview.write_url(row, 5, item["link"], link_fmt, "Open in Google Trends")
            row += 1

            raw_df = pd.read_csv(BASE_DIR / f"{slug}.csv")
            header_rows.append(header_map_frame(slug, raw_df))

        overview.write("A14", "Notes", header_fmt)
        notes = [
            "Trend values are normalized 0-100 within each chart.",
            "Do not compare absolute demand across different sheets directly.",
            "Style words alone are not enough. Use style, silhouette, scene, age, and function together.",
            "Shopping Search is sparse and should be treated as a supporting signal only.",
        ]
        for idx, note in enumerate(notes, start=15):
            overview.write(f"A{idx}", note, wrap_fmt)
        overview.set_column("A:A", 22)
        overview.set_column("B:E", 22)
        overview.set_column("F:F", 28)

        write_df(writer, "Keyword Map", keyword_map_frame(), header_fmt)
        write_df(writer, "Header Map", pd.concat(header_rows, ignore_index=True), header_fmt)

        related_query_rows = [
            {"query_ja": query_ja, "query_zh_cn": query_zh}
            for query_ja, query_zh in RELATED_QUERY_TRANSLATIONS.items()
        ]
        write_df(writer, "Related Query Map", pd.DataFrame(related_query_rows), header_fmt)
        write_df(writer, "Style Universe", pd.DataFrame(STYLE_UNIVERSE), header_fmt)
        write_df(writer, "Trend Plan", pd.DataFrame(TREND_PLAN), header_fmt)
        write_df(writer, "Search Modifiers", pd.DataFrame(SEARCH_MODIFIERS), header_fmt)

        for item in manifest:
            slug = item["slug"]
            summary_sheet, raw_sheet, related_sheet, chart_sheet = SHEET_MAP[slug]

            summary_df = pd.read_csv(BASE_DIR / f"{slug}_summary.csv")
            raw_df = pd.read_csv(BASE_DIR / f"{slug}.csv")
            related_json = json.loads((BASE_DIR / f"{slug}_related.json").read_text(encoding="utf-8"))

            write_df(writer, summary_sheet, summary_frame(summary_df), header_fmt)
            write_df(writer, raw_sheet, raw_df, header_fmt)
            write_df(writer, related_sheet, related_frame(related_json), header_fmt)

            chart_ws = workbook.add_worksheet(chart_sheet)
            writer.sheets[chart_sheet] = chart_ws
            chart_ws.write("A1", slug, title_fmt)
            chart_ws.write("A3", "Chart labels can be cross-checked in 'Keyword Map' and 'Style Universe'.")
            chart_ws.write_url("A4", item["link"], link_fmt, "Open in Google Trends")
            chart_ws.insert_image("A6", str(BASE_DIR / f"{slug}.png"), {"x_scale": 0.75, "y_scale": 0.75})
            chart_ws.set_column("A:A", 58)

    print(OUTPUT_PATH)


if __name__ == "__main__":
    main()
