from __future__ import annotations

from pathlib import Path

import pandas as pd


EXPORT_DATE = "2026-03-19"

BASE_DIR = Path("april_sell_output")
OUT_DIR = Path("step2_output")
OUT_DIR.mkdir(exist_ok=True)

STYLE_PATH = BASE_DIR / "四月风格机会表.csv"
ITEM_PATH = BASE_DIR / "四月商品机会表.csv"
SCENE_PATH = BASE_DIR / "四月场景机会表.csv"
FUNCTION_PATH = BASE_DIR / "四月功能机会表.csv"

OUTPUT_XLSX = OUT_DIR / f"日本男装第二步_早期趋势验证_{EXPORT_DATE}.xlsx"
OUTPUT_MD = OUT_DIR / "第二步_商业分析.md"
OUTPUT_CSV = OUT_DIR / "第二步_机会矩阵.csv"


def load_csv(path: Path) -> pd.DataFrame:
    return pd.read_csv(path, encoding="utf-8-sig")


def pick_row(df: pd.DataFrame, keyword_zh: str) -> pd.Series:
    matched = df.loc[df["中文关键词"] == keyword_zh]
    if matched.empty:
        raise KeyError(f"missing keyword: {keyword_zh}")
    return matched.iloc[0]


def signal_rows() -> list[dict[str, str]]:
    return [
        {
            "来源类型": "WEAR 数据",
            "来源名称": "春季外套趋势",
            "发布时间": "2026-03-10",
            "信号分类": "外套主流",
            "对应方向": "男士夹克",
            "关键信号": "男装春季外套 Top3 是夹克、休闲西装、牛仔夹克。",
            "数据或证据": "WEAR 明确写明男装春季外套前 3 名为ブルゾン、テーラードジャケット、デニムジャケット。",
            "商业含义": "四月外套不该做抽象的轻外套，而应做可落地的短夹克、休闲西装、牛仔夹克。",
            "链接": "https://wear.jp/article/entry/2026/03/10/120000/",
        },
        {
            "来源类型": "WEAR 数据",
            "来源名称": "春季外套趋势",
            "发布时间": "2026-03-10",
            "信号分类": "搭配规则",
            "对应方向": "男士西裤",
            "关键信号": "夹克最推荐搭配宽直筒西裤。",
            "数据或证据": "文章写明夹克尤其适合搭配ワイドストレートシルエットのスラックス。",
            "商业含义": "四月的夹克不应和修身裤绑定，应该和宽直筒西裤做成一组卖法。",
            "链接": "https://wear.jp/article/entry/2026/03/10/120000/",
        },
        {
            "来源类型": "WEAR 数据",
            "来源名称": "春季外套趋势",
            "发布时间": "2026-03-10",
            "信号分类": "搭配规则",
            "对应方向": "男士牛仔裤",
            "关键信号": "休闲西装推荐搭配宽直筒牛仔裤做轻松通勤。",
            "数据或证据": "文章写明テーラードジャケット推荐搭配デニムパンツ，并强调ワイドストレート和ワンクッション。",
            "商业含义": "四月的清爽通勤不是全正装，而是西装外套加宽牛仔的轻正式。",
            "链接": "https://wear.jp/article/entry/2026/03/10/120000/",
        },
        {
            "来源类型": "WEAR 数据",
            "来源名称": "春季外套趋势",
            "发布时间": "2026-03-10",
            "信号分类": "搭配规则",
            "对应方向": "男士套装",
            "关键信号": "牛仔夹克和牛仔裤的同色成套穿法很强。",
            "数据或证据": "文章写明牛仔夹克最受欢迎的搭配之一是同色牛仔裤的套装式穿法。",
            "商业含义": "四月套装不只限于西装套装，牛仔成套也是值得做的小胶囊方向。",
            "链接": "https://wear.jp/article/entry/2026/03/10/120000/",
        },
        {
            "来源类型": "WEAR 数据",
            "来源名称": "春季外套趋势",
            "发布时间": "2026-03-10",
            "信号分类": "机能外套",
            "对应方向": "防泼水男装",
            "关键信号": "尼龙夹克和军装夹克仍有稳定热度。",
            "数据或证据": "文章写明男装里ナイロンジャケット和ミリタリージャケット仍然根强い人気。",
            "商业含义": "四月可以做轻机能夹克，但更适合做小份额补充，不宜抢主线。",
            "链接": "https://wear.jp/article/entry/2026/03/10/120000/",
        },
        {
            "来源类型": "WEAR 数据",
            "来源名称": "大学生春季穿搭",
            "发布时间": "2026-02-26",
            "信号分类": "上装占比",
            "对应方向": "男士衬衫",
            "关键信号": "大学生春季上装里衬衫占 36.1%，排第一。",
            "数据或证据": "WEAR 写明春季大学生上装 Top2 是衬衫 36.1%、T 恤 32.8%。",
            "商业含义": "四月衬衫不是辅助款，而是年轻客群里最直接的成交上装。",
            "链接": "https://wear.jp/article/entry/2026/02/26/120000/",
        },
        {
            "来源类型": "WEAR 数据",
            "来源名称": "大学生春季穿搭",
            "发布时间": "2026-02-26",
            "信号分类": "下装占比",
            "对应方向": "男士牛仔裤",
            "关键信号": "大学生春季下装里牛仔裤占 40.8%，明显第一。",
            "数据或证据": "WEAR 写明春季大学生下装 Top3 是牛仔裤 40.8%、工装裤 17.4%、西裤 15.2%。",
            "商业含义": "年轻人四月买裤子最稳的是牛仔，其次才是工装和西裤。",
            "链接": "https://wear.jp/article/entry/2026/02/26/120000/",
        },
        {
            "来源类型": "WEAR 数据",
            "来源名称": "大学生春季穿搭",
            "发布时间": "2026-02-26",
            "信号分类": "下装占比",
            "对应方向": "男士工装裤",
            "关键信号": "工装裤仍有存在感，但只占 17.4%。",
            "数据或证据": "WEAR 写明工装裤是大学生春季下装第 2 名，但占比明显低于牛仔裤。",
            "商业含义": "工装裤可以保留，但更像点缀而非四月主推。",
            "链接": "https://wear.jp/article/entry/2026/02/26/120000/",
        },
        {
            "来源类型": "WEAR 数据",
            "来源名称": "大学生春季穿搭",
            "发布时间": "2026-02-26",
            "信号分类": "下装占比",
            "对应方向": "男士西裤",
            "关键信号": "大学生春季下装里西裤占 15.2%，已经进入前三。",
            "数据或证据": "WEAR 写明西裤在大学生春季下装里排名第 3。",
            "商业含义": "四月西裤不只是上班人群需求，年轻男装也可以做。",
            "链接": "https://wear.jp/article/entry/2026/02/26/120000/",
        },
        {
            "来源类型": "WEAR 数据",
            "来源名称": "大学生春季穿搭",
            "发布时间": "2026-02-26",
            "信号分类": "外套占比",
            "对应方向": "男士休闲西装",
            "关键信号": "大学生春季外套里休闲西装占 34.8%，排第一。",
            "数据或证据": "WEAR 写明春季大学生外套 Top2 是休闲西装 34.8%、夹克 18.7%。",
            "商业含义": "四月轻正式明显成立，适合做年轻感休闲西装或套装。",
            "链接": "https://wear.jp/article/entry/2026/02/26/120000/",
        },
        {
            "来源类型": "WEAR 数据",
            "来源名称": "美式休闲春季趋势",
            "发布时间": "2026-02-24",
            "信号分类": "风格基调",
            "对应方向": "美式休闲男装",
            "关键信号": "美式休闲建议建立在清爽休闲基础上，不是重工复古。",
            "数据或证据": "文章明确建议以きれいめカジュアル为土台做春季美式休闲。",
            "商业含义": "你如果做美式，不要太老炮或太复古，应该偏清爽、都市、易穿。",
            "链接": "https://wear.jp/article/entry/2026/02/20/120000/",
        },
        {
            "来源类型": "WEAR 数据",
            "来源名称": "美式休闲春季趋势",
            "发布时间": "2026-02-24",
            "信号分类": "上装占比",
            "对应方向": "男士衬衫",
            "关键信号": "美式休闲里 T 恤约 41%，衬衫约 28%。",
            "数据或证据": "WEAR 写明 #アメカジ 投稿中 T 恤约 41%，衬衫约 28%。",
            "商业含义": "四月做美式风格时，T 恤是基底，衬衫是提升清爽感和层次的关键。",
            "链接": "https://wear.jp/article/entry/2026/02/20/120000/",
        },
        {
            "来源类型": "WEAR 数据",
            "来源名称": "美式休闲春季趋势",
            "发布时间": "2026-02-24",
            "信号分类": "下装占比",
            "对应方向": "男士牛仔裤",
            "关键信号": "美式休闲里牛仔裤约 48%，远高于卡其裤约 11%。",
            "数据或证据": "WEAR 写明 #アメカジ 投稿中牛仔裤约 48%，卡其裤约 11%。",
            "商业含义": "四月美式线最该押的是牛仔，不是卡其裤。",
            "链接": "https://wear.jp/article/entry/2026/02/20/120000/",
        },
        {
            "来源类型": "WEAR 数据",
            "来源名称": "美式休闲春季趋势",
            "发布时间": "2026-02-24",
            "信号分类": "版型规则",
            "对应方向": "男士阔腿裤",
            "关键信号": "美式休闲明确推荐 A 线或加宽 I 线，且下半身用深色压重心。",
            "数据或证据": "WEAR 写明当前春季美式休闲宜用 A ライン、太めの I ライン、下重心。",
            "商业含义": "四月裤型应继续偏宽直筒、阔腿、直落，不要回修身。",
            "链接": "https://wear.jp/article/entry/2026/02/20/120000/",
        },
        {
            "来源类型": "ZOZOTOWN 排名",
            "来源名称": "男士西裤排名",
            "发布时间": "抓取于 2026-03-19",
            "信号分类": "商品结构",
            "对应方向": "男士西裤",
            "关键信号": "热卖西裤集中在亚麻混纺、3D 打褶、双褶、宽松垂坠、可成套。",
            "数据或证据": "排名摘要中出现 Linen Blend 3D Tuck Slacks、Classic Double Tuck Wide Slacks、TR Stretch Loose 系列等。",
            "商业含义": "四月西裤不能只做普通通勤裤，应该做轻面料、双褶、宽松、可成套版本。",
            "链接": "https://zozo.jp/ranking/category/pants/slacks/all-sales-men.html",
        },
        {
            "来源类型": "ZOZOTOWN 排名",
            "来源名称": "男士衬衫排名",
            "发布时间": "抓取于 2026-03-19",
            "信号分类": "商品结构",
            "对应方向": "男士衬衫",
            "关键信号": "热卖衬衫集中在常规领、短款、条纹、宽松、易打理和功能面料。",
            "数据或证据": "排名摘要中出现 Regular Collar、Cropped Shirt、Stripe Shirt、Easy Care、UV Cut、Quick Dry。",
            "商业含义": "四月衬衫不应只做基础白衬衫，应补充条纹、蓝色、短款和易护理版本。",
            "链接": "https://zozo.jp/ranking/category/tops/shirt-blouse/all-sales-men.html",
        },
        {
            "来源类型": "ZOZOTOWN 排名",
            "来源名称": "男士休闲西装排名",
            "发布时间": "抓取于 2026-03-19",
            "信号分类": "商品结构",
            "对应方向": "男士休闲西装",
            "关键信号": "热卖休闲西装集中在双排扣、宽松、旅行感、全天候、可成套。",
            "数据或证据": "排名摘要中出现 loose basic double tailored jacket、travel setup、all weather active wear、setup compatible。",
            "商业含义": "四月的西装外套要更轻、更松、更可日常穿，而不是传统商务正装。",
            "链接": "https://zozo.jp/ranking/category/jacket-outerwear/tailored-jacket/all-sales-men.html",
        },
        {
            "来源类型": "ZOZOTOWN 排名",
            "来源名称": "男士套装排名",
            "发布时间": "抓取于 2026-03-19",
            "信号分类": "商品结构",
            "对应方向": "男士套装",
            "关键信号": "套装排名里同时存在运动套装、意式领套装、温控面料套装。",
            "数据或证据": "排名摘要出现 Nike Poly-Knit Tracksuit、Italian Collar Set-up、Outlast Basic Jacket ｜SETUP可。",
            "商业含义": "四月套装需求很广，但你的品牌更适合切清爽通勤和轻功能，不要往纯运动走。",
            "链接": "https://zozo.jp/ranking/category/jacket-outerwear/setup/all-sales-men.html",
        },
        {
            "来源类型": "ZOZOTOWN 排名",
            "来源名称": "男士牛仔裤排名",
            "发布时间": "抓取于 2026-03-19",
            "信号分类": "版型信号",
            "对应方向": "男士牛仔裤",
            "关键信号": "牛仔裤热卖版型集中在宽、弯、桶、喇叭、巴吉。",
            "数据或证据": "排名摘要出现 Wide Baggy、Curve、Barrel Leg、Flare、Balloon、Wide Straight 等多个版型词。",
            "商业含义": "四月牛仔绝对不该回修身，应继续往宽直筒和弧线轮廓做。",
            "链接": "https://zozo.jp/ranking/category/pants/denim-pants/all-sales-men.html",
        },
        {
            "来源类型": "ZOZOTOWN 排名",
            "来源名称": "男士牛仔夹克排名",
            "发布时间": "抓取于 2026-03-19",
            "信号分类": "版型信号",
            "对应方向": "男士牛仔夹克",
            "关键信号": "热卖牛仔夹克集中在短款、Trucker、工装感和可做牛仔成套。",
            "数据或证据": "排名摘要出现 Type III Trucker、Short Jacket、Work Jacket、デニムセットアップ推奨。",
            "商业含义": "四月牛仔夹克更适合做短款和成套，而不是长款或厚重复古版。",
            "链接": "https://zozo.jp/ranking/category/jacket-outerwear/denim-jacket/all-sales-men.html",
        },
        {
            "来源类型": "ZOZOTOWN 排名",
            "来源名称": "男士夹克排名",
            "发布时间": "抓取于 2026-03-19",
            "信号分类": "外套风格",
            "对应方向": "男士夹克",
            "关键信号": "热卖夹克里既有消防扣、粗花呢，也有运动轨道夹克。",
            "数据或证据": "排名摘要出现 Fireman Blouson、Velour Track Jacket 等。",
            "商业含义": "四月夹克可以做材质变化和设计点，但核心仍应保持短、利落、易搭配。",
            "链接": "https://zozo.jp/ranking/category/jacket-outerwear/jacket/all-sales-men.html",
        },
        {
            "来源类型": "ZOZOTOWN 排名",
            "来源名称": "男士开衫排名",
            "发布时间": "抓取于 2026-03-19",
            "信号分类": "品类观察",
            "对应方向": "男士开衫",
            "关键信号": "开衫热卖多为毛海、毛茸、针织 POLO 或拉链针织。",
            "数据或证据": "排名摘要里主要是 Shaggy、Mohair、Knit Polo、Zip Cardigan。",
            "商业含义": "开衫在四月适合做补充，但更偏材质款和风格款，不像衬衫那样是广谱刚需。",
            "链接": "https://zozo.jp/ranking/category/tops/cardigan/all-sales-men.html",
        },
        {
            "来源类型": "ZOZOTOWN 排名",
            "来源名称": "男士卡其裤排名",
            "发布时间": "抓取于 2026-03-19",
            "信号分类": "版型信号",
            "对应方向": "男士卡其裤",
            "关键信号": "卡其裤热卖集中在双褶、半喇叭、气球和宽直筒。",
            "数据或证据": "排名摘要出现 One Tuck Semi Flare、Wide Chino、Balloon Pants、2 Tuck Wide Pants。",
            "商业含义": "如果做卡其线，也应该用宽松和打褶做法，避免传统 slim chino。",
            "链接": "https://zozo.jp/ranking/category/pants/chino-pants/all-sales-men.html",
        },
        {
            "来源类型": "ZOZOTOWN 排名",
            "来源名称": "功能西裤排名",
            "发布时间": "抓取于 2026-03-19",
            "信号分类": "功能卖点",
            "对应方向": "防泼水男装",
            "关键信号": "品牌排名里已出现弹力、防泼水、可成套的西裤。",
            "数据或证据": "NANO universe 排名摘要出现 TEXBRID Twill Pants / Stretch / Water-repellent / Setup-compatible。",
            "商业含义": "功能不是文案点缀，而是真正在卖的卖点，适合放进四月通勤裤和套装。",
            "链接": "https://zozo.jp/ranking/brand/nanouniverse/pants/slacks/all-sales-men.html?p_gttagid=5500_41106",
        },
        {
            "来源类型": "ZOZOTOWN 排名",
            "来源名称": "商务衬衫排名",
            "发布时间": "抓取于 2026-03-19",
            "信号分类": "功能卖点",
            "对应方向": "可机洗男装",
            "关键信号": "商务衬衫排名里高频出现免烫、易护理、条纹和宽领。",
            "数据或证据": "SHIPS Colors、COMME CA ISM 等排名摘要出现 No-iron、Easy-care、Striped、Wide Collar。",
            "商业含义": "四月衬衫除了版型和颜色，必须补足免烫、易护理、吸湿快干这些易成交卖点。",
            "链接": "https://zozo.jp/ranking/brand/shipscolors/tops/business-shirt/all-sales-men.html",
        },
    ]


def build_opportunity_matrix(
    style_df: pd.DataFrame,
    item_df: pd.DataFrame,
    function_df: pd.DataFrame,
) -> pd.DataFrame:
    source_support = {
        "办公室休闲男装": 4,
        "美式休闲男装": 5,
        "简约男装": 2,
        "时装感男装": 2,
        "街头风男装": 1,
        "古着男装": 1,
        "男士套装": 4,
        "男士牛仔裤": 7,
        "男士衬衫": 6,
        "男士西裤": 6,
        "男士阔腿裤": 4,
        "男士夹克": 4,
        "男士开衫": 2,
        "男士工装裤": 2,
        "弹力男装": 2,
        "防泼水男装": 2,
        "可机洗男装": 2,
    }

    rows = []
    targets = [
        ("风格", pick_row(style_df, "办公室休闲男装"), "主线", "做四月品牌主线，承担通勤和清爽需求。"),
        ("风格", pick_row(style_df, "美式休闲男装"), "主线", "做第二主线，偏牛仔与清爽美式，不宜做厚重复古。"),
        ("风格", pick_row(style_df, "简约男装"), "辅助", "做视觉语言和配色方向，不必独立成风格线。"),
        ("风格", pick_row(style_df, "时装感男装"), "辅助", "做版型和设计点缀，不做太概念化。"),
        ("风格", pick_row(style_df, "街头风男装"), "谨慎", "保留少量街头感单品，但不建议做主轴。"),
        ("风格", pick_row(style_df, "古着男装"), "谨慎", "可借元素，不建议整线押注。"),
        ("商品", pick_row(item_df, "男士套装"), "主推", "做轻正式与轻机能两条套装线，控制运动感。"),
        ("商品", pick_row(item_df, "男士牛仔裤"), "主推", "做宽直筒、弧形或巴吉牛仔，并支持成套穿法。"),
        ("商品", pick_row(item_df, "男士衬衫"), "主推", "重点放在蓝白条、白衬衫、常规领、易打理版本。"),
        ("商品", pick_row(item_df, "男士西裤"), "主推", "做双褶、垂坠、亚麻混纺、可成套西裤。"),
        ("商品", pick_row(item_df, "男士阔腿裤"), "稳卖", "作为所有外套和衬衫的底盘裤型。"),
        ("商品", pick_row(item_df, "男士夹克"), "稳卖", "优先做短夹克或轻夹克，不做厚重长款。"),
        ("商品", pick_row(item_df, "男士开衫"), "补充", "做少量针织和衬衫内搭线，不做主力。"),
        ("商品", pick_row(item_df, "男士工装裤"), "谨慎", "只做小份额点缀，避免占库存。"),
        ("功能", pick_row(function_df, "弹力男装"), "主推", "主力写进裤装、套装和通勤外套。"),
        ("功能", pick_row(function_df, "防泼水男装"), "稳卖", "放进轻外套和轻机能套装最合理。"),
        ("功能", pick_row(function_df, "可机洗男装"), "稳卖", "放进衬衫和通勤裤，提升转化。"),
    ]

    for category, row, strategy, action in targets:
        keyword = row["中文关键词"]
        support = source_support.get(keyword, 0)
        trend_yoy = float(row["同比变化（%）"]) if str(row["同比变化（%）"]).strip() else 0.0
        april_index = float(row["4月季节指数"]) if str(row["4月季节指数"]).strip() else 0.0
        score = round(support * 15 + max(trend_yoy, 0) * 0.8 + min(april_index, 180) * 0.2, 1)
        rows.append(
            {
                "分类": category,
                "中文关键词": keyword,
                "日文关键词": row["日文关键词"],
                "第一步判断": row["判断"],
                "近52周均值": row["近52周均值"],
                "同比变化（%）": row["同比变化（%）"],
                "4月季节指数": row["4月季节指数"],
                "第二步信号数": support,
                "综合机会分": score,
                "策略级别": strategy,
                "建议动作": action,
            }
        )

    result = pd.DataFrame(rows).sort_values(["策略级别", "综合机会分"], ascending=[True, False])
    return result


def build_overview() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"项目": "分析对象", "内容": "日本男装，四月销售窗口"},
            {"项目": "这一步做什么", "内容": "按照截图中可见的第二步“AI 挖早期趋势”思路，用内容平台和交易平台信号，验证第一步 Google Trends 的结论。"},
            {"项目": "使用来源", "内容": "WEAR 官方数据文章、ZOZOTOWN 排名页搜索摘要、第一步四月 Google Trends 数据。"},
            {"项目": "这一步解决的问题", "内容": "把“方向对不对”进一步缩到“具体卖什么、怎么卖、哪些只该少量做”。"},
            {"项目": "最重要发现", "内容": "四月更适合办公室休闲、轻美式、衬衫、牛仔、西裤、夹克、轻套装；不适合把重街头和工装裤当主力。"},
        ]
    )


def build_final_summary(matrix: pd.DataFrame) -> pd.DataFrame:
    top = matrix.loc[matrix["策略级别"].isin(["主线", "主推", "稳卖"])].copy()
    top = top.sort_values("综合机会分", ascending=False).head(10)
    return top[["分类", "中文关键词", "第一步判断", "第二步信号数", "综合机会分", "策略级别", "建议动作"]]


def write_markdown(
    matrix: pd.DataFrame,
    source_df: pd.DataFrame,
) -> None:
    top_core = matrix.loc[matrix["策略级别"].isin(["主线", "主推"])].sort_values("综合机会分", ascending=False)
    cautious = matrix.loc[matrix["策略级别"] == "谨慎"].sort_values("综合机会分", ascending=False)

    lines = [
        "# 日本男装第二步：早期趋势验证",
        "",
        f"- 分析日期：{EXPORT_DATE}",
        "- 目标：承接第一步 Google Trends，把四月卖货方向进一步缩到能执行的风格、商品和卖点。",
        "",
        "## 最终结论",
        "",
        "四月最值得做的不是重街头，而是“办公室休闲 + 轻美式 + 简约利落”。第二步的内容平台和交易平台信号，与第一步的 Google Trends 基本一致，而且把单品答案收得更清楚了：真正该押的是套装、牛仔裤、衬衫、西裤、短夹克；工装裤和重古着只适合做点缀。",
        "",
        "## 主推方向",
        "",
    ]

    for _, row in top_core.iterrows():
        lines.append(
            f"- {row['中文关键词']}：第一步是“{row['第一步判断']}”，第二步拿到 {int(row['第二步信号数'])} 个支持信号，综合机会分 {row['综合机会分']}。{row['建议动作']}"
        )

    lines.extend(
        [
            "",
            "## 谨慎方向",
            "",
        ]
    )

    for _, row in cautious.iterrows():
        lines.append(
            f"- {row['中文关键词']}：第一步虽有一定季节性，但第二步支持信号弱，且容易偏离四月主流成交逻辑。{row['建议动作']}"
        )

    lines.extend(
        [
            "",
            "## 第二步带来的新增判断",
            "",
            "- 外套要从抽象的“轻外套”落到具体的“短夹克、休闲西装、牛仔夹克”。",
            "- 衬衫不能只做基础白衬衫，四月更适合蓝白条、常规领、短款、易护理、快干版本。",
            "- 裤型应继续押宽直筒、阔腿、双褶、弧线轮廓，不能回到修身裤。",
            "- 套装是强需求，但更适合做轻正式和轻机能，不适合做纯运动套装。",
            "- 功能卖点里最能转化的是弹力、防泼水、免烫和易护理；凉感还太早。",
            "",
            "## 主要来源",
            "",
        ]
    )

    for _, row in source_df.drop_duplicates(subset=["来源名称", "链接"]).iterrows():
        lines.append(f"- {row['来源名称']}：{row['链接']}")

    OUTPUT_MD.write_text("\n".join(lines), encoding="utf-8")


def autosize(worksheet, dataframe: pd.DataFrame) -> None:
    for idx, col in enumerate(dataframe.columns):
        values = dataframe[col].astype(str).tolist()
        max_len = max([len(str(col))] + [len(value) for value in values])
        worksheet.set_column(idx, idx, min(max_len + 2, 60))


def write_sheet(writer: pd.ExcelWriter, name: str, df: pd.DataFrame, link_column: str | None = None) -> None:
    df.to_excel(writer, sheet_name=name, index=False)
    ws = writer.sheets[name]
    autosize(ws, df)
    ws.freeze_panes(1, 0)

    if link_column and link_column in df.columns:
        link_idx = list(df.columns).index(link_column)
        link_fmt = writer.book.add_format({"font_color": "blue", "underline": 1})
        for row_idx, url in enumerate(df[link_column], start=1):
            if isinstance(url, str) and url.startswith("http"):
                ws.write_url(row_idx, link_idx, url, link_fmt, string=url)


def main() -> None:
    style_df = load_csv(STYLE_PATH)
    item_df = load_csv(ITEM_PATH)
    scene_df = load_csv(SCENE_PATH)
    function_df = load_csv(FUNCTION_PATH)

    source_df = pd.DataFrame(signal_rows())
    matrix_df = build_opportunity_matrix(style_df, item_df, function_df)
    summary_df = build_final_summary(matrix_df)
    overview_df = build_overview()

    scene_focus_df = scene_df.loc[
        scene_df["中文关键词"].isin(["办公室休闲男装", "通勤男装", "高尔夫男装", "春装男装"])
    ].copy()

    write_markdown(matrix_df, source_df)
    matrix_df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")

    with pd.ExcelWriter(OUTPUT_XLSX, engine="xlsxwriter") as writer:
        write_sheet(writer, "概览", overview_df)
        write_sheet(writer, "第二步来源信号", source_df, link_column="链接")
        write_sheet(writer, "第二步机会矩阵", matrix_df)
        write_sheet(writer, "第二步最终结论", summary_df)
        write_sheet(writer, "第一步场景参考", scene_focus_df)


if __name__ == "__main__":
    main()
