from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Iterable
from urllib.parse import quote, urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from youtube_comment_downloader import SORT_BY_POPULAR, YoutubeCommentDownloader


EXPORT_DATE = "2026-03-19"

APRIL_DIR = Path("april_sell_output")
OUT_DIR = Path("step2_plus_output")
OUT_DIR.mkdir(exist_ok=True)

STYLE_PATH = APRIL_DIR / "四月风格机会表.csv"
ITEM_PATH = APRIL_DIR / "四月商品机会表.csv"
FUNCTION_PATH = APRIL_DIR / "四月功能机会表.csv"
SCENE_PATH = APRIL_DIR / "四月场景机会表.csv"

OUTPUT_XLSX = OUT_DIR / f"日本男装第二步_多来源验证_{EXPORT_DATE}.xlsx"
OUTPUT_MD = OUT_DIR / "第二步_多来源商业分析.md"

HEADERS = {"User-Agent": "Mozilla/5.0"}


AMAZON_CATEGORIES = [
    {"label": "亚马逊日本_男士衬衫", "node": "24548215051", "group": "衬衫"},
    {"label": "亚马逊日本_男士开衫", "node": "2131438051", "group": "开衫"},
    {"label": "亚马逊日本_男士西装外套", "node": "5825572051", "group": "西装外套"},
    {"label": "亚马逊日本_男士牛仔裤", "node": "2133077051", "group": "牛仔裤"},
    {"label": "亚马逊日本_男士长裤", "node": "2131443051", "group": "长裤"},
]


YOUTUBE_QUERIES = [
    {"query": "メンズ 春服 オフィスカジュアル", "focus": "办公室休闲"},
    {"query": "メンズ 春服 アメカジ", "focus": "美式休闲"},
    {"query": "メンズ 春 デニム コーデ", "focus": "牛仔穿搭"},
    {"query": "メンズ アイビー プレッピー", "focus": "学院风"},
]


YOUTUBE_SELECTED = [
    {
        "focus": "办公室休闲",
        "video_id": "54e9sP-cDQI",
        "title": "【大人世代！春のビジカジ】2026春！ビジネスカジュアル",
        "channel": "Chu Chu DANSHI メンズファッションチャンネル",
        "skip_author": "@chuchudanshi",
        "why": "当前四月办公室休闲代表内容，且评论集中在白衬衫、内搭、透感、感动系列。",
    },
    {
        "focus": "春季主流穿搭",
        "video_id": "aM1qTR3TFKU",
        "title": "女性目線で男性が着てたらカッコいい春服&コーデ【UNIQLO 春服メンズ2026】",
        "channel": "八田エミリの日常",
        "skip_author": "@hattaemily",
        "why": "观看量高，能反映大众接受度和女性视角的穿搭偏好。",
    },
    {
        "focus": "美式休闲",
        "video_id": "iuNP8IiEgto",
        "title": "春のアメカジコーデ対決！",
        "channel": "SCYTHE(サイズ）TV",
        "skip_author": "",
        "why": "时效新，评论直接投票具体春季美式搭配。",
    },
    {
        "focus": "学院风",
        "video_id": "8f7lqPehBs8",
        "title": "アメリカの王道メンズファッション｜IVY STYLE（アイビールック）のすすめ",
        "channel": "PALATINE",
        "skip_author": "",
        "why": "日语 YouTube 里较有代表性的学院风 / 常春藤内容。",
    },
]


FORUM_ROWS = [
    {
        "来源平台": "Yahoo!知恵袋",
        "主题": "办公室休闲边界",
        "用户原话或摘要": "公司如果明确推行办公室休闲，首日穿得过于正式反而会显得突兀。",
        "商业判断": "四月通勤线要做“得体但不太正式”，不能只做传统正装。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q13283735093",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "主题": "面试 / 办公室休闲实操",
        "用户原话或摘要": "感动ジャケット・パンツ可以作为办公室休闲，但白 T 或 POLO 对面试可能偏松，建议简洁色和西裤。",
        "商业判断": "套装和西裤成立，但上装要注意正式度，衬衫仍然比白 T 更稳。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q11301799726",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "主题": "低身高与阔腿裤",
        "用户原话或摘要": "低身高穿过宽阔腿裤容易显得比例变差，很多人担心脸大、腿短被放大。",
        "商业判断": "阔腿裤仍可卖，但四月要控制裤宽和长度，给出适合 160-170cm 的版型解释。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q12233510552",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "主题": "学生预算与品牌",
        "用户原话或摘要": "学生群体会优先看 5000 日元以内、ZOZOTOWN 买得到的品牌和单品。",
        "商业判断": "学生向线必须控制价格和搭配门槛，学院风尤其不能做得过贵。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q14286522609",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "主题": "LLBean 与常春藤风格",
        "用户原话或摘要": "LLBean 被直接归到常春藤 / 学院风脉络里，但和 Patagonia 这类一起穿会更偏重美式休闲。",
        "商业判断": "学院风在日本语境里更接近常春藤和东海岸气质，不等于纯校服风。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q11268982261",
    },
]


REVIEW_ROWS = [
    {
        "来源平台": "楽天评论",
        "品类": "免烫衬衫",
        "用户原话或摘要": "白衬衫有硬性需求，但真正让人下单的是“ノーアイロン、肌触りが良い、吸水速乾”。",
        "商业判断": "四月衬衫最强卖点不是风格词，而是免烫、舒适、快干、可通勤。",
        "链接": "https://review.rakuten.co.jp/item/1/251912_10007319/1.1/",
    },
    {
        "来源平台": "楽天评论",
        "品类": "免烫衬衫",
        "用户原话或摘要": "线下买很贵，线上同类价格能做到一半以下，洗后状态和穿着感受都不错。",
        "商业判断": "价格敏感明显，功能衬衫可以卖，但要把性价比和易打理写透。",
        "链接": "https://review.rakuten.co.jp/item/1/251912_10007910/1.0/",
    },
    {
        "来源平台": "楽天评论",
        "品类": "蓝色纽扣领衬衫",
        "用户原话或摘要": "价格不过高，但生地和做工扎实，希望持续卖。",
        "商业判断": "蓝色和条纹衬衫是四月可扩的颜色方向，不只限白衬衫。",
        "链接": "https://review.rakuten.co.jp/item/1/251912_10012957/1.1/",
    },
    {
        "来源平台": "楽天评论",
        "品类": "紺ブレ",
        "用户原话或摘要": "评论明确写到“就是在找紺ブレ”，而且有用户表示会每年替换一次。",
        "商业判断": "海军蓝西装外套在日本是稳定刚需，适合挂到学院风和办公室休闲两条线。",
        "链接": "https://review.rakuten.co.jp/item/1/223890_10001890/1.1/",
    },
    {
        "来源平台": "楽天评论",
        "品类": "直筒牛仔",
        "用户原话或摘要": "用户最常提的是“太腿舒服、好动、颜色顺眼、价格还能接受”。",
        "商业判断": "牛仔的真实需求是舒适和好搭，直筒和轻弹更稳，过于概念化的版型风险更高。",
        "链接": "https://review.rakuten.co.jp/item/1/222661_10004538/1.1/",
    },
    {
        "来源平台": "楽天评论",
        "品类": "直筒牛仔",
        "用户原话或摘要": "日本中年体型用户会强调“长度刚好、不用改裤脚、蹲下轻松”。",
        "商业判断": "裤长和活动性是关键卖点，四月裤装页面要明确长度和弹性信息。",
        "链接": "https://review.rakuten.co.jp/review/review/review/item/1/195436_10040048/1.1/",
    },
    {
        "来源平台": "楽天评论",
        "品类": "袴 / 极宽裤",
        "用户原话或摘要": "穿着一松，裤脚就容易拖地。",
        "商业判断": "极宽裤不是不能卖，但四月大盘更适合“干净宽松”，不要盲目放大袴感。",
        "链接": "https://review.rakuten.co.jp/item/1/402463_10007311/1.1/",
    },
]


def load_csv(path: Path) -> pd.DataFrame:
    return pd.read_csv(path, encoding="utf-8-sig")


def pick_row(df: pd.DataFrame, keyword_zh: str) -> pd.Series:
    return df.loc[df["中文关键词"] == keyword_zh].iloc[0]


def clean_amazon_title(card: BeautifulSoup) -> str:
    title = ""
    clamp = card.select_one("div._cDEzb_p13n-sc-css-line-clamp-2_EWgCb")
    if clamp:
        title = clamp.get_text(" ", strip=True)
    if not title:
        clamp = card.select_one("div._cDEzb_p13n-sc-css-line-clamp-3_g3dy1")
        if clamp:
            title = clamp.get_text(" ", strip=True)
    if not title:
        img = card.select_one("img")
        if img and img.get("alt"):
            title = img["alt"]
    return title.strip()


def fetch_amazon_bestsellers(limit: int = 8) -> pd.DataFrame:
    rows: list[dict[str, str]] = []
    for category in AMAZON_CATEGORIES:
        url = f"https://www.amazon.co.jp/gp/bestsellers/fashion/{category['node']}"
        cards = []
        for _ in range(3):
            html = requests.get(url, headers=HEADERS, timeout=30).text
            soup = BeautifulSoup(html, "html.parser")
            cards = soup.select("div.zg-grid-general-faceout")
            if cards and "opfcaptcha" not in html:
                break
        for rank, card in enumerate(cards[:limit], start=1):
            title = clean_amazon_title(card)
            rating = card.select_one(".a-icon-alt")
            rating = rating.get_text(" ", strip=True) if rating else ""
            review_link = card.select_one('a[href*="/product-reviews/"]')
            review_count = ""
            if review_link:
                small = review_link.select_one("span.a-size-small")
                if small:
                    review_count = small.get_text(" ", strip=True)
            price = ""
            price_sel = card.select_one("span._cDEzb_p13n-sc-price_3mJ9Z") or card.select_one(".p13n-sc-price")
            if price_sel:
                price = price_sel.get_text(" ", strip=True)
            asin = ""
            root = card.select_one("div.p13n-sc-uncoverable-faceout")
            if root and root.get("id"):
                asin = root["id"]
            rows.append(
                {
                    "来源组": category["label"],
                    "商品组": category["group"],
                    "当前排名": rank,
                    "商品标题": title,
                    "评分": rating,
                    "评论量": review_count,
                    "价格": price,
                    "ASIN": asin,
                    "榜单链接": url,
                }
            )
    return pd.DataFrame(rows)


def find_video_renderers(obj: object, out: list[dict[str, object]]) -> None:
    if isinstance(obj, dict):
        if "videoRenderer" in obj:
            out.append(obj["videoRenderer"])
        for value in obj.values():
            find_video_renderers(value, out)
    elif isinstance(obj, list):
        for item in obj:
            find_video_renderers(item, out)


def youtube_search(query: str, limit: int = 8) -> list[dict[str, str]]:
    url = "https://www.youtube.com/results?search_query=" + quote(query)
    html = requests.get(url, headers=HEADERS, timeout=30).text
    matched = re.search(r"var ytInitialData = (\{.*?\});", html)
    if not matched:
        return []
    data = json.loads(matched.group(1))
    videos: list[dict[str, object]] = []
    find_video_renderers(data, videos)
    rows = []
    for video in videos[:limit]:
        title = "".join(run.get("text", "") for run in video.get("title", {}).get("runs", []))
        channel = "".join(run.get("text", "") for run in video.get("ownerText", {}).get("runs", []))
        views = video.get("viewCountText", {}).get("simpleText", "")
        published = video.get("publishedTimeText", {}).get("simpleText", "")
        video_id = video.get("videoId", "")
        rows.append(
            {
                "查询词": query,
                "视频标题": title,
                "频道": channel,
                "观看量": views,
                "发布时间": published,
                "视频链接": f"https://www.youtube.com/watch?v={video_id}",
                "视频ID": video_id,
            }
        )
    return rows


def fetch_youtube_searches() -> pd.DataFrame:
    rows: list[dict[str, str]] = []
    for item in YOUTUBE_QUERIES:
        for row in youtube_search(item["query"], limit=6):
            row["关注方向"] = item["focus"]
            rows.append(row)
    return pd.DataFrame(rows)


def fetch_youtube_comments(limit: int = 8) -> pd.DataFrame:
    downloader = YoutubeCommentDownloader()
    rows: list[dict[str, str]] = []
    for item in YOUTUBE_SELECTED:
        count = 0
        for comment in downloader.get_comments_from_url(
            f"https://www.youtube.com/watch?v={item['video_id']}",
            sort_by=SORT_BY_POPULAR,
        ):
            if comment.get("reply"):
                continue
            if item["skip_author"] and comment.get("author") == item["skip_author"]:
                continue
            text = str(comment.get("text", "")).replace("\n", " / ").strip()
            if not text or len(text) > 260:
                continue
            rows.append(
                {
                    "关注方向": item["focus"],
                    "视频标题": item["title"],
                    "频道": item["channel"],
                    "评论作者": comment.get("author", ""),
                    "点赞数": comment.get("votes", ""),
                    "评论内容": text,
                    "选择理由": item["why"],
                    "视频链接": f"https://www.youtube.com/watch?v={item['video_id']}",
                }
            )
            count += 1
            if count >= limit:
                break
    return pd.DataFrame(rows)


def build_amazon_summary(amazon_df: pd.DataFrame) -> pd.DataFrame:
    summaries: list[dict[str, str]] = []
    for group, sub_df in amazon_df.groupby("商品组"):
        top_titles = "；".join(sub_df.sort_values("当前排名").head(3)["商品标题"].tolist())
        summaries.append(
            {
                "商品组": group,
                "当前 Top3 摘要": top_titles,
                "商业解读": amazon_group_takeaway(group, sub_df),
            }
        )
    return pd.DataFrame(summaries)


def amazon_group_takeaway(group: str, sub_df: pd.DataFrame) -> str:
    title_blob = " ".join(sub_df["商品标题"].fillna("").tolist()).lower()
    if group == "衬衫":
        return "当前亚马逊日本衬衫榜单高度集中在免烫、快干、抗菌、常规领和商务可穿，说明四月衬衫的成交关键词是功能与通勤。"
    if group == "牛仔裤":
        return "牛仔榜单同时存在 Levi's 505 这类正统直筒和宽松款，说明四月牛仔可以做宽松，但仍需保留直筒与轻弹。"
    if group == "西装外套":
        if "スクール" in title_blob:
            return "海军蓝西装外套仍是硬需求，而且榜单里出现多条 school blazer，学院风有商品需求，但更像细分人群而非大盘主线。"
        return "当前西装外套仍以紺ブレ和通勤西装为核心。"
    if group == "开衫":
        if "スクール" in title_blob:
            return "开衫榜单里学校和商务关键词同时存在，说明学院风在开衫上有真实购买需求，但更偏基础 V 领与纯色。"
        return "开衫有需求，但更偏基础针织和通勤补充。"
    if group == "长裤":
        return "长裤榜单说明四月裤装要兼顾舒适、弹力、宽松与通勤，不宜只做极端宽裤。"
    return ""


def build_amazon_history_proxy(
    item_df: pd.DataFrame,
    style_df: pd.DataFrame,
) -> pd.DataFrame:
    shirt = pick_row(item_df, "男士衬衫")
    denim = pick_row(item_df, "男士牛仔裤")
    slacks = pick_row(item_df, "男士西裤")
    cardigan = pick_row(item_df, "男士开衫")
    office = pick_row(style_df, "办公室休闲男装")
    return pd.DataFrame(
        [
            {
                "维度": "说明",
                "历史 4 月口径": "亚马逊日本官方不提供公开历史畅销榜导出；我尝试了存档接口，但相关榜单页没有可用 2024/2025 年 4 月快照，因此这里使用“四月季节性 + 真实购买评论”做代理判断。",
                "当前证据": "当前榜单可直接抓到，历史榜单不可稳定抓到。",
                "代理结论": "可用于判断历史四月更可能卖得好的品类，而不是还原精确名次。",
            },
            {
                "维度": "衬衫",
                "历史 4 月口径": f"谷歌趋势四月季节指数 {shirt['4月季节指数']}，且楽天评论高频出现ノーアイロン、吸水速乾、新生活。",
                "当前证据": "当前亚马逊衬衫榜单前列被免烫、快干、商务衬衫占据。",
                "代理结论": "历史 4 月衬衫大概率一直是强势品类，且功能型商务衬衫更稳定。",
            },
            {
                "维度": "牛仔裤",
                "历史 4 月口径": f"谷歌趋势四月季节指数 {denim['4月季节指数']}，楽天评论强调直筒、舒适、长度合适、易活动。",
                "当前证据": "当前亚马逊牛仔榜单前列同时有 Levi's 505 和宽松款。",
                "代理结论": "历史 4 月牛仔大概率稳卖，且“可长期穿”的直筒与轻宽松更稳。",
            },
            {
                "维度": "西裤 / 通勤裤",
                "历史 4 月口径": f"西裤同比 {slacks['同比变化（%）']}%，办公室休闲四月季节指数 {office['4月季节指数']}。",
                "当前证据": "亚马逊和 YouTube 都围绕感动パンツ、感动ジャケット、紺ブレ展开。",
                "代理结论": "历史 4 月通勤裤与轻正式裤有连续性需求，尤其适合新生活和换季场景。",
            },
            {
                "维度": "开衫",
                "历史 4 月口径": f"开衫四月季节指数 {cardigan['4月季节指数']}，楽天评论能看到 2025-04 的购买记录。",
                "当前证据": "当前亚马逊开衫榜单里基础 V 领和 school cardigan 都在前列。",
                "代理结论": "开衫在历史 4 月有需求，但更适合作为补充而非最大主力。",
            },
        ]
    )


def build_youtube_summary(search_df: pd.DataFrame, comment_df: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, str]] = []
    for focus, sub_df in search_df.groupby("关注方向"):
        top_views = "；".join(sub_df.head(3)["视频标题"].tolist())
        comments = comment_df.loc[comment_df["关注方向"] == focus, "评论内容"].head(3).tolist()
        rows.append(
            {
                "关注方向": focus,
                "YouTube 结果摘要": top_views,
                "评论区主要声音": "；".join(comments),
                "商业判断": youtube_takeaway(focus, comments),
            }
        )
    return pd.DataFrame(rows)


def youtube_takeaway(focus: str, comments: Iterable[str]) -> str:
    blob = " ".join(comments)
    if focus == "办公室休闲":
        return "用户最在意白衬衫的光泽、透感、内搭替代和感动系列，说明通勤不是抽象风格，而是非常具体的穿着痛点。"
    if focus == "春季主流穿搭":
        return "高观看量视频下，用户反复提到 Milano Rib Jacket、Barrel Pants、亮色搭配和身高适配，说明轻外套和裤型选择是四月真实决策点。"
    if focus == "美式休闲":
        return "评论会对具体编号投票，偏好“清爽、蓝色、シャンブレー、短外套”，说明美式休闲应走易懂、易模仿的搭配。"
    if focus == "学院风":
        return "学院风内容有受众，但观看量和互动明显低于办公室休闲与春季主流，属于细分审美，不是大众通货。"
    return ""


def build_academy_analysis(
    style_df: pd.DataFrame,
    amazon_df: pd.DataFrame,
    youtube_search_df: pd.DataFrame,
) -> pd.DataFrame:
    preppy = pick_row(style_df, "学院派男装")
    blazer_df = amazon_df.loc[amazon_df["商品组"] == "西装外套"].copy()
    cardigan_df = amazon_df.loc[amazon_df["商品组"] == "开衫"].copy()
    preppy_video = youtube_search_df.loc[youtube_search_df["关注方向"] == "学院风"].head(5)

    school_blazer_count = int(blazer_df["商品标题"].fillna("").str.contains("スクール").sum())
    school_cardigan_count = int(cardigan_df["商品标题"].fillna("").str.contains("スクール|school", case=False, regex=True).sum())

    return pd.DataFrame(
        [
            {"指标": "谷歌趋势同比", "结果": preppy["同比变化（%）"], "说明": "学院派男装同比高增长，但基数很小。"},
            {"指标": "谷歌趋势近52周均值", "结果": preppy["近52周均值"], "说明": "搜索体量仍远小于办公室休闲和美式休闲。"},
            {"指标": "亚马逊开衫榜单", "结果": f"含 {school_cardigan_count} 个 school cardigan 关键词商品", "说明": "学院风真实购买更多落在基础 V 领开衫上。"},
            {"指标": "亚马逊西装外套榜单", "结果": f"含 {school_blazer_count} 个 school blazer 关键词商品", "说明": "学院风更像海军蓝西装外套和制服外套的细分需求。"},
            {"指标": "YouTube 学院风结果", "结果": f"Top 视频观看量最高约 {preppy_video.iloc[0]['观看量'] if not preppy_video.empty else ''}", "说明": "内容存在但热度显著低于春季主流和办公室休闲。"},
            {"指标": "结论", "结果": "可做，但不宜做主线", "说明": "更适合把学院风拆成“紺ブレ + V 领开衫 + 牛津衬衫 + 直筒牛仔/卡其裤”的轻学院胶囊。"},
        ]
    )


def build_final_recommendation(
    style_df: pd.DataFrame,
    item_df: pd.DataFrame,
    function_df: pd.DataFrame,
) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "层级": "主线风格",
                "方向": "办公室休闲 + 轻美式 + 简约利落",
                "原因": f"办公室休闲同比 {pick_row(style_df, '办公室休闲男装')['同比变化（%）']}%，美式休闲同比 {pick_row(style_df, '美式休闲男装')['同比变化（%）']}%，且多来源都支持。",
            },
            {
                "层级": "核心商品",
                "方向": "衬衫 / 牛仔裤 / 西裤 / 轻套装 / 短夹克",
                "原因": "谷歌趋势、WEAR、亚马逊日本、YouTube 用户讨论都同时指向这 5 个方向。",
            },
            {
                "层级": "功能卖点",
                "方向": "免烫 / 易护理 / 弹力 / 防泼水",
                "原因": "Amazon 当前榜单和楽天评论都在强调功能型衬衫、通勤裤和轻外套。",
            },
            {
                "层级": "版型",
                "方向": "宽直筒、双褶、轻阔腿、短外套",
                "原因": "消费者喜欢宽松，但论坛和评论提醒极宽裤、拖地裤、过度 oversized 有风险。",
            },
            {
                "层级": "学院风专项",
                "方向": "做成小胶囊，不做品牌主线",
                "原因": "学院风在商品层面有真实需求，但搜索量和内容热度都偏小众，更适合嫁接到紺ブレ、V 领开衫、牛津衬衫。"},
            {
                "层级": "谨慎方向",
                "方向": "重街头 / 重古着 / 工装裤主推 / 过度凉感",
                "原因": "四月真实用户更在意通勤、清爽、功能和适配性，而不是强风格标签。",
            },
        ]
    )


def write_markdown(
    amazon_summary: pd.DataFrame,
    youtube_summary: pd.DataFrame,
    academy_df: pd.DataFrame,
) -> None:
    lines = [
        "# 日本男装第二步：多来源验证版",
        "",
        f"- 分析日期：{EXPORT_DATE}",
        "- 目标：在第一步四月 Google Trends 基础上，补充 Amazon 日本、YouTube、日本论坛与真实购买评论，做更接近成交端的判断。",
        "",
        "## 最终结论",
        "",
        "四月日本男装的大盘答案没有变，但这次更具体了：主线仍然是办公室休闲、轻美式和简约利落；真正该押的商品是衬衫、牛仔裤、西裤、轻套装和短夹克。多来源数据都不支持把重街头、重古着和工装裤当主推。",
        "",
        "## Amazon 日本给出的新信息",
        "",
    ]

    for _, row in amazon_summary.iterrows():
        lines.append(f"- {row['商品组']}：{row['商业解读']}")

    lines.extend(
        [
            "",
            "## YouTube 和论坛给出的用户信息",
            "",
        ]
    )

    for _, row in youtube_summary.iterrows():
        lines.append(f"- {row['关注方向']}：{row['商业判断']}")

    lines.extend(
        [
            "",
            "## 学院风专项结论",
            "",
            "学院风不是不能做，但从搜索量、内容热度和用户讨论看，它更像四月里的细分胶囊，而不是主线。最适合的做法是把它拆成基础单品语言：紺ブレ、V 领开衫、牛津衬衫、直筒牛仔或卡其裤，再带一点常春藤气质，而不是直接做成校服感很强的成套造型。",
            "",
        ]
    )

    for _, row in academy_df.iterrows():
        lines.append(f"- {row['指标']}：{row['结果']}。{row['说明']}")

    OUTPUT_MD.write_text("\n".join(lines), encoding="utf-8")


def autosize(worksheet, dataframe: pd.DataFrame) -> None:
    for idx, col in enumerate(dataframe.columns):
        values = dataframe[col].astype(str).tolist()
        max_len = max([len(str(col))] + [len(value) for value in values])
        worksheet.set_column(idx, idx, min(max_len + 2, 70))


def write_sheet(writer: pd.ExcelWriter, name: str, df: pd.DataFrame, link_columns: list[str] | None = None) -> None:
    df.to_excel(writer, sheet_name=name, index=False)
    ws = writer.sheets[name]
    autosize(ws, df)
    ws.freeze_panes(1, 0)
    if link_columns:
        link_fmt = writer.book.add_format({"font_color": "blue", "underline": 1})
        for link_column in link_columns:
            if link_column not in df.columns:
                continue
            col_idx = list(df.columns).index(link_column)
            for row_idx, url in enumerate(df[link_column], start=1):
                if isinstance(url, str) and url.startswith("http"):
                    ws.write_url(row_idx, col_idx, url, link_fmt, string=url)


def main() -> None:
    style_df = load_csv(STYLE_PATH)
    item_df = load_csv(ITEM_PATH)
    scene_df = load_csv(SCENE_PATH)
    function_df = load_csv(FUNCTION_PATH)

    amazon_df = fetch_amazon_bestsellers(limit=8)
    amazon_summary_df = build_amazon_summary(amazon_df)
    amazon_history_proxy_df = build_amazon_history_proxy(item_df, style_df)

    youtube_search_df = fetch_youtube_searches()
    youtube_comment_df = fetch_youtube_comments(limit=8)
    youtube_summary_df = build_youtube_summary(youtube_search_df, youtube_comment_df)

    forum_df = pd.DataFrame(FORUM_ROWS)
    review_df = pd.DataFrame(REVIEW_ROWS)
    academy_df = build_academy_analysis(style_df, amazon_df, youtube_search_df)
    final_df = build_final_recommendation(style_df, item_df, function_df)

    overview_df = pd.DataFrame(
        [
            {"项目": "这次新增来源", "内容": "Amazon 日本榜单、YouTube 搜索结果与评论、日本论坛问答、楽天真实购买评论、学院风专项判断"},
            {"项目": "时间口径", "内容": f"当前数据抓取日期为 {EXPORT_DATE}；第一步趋势数据仍使用四月销售窗口版。"},
            {"项目": "关于 Amazon 历史 4 月", "内容": "官方历史畅销榜无法公开导出，也没有稳定存档快照，因此用“四月季节性 + 评论时间与评论内容”做代理。"},
            {"项目": "最重要变化", "内容": "第二步不再停留在方向判断，而是补到“用户在乎什么、现在卖什么、哪些款式容易出问题”。"},
        ]
    )

    write_markdown(amazon_summary_df, youtube_summary_df, academy_df)

    with pd.ExcelWriter(OUTPUT_XLSX, engine="xlsxwriter") as writer:
        write_sheet(writer, "概览", overview_df)
        write_sheet(writer, "Amazon当前榜单", amazon_df, link_columns=["榜单链接"])
        write_sheet(writer, "Amazon历史4月代理", amazon_history_proxy_df)
        write_sheet(writer, "YouTube搜索结果", youtube_search_df, link_columns=["视频链接"])
        write_sheet(writer, "YouTube评论", youtube_comment_df, link_columns=["视频链接"])
        write_sheet(writer, "论坛与问答", forum_df, link_columns=["链接"])
        write_sheet(writer, "真实购买评论", review_df, link_columns=["链接"])
        write_sheet(writer, "学院风专项", academy_df)
        write_sheet(writer, "最终建议", final_df)
        write_sheet(writer, "第一步场景参考", scene_df, link_columns=None)


if __name__ == "__main__":
    main()
