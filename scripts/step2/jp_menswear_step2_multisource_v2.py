from __future__ import annotations

import json
import re
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from deep_translator import GoogleTranslator
from youtube_comment_downloader import SORT_BY_POPULAR, YoutubeCommentDownloader


EXPORT_DATE = "2026-03-19"

APRIL_DIR = Path("april_sell_output")
OUT_DIR = Path("step2_plus_output")
OUT_DIR.mkdir(exist_ok=True)

STYLE_PATH = APRIL_DIR / "四月风格机会表.csv"
ITEM_PATH = APRIL_DIR / "四月商品机会表.csv"
FUNCTION_PATH = APRIL_DIR / "四月功能机会表.csv"
SCENE_PATH = APRIL_DIR / "四月场景机会表.csv"

OUTPUT_XLSX = OUT_DIR / f"日本男装第二步_多来源验证_v3_{EXPORT_DATE}.xlsx"
OUTPUT_MD = OUT_DIR / "第二步_多来源商业分析_v3.md"

HEADERS = {"User-Agent": "Mozilla/5.0"}
TRANSLATOR = GoogleTranslator(source="auto", target="zh-CN")
TRANSLATION_CACHE: dict[str, str] = {}


AMAZON_CATEGORIES = [
    {"label": "亚马逊日本_男士衬衫", "node": "24548215051", "group": "衬衫"},
    {"label": "亚马逊日本_男士开衫", "node": "2131438051", "group": "开衫"},
    {"label": "亚马逊日本_男士西装外套", "node": "5825572051", "group": "西装外套"},
    {"label": "亚马逊日本_男士牛仔裤", "node": "2133077051", "group": "牛仔裤"},
    {"label": "亚马逊日本_男士长裤", "node": "2131443051", "group": "长裤"},
]


YOUTUBE_MAIN_QUERIES = [
    {"query": "メンズ 春服 2026", "focus": "春季主流"},
    {"query": "メンズ 春服 オフィスカジュアル", "focus": "办公室休闲"},
    {"query": "UNIQLO 春服 メンズ 2026", "focus": "优衣库男装"},
    {"query": "メンズ 春 デニム コーデ", "focus": "牛仔穿搭"},
    {"query": "メンズ 春服 アメカジ", "focus": "美式休闲"},
    {"query": "メンズ アイビー プレッピー", "focus": "学院风"},
    {"query": "メンズ 春服 30代40代", "focus": "30-40岁"},
    {"query": "メンズ 春服 20代", "focus": "20代"},
]

YOUTUBE_AUX_QUERIES = [
    {"query": "女性目線 メンズ 春服 2026", "focus": "女性视角辅助"},
]

FEMALE_YOUTUBE_PATTERNS = [
    "女性目線",
    "女子目線",
    "女子ウケ",
    "彼氏",
    "モテ",
    "女性が着てたら",
]


FORUM_QUERIES = [
    ("办公室休闲", "site:detail.chiebukuro.yahoo.co.jp メンズ オフィスカジュアル"),
    ("办公室休闲", "site:detail.chiebukuro.yahoo.co.jp メンズ ビジカジ シャツ"),
    ("办公室休闲", "site:detail.chiebukuro.yahoo.co.jp メンズ ビジカジ ジャケット"),
    ("阔腿裤", "site:detail.chiebukuro.yahoo.co.jp メンズ ワイドパンツ 低身長"),
    ("阔腿裤", "site:detail.chiebukuro.yahoo.co.jp メンズ ワイドパンツ ダサい"),
    ("牛仔", "site:detail.chiebukuro.yahoo.co.jp メンズ デニムジャケット 春 コーデ"),
    ("海军蓝西装", "site:detail.chiebukuro.yahoo.co.jp メンズ 紺ブレ オフィスカジュアル"),
    ("套装", "site:detail.chiebukuro.yahoo.co.jp メンズ セットアップ オフィスカジュアル"),
    ("学院风", "site:detail.chiebukuro.yahoo.co.jp アイビー プレッピー メンズ ブランド"),
    ("学院风", "site:detail.chiebukuro.yahoo.co.jp スクールカーディガン メンズ"),
    ("学生预算", "site:detail.chiebukuro.yahoo.co.jp メンズ 学生 5000円 コーデ"),
    ("衬衫", "site:detail.chiebukuro.yahoo.co.jp メンズ 白シャツ 透ける"),
]


REVIEW_QUERIES = [
    ("免烫衬衫", "site:review.rakuten.co.jp メンズ ノーアイロン シャツ レビュー"),
    ("纽扣领衬衫", "site:review.rakuten.co.jp メンズ ボタンダウン シャツ レビュー"),
    ("海军蓝西装", "site:review.rakuten.co.jp 紺ブレザー メンズ レビュー"),
    ("直筒牛仔", "site:review.rakuten.co.jp メンズ レギュラーストレート デニム レビュー"),
    ("阔腿裤", "site:review.rakuten.co.jp メンズ ワイドパンツ レビュー"),
    ("开衫", "site:review.rakuten.co.jp メンズ カーディガン レビュー"),
    ("学院风开衫", "site:review.rakuten.co.jp スクールカーディガン メンズ レビュー"),
    ("卡其裤", "site:review.rakuten.co.jp メンズ チノパン レビュー"),
    ("弹力长裤", "site:review.rakuten.co.jp メンズ ストレッチ パンツ レビュー"),
    ("轻外套", "site:review.rakuten.co.jp メンズ ジャケット 春 レビュー"),
    ("套装", "site:review.rakuten.co.jp メンズ セットアップ レビュー"),
    ("防泼水", "site:review.rakuten.co.jp メンズ 撥水 ジャケット レビュー"),
]


MANUAL_FORUM_ROWS = [
    {
        "来源平台": "Yahoo!知恵袋",
        "目标适配性": "男装主来源",
        "主题": "办公室休闲",
        "帖子标题原文": "入社式に着ていく服装でオフィスカジュアルと指定されていて既にスーツを購入してしまっているのですがスーツを着ていくのは良くないですか？",
        "帖子摘要原文": "入社式は企業の経営陣、上位管理職が揃う重要な式典です。そこに参加する服装について企業側からわざわざ指定されてるならスーツ良くないですね。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q10259369790",
        "查询词": "manual",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "目标适配性": "男装主来源",
        "主题": "办公室休闲",
        "帖子标题原文": "入社式に着ていく服装でオフィスカジュアルと指定されていて既にスーツを購入してしまっているのですがスーツを着ていくのは良くないですか？",
        "帖子摘要原文": "オフィスカジュアルとはジャケットにチノパンなどスーツではないけど仕事で着ててもチャラく見えない服装です。とりあえず紺やグレーのジャケットかっとけば下はポロシャツやチノパンなどカジュアルでいい。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q10259369790",
        "查询词": "manual",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "目标适配性": "男装主来源",
        "主题": "办公室休闲",
        "帖子标题原文": "男性ですが、オフィスカジュアルとは具体的にどんな服装でしょうか？",
        "帖子摘要原文": "もしそうなら、オフィスカジュアルの要件はオフィス毎に異なります。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q10318500465",
        "查询词": "manual",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "目标适配性": "男装主来源",
        "主题": "办公室休闲",
        "帖子标题原文": "salomonがオフィスカジュアル・ビジネスカジュアルに適しているかについて",
        "帖子摘要原文": "ビジネスカジュアルは革靴。トレランシューズであるXT-6は完全にミスマッチかと。オフィスカジュアルとして履くならテープまでオールブラックのモデルを選ぶべきでしょう。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q12295877782",
        "查询词": "manual",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "目标适配性": "男装主来源",
        "主题": "衬衫 / 白T",
        "帖子标题原文": "メンズです。白Tシャツを着るとタンクトップが透ける現象がよく起きるのですが、皆さんはどうやって下着が見えないように白色の服を着こなしているんですか？",
        "帖子摘要原文": "タンクトップの色が問題なんですよ。見えにくい肌に近い色がいいんですよ。ベージュやオレンジにはなりますがメンズではあまりないのでレディース用を着ていますよ。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q13304268113",
        "查询词": "manual",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "目标适配性": "男装主来源",
        "主题": "衬衫 / 白T",
        "帖子标题原文": "メンズです。白Tシャツを着るとタンクトップが透ける現象がよく起きるのですが、皆さんはどうやって下着が見えないように白色の服を着こなしているんですか？",
        "帖子摘要原文": "自分は生地の問題かなーと思ってます。色が付いた服でも薄い生地だと下着が透けてしまいますし。なので厚手のTシャツ好きです。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q13304268113",
        "查询词": "manual",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "目标适配性": "男装主来源",
        "主题": "衬衫 / 白T",
        "帖子标题原文": "メンズです。白Tシャツを着るとタンクトップが透ける現象がよく起きるのですが、皆さんはどうやって下着が見えないように白色の服を着こなしているんですか？",
        "帖子摘要原文": "Tシャツの生地とタンクトップの色ですね。タンクトップは肌に近い色が透けにくいですよ。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q13304268113",
        "查询词": "manual",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "目标适配性": "男装主来源",
        "主题": "学院风",
        "帖子标题原文": "アイビー、トラッドファッションのブランドでかつて日本でも人気が高かったブルックスブラザーズとJプレスがありましたけどアメリカ本土では今でも高級紳士服ショップなんですか？",
        "帖子摘要原文": "当時の愛読書がメンズクラブでアイビートラッドはVANが倒産していたし、本格的なアメトラはブルックスだけしかないと憧れたブランドでしたね。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q10310329690",
        "查询词": "manual",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "目标适配性": "男装主来源",
        "主题": "学院风",
        "帖子标题原文": "LLBeanはアイビーやプレッピーに入りますか？",
        "帖子摘要原文": "LLBeanはアイビーやプレッピーに欠かせないブランドです。ただPatagoniaなどと合わせるともっとアメカジ寄りになります。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q11268982261",
        "查询词": "manual",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "目标适配性": "男装主来源",
        "主题": "阔腿裤",
        "帖子标题原文": "低身長だとワイドパンツは難しいですか？",
        "帖子摘要原文": "低身長だとワイドパンツは顔や上半身の比率が目立ちやすく、丈が長すぎるとバランスが崩れやすいです。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q12233510552",
        "查询词": "manual",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "目标适配性": "男装主来源",
        "主题": "学生预算",
        "帖子标题原文": "学生が5000円以内で買えるメンズブランドはありますか？",
        "帖子摘要原文": "学生なら5000円以内で、ZOZOTOWNでも買えて、コーデしやすいブランドを優先します。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q14286522609",
        "查询词": "manual",
    },
]


MANUAL_FORUM_ROWS += [
    {
        "来源平台": "Yahoo!知恵袋",
        "目标适配性": "男装主来源",
        "主题": "办公室休闲",
        "帖子标题原文": "50代男性です。会社ではオフィスカジュアルが推奨されていますが、ファッションに疎いので困っています。",
        "帖子摘要原文": "会社ではオフィスカジュアルが推奨されているが、何を買えば良いのか分からず、コスパの高いものも知りたいという相談。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q12307599952",
        "查询词": "web_search",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "目标适配性": "男装主来源",
        "主题": "办公室休闲",
        "帖子标题原文": "内定式があるのですが、スーツ禁止のオフィスカジュアルと言われました。",
        "帖子摘要原文": "オフィスカジュアル指定のところにスーツで行くと、逆に浮くし暑さ配慮の面でも不自然という声がある。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q14304769184",
        "查询词": "web_search",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "目标适配性": "男装主来源",
        "主题": "办公室休闲",
        "帖子标题原文": "就活中の大学生(男)です。『オフィスカジュアル』に適した服を持っていないです。",
        "帖子摘要原文": "男性就活生が、手持ちにオフィスカジュアルがなく、何を揃えるべきか迷っている。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q11319796976",
        "查询词": "web_search",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "目标适配性": "男装主来源",
        "主题": "办公室休闲",
        "帖子标题原文": "『オフィスカジュアル』の場合、スーツを着用しても大丈夫でしょうか？",
        "帖子摘要原文": "オフィスカジュアル指定の場にスーツで行くと、場に対して堅すぎるのではという不安が繰り返し出ている。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q13300013608",
        "查询词": "web_search",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "目标适配性": "男装主来源",
        "主题": "办公室休闲",
        "帖子标题原文": "現在就活中なのですが、オフィスカジュアルはどういった服装でしょうか？",
        "帖子摘要原文": "男性ならノーネクタイのボタンダウンシャツ、センタープレスパンツ、革ベルト、革靴が基準という具体的な回答がついている。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q13305542978",
        "查询词": "web_search",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "目标适配性": "男装主来源",
        "主题": "办公室休闲",
        "帖子标题原文": "男性のオフィスカジュアルの服はどこで買うのですか？",
        "帖子摘要原文": "ユニクロでもチノパンやスラックス、ボタンダウンシャツで組めるが、どこまでがオフィスカジュアルか見極めが必要という意見。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q13265325543",
        "查询词": "web_search",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "目标适配性": "男装主来源",
        "主题": "阔腿裤",
        "帖子标题原文": "低身長が着てたらとても違和感のあるコーディネートって何ですか？",
        "帖子摘要原文": "低身長だとワイドパンツのように下半身がダボダボする服は難しいのか、という不安が率直に出ている。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q13265450008",
        "查询词": "web_search",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "目标适配性": "男装主来源",
        "主题": "学院风",
        "帖子标题原文": "プレッピー、アイビー系の洋服のブランドについて。",
        "帖子摘要原文": "プレッピーだと思える着こなしとして、ステンカラーコート、コットンニット、白パン、ローファー、クラシックサングラスの組み合わせが語られている。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q11121108726",
        "查询词": "web_search",
    },
    {
        "来源平台": "Yahoo!知恵袋",
        "目标适配性": "男装主来源",
        "主题": "办公室休闲",
        "帖子标题原文": "入社式に着ていく服装でオフィスカジュアル指定なのですが、何を基準にすればいいですか？",
        "帖子摘要原文": "回答では『ジャケパンで検索してみてください』『明日はノータイでよろしい』とあり、入門者がまず理解すべき型としてジャケパンが挙がっている。",
        "链接": "https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q10259369790",
        "查询词": "web_search",
    },
]


MANUAL_REVIEW_ROWS = [
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "免烫衬衫",
        "评论页标题原文": "ワイシャツ 半袖 メンズ ノーアイロン ストレッチ ニット 半袖ワイシャツ",
        "评论摘要原文": "ノーアイロンで速乾性がいちばん大事。生地はメッシュのポロシャツって感じですが周りもクールビズなので違和感なくサラッと着れるみたいです。",
        "链接": "https://review.rakuten.co.jp/review/review/item/1/253792_10000617/1.1/",
        "查询词": "manual",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "免烫衬衫",
        "评论页标题原文": "ノーアイロン ニットシャツ 半袖 ワイシャツ メンズ",
        "评论摘要原文": "シャツなのに伸びて便利、スリムぎみなのでワンサイズ大き目を選ぶのもオススメです。シャツの生地も柔らかく伸びて非常に着やすい。",
        "链接": "https://review.rakuten.co.jp/item/1/253792_10008226/1.1/",
        "查询词": "manual",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "海军蓝西装",
        "评论页标题原文": "ジャケット メンズ 紺ブレ 紺 ブレザー チャコールグレー",
        "评论摘要原文": "グレー系のブレザー（銀釦）を探していました。初めて袖を通しましたが、ええ感じです。この価格なら十分に満足です。171cm68キロで標準、Mサイズでジャストでした。",
        "链接": "https://review.rakuten.co.jp/item/1/277415_10016219/1.1/",
        "查询词": "manual",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "海军蓝西装",
        "评论页标题原文": "紺ブレザー メンズ メタルボタン 金・銀釦 シングル 2つボタン",
        "评论摘要原文": "通勤カバンをバックパックに変えたため、ジャケットの肩や腰辺りの生地が傷むことを想定し、この価格帯で紺ブレを探していました。",
        "链接": "https://review.rakuten.co.jp/item/1/310736_10000215/1.1/",
        "查询词": "manual",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "海军蓝西装",
        "评论页标题原文": "紺ブレザー メンズ | ウール100% 紺ブレ ジャケット",
        "评论摘要原文": "レビューにもあったように、すごく生地もしっかりしていて、1万円以下の商品ではないです。値段のわりにクオリティーは高かったです。",
        "链接": "https://review.rakuten.co.jp/item/1/223890_10005789/1.1/",
        "查询词": "manual",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "直筒牛仔",
        "评论页标题原文": "エドウィン 503 レギュラーストレートパンツ メンズ",
        "评论摘要原文": "サイズ合わせてが出来て便利。いつものサイズと同じで問題なし。",
        "链接": "https://review.rakuten.co.jp/item/1/227223_10009113/1.1/",
        "查询词": "manual",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "直筒牛仔",
        "评论页标题原文": "EDWIN 505ZXX レギュラーストレート デニム ジーンズ",
        "评论摘要原文": "大き目サイズでしたが、ベルトをすると問題ありません。実際の商品も、新品の状態から風合いがソフトで履きやすいです。これぐらいがお手頃価格だと思います。",
        "链接": "https://review.rakuten.co.jp/item/1/222661_10004538/1.1/",
        "查询词": "manual",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "直筒牛仔",
        "评论页标题原文": "Levi's 505 レギュラーストレート ジーンズ",
        "评论摘要原文": "ワンサイズ上のサイズなら緩く穿けるので、503の代用品を探していてダボっとジーンズを穿きたい人にはお勧めです。黒がなかなか売っていないので探していたところ、こちらにたどり着きました。",
        "链接": "https://review.rakuten.co.jp/review/review/review/item/1/222661_10006717/1.1/",
        "查询词": "manual",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "直筒牛仔",
        "评论页标题原文": "Levi's 505 レギュラーストレート ジーンズ メンズ",
        "评论摘要原文": "ジーンズをジャストサイズで履きたいけど窮屈なのは困ると痛し痒しな状態の時に本製品を発見しました。見た目は普通にリーバイスで良いお買い物になりました。",
        "链接": "https://review.rakuten.co.jp/review/item/1/339642_10001391/1.1/",
        "查询词": "manual",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "学院风开衫",
        "评论页标题原文": "EASTBOY 綿カーディガン ハイゲージ12G スクールカーディガン",
        "评论摘要原文": "綿100%で毛玉ができにくく、ハイゲージなのでジャケットのインにすっきり着られて気に入っています。",
        "链接": "https://review.rakuten.co.jp/item/1/225462_10005692/1.1/",
        "查询词": "manual",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "开衫",
        "评论页标题原文": "カーディガン POLO BCS 大きいサイズ メンズ 薄手",
        "评论摘要原文": "とても、肌触りの良い商品でした。",
        "链接": "https://review.rakuten.co.jp/item/1/270856_10017565/1.0/",
        "查询词": "manual",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "宽裤 / 工作裤",
        "评论页标题原文": "ウォバッシュ ストライプ ワークパンツ メンズ",
        "评论摘要原文": "記載の寸法と実物では寸法に結構な差がある場合もあるようですので細かな寸法を気にされる方は事前に確認することをお勧めします。",
        "链接": "https://review.rakuten.co.jp/review/review/review/item/1/428943_10000016/1.1/",
        "查询词": "manual",
    },
]


MANUAL_REVIEW_ROWS += [
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "免烫衬衫",
        "评论页标题原文": "別格ノーアイロン シャツ メンズ Yシャツ クールビズ 形態安定",
        "评论摘要原文": "ノーアイロンシャツは手入れが楽で助かる。生地には伸びがあり、動きやすそうだという評価。",
        "链接": "https://review.rakuten.co.jp/review/item/1/251912_10007766/1.1/",
        "查询词": "web_search",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "免烫衬衫",
        "评论页标题原文": "別格ノーアイロン シャツ メンズ Yシャツ クールビズ 形態安定",
        "评论摘要原文": "サイズは全体的に小さめで、LLでも余裕が少ない。襟が高くて少し苦しいという不満もある。",
        "链接": "https://review.rakuten.co.jp/review/item/1/251912_10007766/1.1/",
        "查询词": "web_search",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "免烫衬衫",
        "评论页标题原文": "別格ノーアイロン シャツ メンズ Yシャツ クールビズ 形態安定",
        "评论摘要原文": "長袖で気に入っていたので半袖も購入。少しボタンホールはきついが許容範囲という声。",
        "链接": "https://review.rakuten.co.jp/review/item/1/251912_10007766/1.1/",
        "查询词": "web_search",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "免烫衬衫",
        "评论页标题原文": "別格ノーアイロン シャツ メンズ Yシャツ クールビズ 形態安定",
        "评论摘要原文": "毎日の仕事用として、生地の厚さが程よく、梱包も丁寧で、また買いたいという評価。",
        "链接": "https://review.rakuten.co.jp/review/item/1/251912_10007766/1.1/",
        "查询词": "web_search",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "免烫衬衫",
        "评论页标题原文": "別格ノーアイロン シャツ メンズ Yシャツ クールビズ 形態安定",
        "评论摘要原文": "洗濯しても形状が崩れず満足という意見があり、家庭洗濯耐性が評価されている。",
        "链接": "https://review.rakuten.co.jp/review/item/1/251912_10007766/1.1/",
        "查询词": "web_search",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "免烫衬衫",
        "评论页标题原文": "別格ノーアイロン シャツ メンズ Yシャツ クールビズ 形態安定",
        "评论摘要原文": "生地はやや薄めで、ストレッチ性は他社と比べてあまり感じない、洗濯後に若干シワがつくという声もある。",
        "链接": "https://review.rakuten.co.jp/review/item/1/251912_10007766/1.1/",
        "查询词": "web_search",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "直筒牛仔",
        "评论页标题原文": "Lee AMERICAN RIDERS 101 ストレート カウボーイパンツ メンズ デニム",
        "评论摘要原文": "裾直しをしなくてよかったほど丈が長すぎず、ネット購入でも良い買い物だったという評価。",
        "链接": "https://review.rakuten.co.jp/item/1/339642_10001283/1.1/",
        "查询词": "web_search",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "西裤",
        "评论页标题原文": "スラックス メンズ パンツ 紳士 ウォッシャブル 春夏用 パンツ",
        "评论摘要原文": "COOLBIZが始まってズボンだけ大量に必要になったが、この商品は品質が良く、サラサラで汗も乾きやすいという高評価。",
        "链接": "https://review.rakuten.co.jp/review/review/review/item/1/260455_10001231/1.1/",
        "查询词": "web_search",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "西裤",
        "评论页标题原文": "スラックス メンズ パンツ 紳士 ウォッシャブル 春夏用 パンツ",
        "评论摘要原文": "この値段でこの品質なら十分。バリエーションが増えたらリピートしたいという声がある。",
        "链接": "https://review.rakuten.co.jp/review/review/review/item/1/260455_10001231/1.1/",
        "查询词": "web_search",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "西裤",
        "评论页标题原文": "スラックス メンズ パンツ 紳士 ウォッシャブル 春夏用 パンツ",
        "评论摘要原文": "春夏用で家で洗える点を重視して購入し、履きやすいのでリピートしたいという家族目線の評価。",
        "链接": "https://review.rakuten.co.jp/review/review/review/item/1/260455_10001231/1.1/",
        "查询词": "web_search",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "西裤",
        "评论页标题原文": "スラックス クールビズ 夏用 清涼素材 サラサラ ゴルフパンツ メンズ",
        "评论摘要原文": "酷暑が続きそうなので購入。先に買った春夏用よりも柔らかく、通気性が良く満足という声。",
        "链接": "https://review.rakuten.co.jp/review/item/1/277415_10006322/1.1/",
        "查询词": "web_search",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "西裤",
        "评论页标题原文": "スラックス クールビズ 夏用 清涼素材 サラサラ ゴルフパンツ メンズ",
        "评论摘要原文": "想像以上にダボダボで一昔前のシルエットに見えるという否定的レビューがあり、太さのコントロールが重要だと分かる。",
        "链接": "https://review.rakuten.co.jp/review/item/1/277415_10006322/1.1/",
        "查询词": "web_search",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "西裤",
        "评论页标题原文": "スラックス クールビズ 夏用 清涼素材 サラサラ ゴルフパンツ メンズ",
        "评论摘要原文": "濃紺85センチを購入し、薄手の生地でこれからの時期にぴったり、またリピートしたいという声。",
        "链接": "https://review.rakuten.co.jp/review/item/1/277415_10006322/1.1/",
        "查询词": "web_search",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "春季夹克",
        "评论页标题原文": "サマージャケット メンズ テーラードジャケット 春服 春物 夏服",
        "评论摘要原文": "サイズがぴったりで、デザインも良く、通気性も良いので気に入ったという男性レビュー。",
        "链接": "https://review.rakuten.co.jp/item/1/420334_10000134/1.1/",
        "查询词": "web_search",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "春季夹克",
        "评论页标题原文": "サマージャケット メンズ テーラードジャケット 春服 春物 夏服",
        "评论摘要原文": "オリジワはあったが、スチームで簡単に取れたという声があり、皱褶而非版型是主要小问题。",
        "链接": "https://review.rakuten.co.jp/item/1/420334_10000134/1.1/",
        "查询词": "web_search",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "春季夹克",
        "评论页标题原文": "ジャケット メンズ 春夏 サマージャケット 吸水速乾素材 日本製生地",
        "评论摘要原文": "グリーン系は初めてだが風合いが良く、軽いのにしっかりしていて使用シーンが多そうという評価。",
        "链接": "https://review.rakuten.co.jp/review/item/1/277415_10013518/1.1/",
        "查询词": "web_search",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "春季夹克",
        "评论页标题原文": "ジャケット メンズ 春夏 サマージャケット 吸水速乾素材 日本製生地",
        "评论摘要原文": "暑いところでも着られそうで良かったという声があり、春夏の通勤ジャケットに必要な軽さが評価されている。",
        "链接": "https://review.rakuten.co.jp/review/item/1/277415_10013518/1.1/",
        "查询词": "web_search",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "阔腿裤",
        "评论页标题原文": "ワイドスラックス メンズ ワイドパンツ メンズ 夏 スラックス",
        "评论摘要原文": "見た目の美しさ、レーヨンの質感、履き心地、手触り、シルエットまで高く評価されている。",
        "链接": "https://review.rakuten.co.jp/review/review/review/item/1/223605_10019785",
        "查询词": "web_search",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "阔腿裤",
        "评论页标题原文": "ワイドスラックス メンズ ワイドパンツ メンズ 夏 スラックス",
        "评论摘要原文": "ストレッチ感はないが、薄手でドレープ感があり、楽に履けておしゃれになれるという評価。",
        "链接": "https://review.rakuten.co.jp/review/review/review/item/1/223605_10019785",
        "查询词": "web_search",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "卡其裤",
        "评论页标题原文": "チノパン メンズ テーパードパンツ ストレッチ オフィス カジュアル",
        "评论摘要原文": "Lだと腰回りが余り、Mだと太ももから下の余裕が少ない。サイズを一つ変えるだけでヒップのシルエットが変わるという声。",
        "链接": "https://review.rakuten.co.jp/review/item/1/227930_10019184/1.1/",
        "查询词": "web_search",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "西裤",
        "评论页标题原文": "スラックス メンズ 紳士 ビジネス ノータック アジャスター付 ストレッチ",
        "评论摘要原文": "ウエスト調整と股下条件に合い、軽くて動きやすくて楽だという声がある。",
        "链接": "https://review.rakuten.co.jp/item/1/396942_10000035/1.1/",
        "查询词": "web_search",
    },
    {
        "来源平台": "楽天评论",
        "目标适配性": "男装主来源",
        "品类主题": "西裤",
        "评论页标题原文": "スラックス メンズ ビジネスパンツ ストレッチ パンツ メンズ",
        "评论摘要原文": "ワンサイズ上は大きすぎて交換になった。さらに股上が長く、腰履きのスリムパンツとしては見え方が悪いという厳しめの声。",
        "链接": "https://review.rakuten.co.jp/item/1/385618_10003984/1.1/",
        "查询词": "web_search",
    },
]


def load_csv(path: Path) -> pd.DataFrame:
    return pd.read_csv(path, encoding="utf-8-sig")


def pick_row(df: pd.DataFrame, keyword_zh: str) -> pd.Series:
    return df.loc[df["中文关键词"] == keyword_zh].iloc[0]


def translate_text(text: str) -> str:
    text = (text or "").strip()
    if not text:
        return ""
    if text in TRANSLATION_CACHE:
        return TRANSLATION_CACHE[text]
    try:
        translated = TRANSLATOR.translate(text)
    except Exception:
        translated = text
    TRANSLATION_CACHE[text] = translated
    return translated


def add_translation_column(df: pd.DataFrame, source_col: str, target_col: str) -> pd.DataFrame:
    if df.empty or source_col not in df.columns:
        df[target_col] = ""
        return df
    unique_texts = [str(text).strip() for text in df[source_col].fillna("").tolist() if str(text).strip()]
    unique_texts = list(dict.fromkeys(unique_texts))
    missing = [text for text in unique_texts if text not in TRANSLATION_CACHE]
    for i in range(0, len(missing), 20):
        chunk = missing[i : i + 20]
        try:
            translated = TRANSLATOR.translate_batch(chunk)
        except Exception:
            translated = [translate_text(text) for text in chunk]
        for src, zh in zip(chunk, translated):
            TRANSLATION_CACHE[src] = zh if zh else src
    df[target_col] = df[source_col].fillna("").map(lambda text: TRANSLATION_CACHE.get(str(text).strip(), ""))
    return df


def ddg_unwrap(url: str) -> str:
    if "duckduckgo.com/l/?" not in url:
        return url
    parsed = urlparse(url)
    uddg = parse_qs(parsed.query).get("uddg", [""])[0]
    return unquote(uddg) if uddg else url


def ddg_search(query: str, domain_keyword: str, limit: int = 4) -> list[dict[str, str]]:
    url = "https://html.duckduckgo.com/html/?q=" + quote(query)
    html = requests.get(url, headers=HEADERS, timeout=30).text
    soup = BeautifulSoup(html, "html.parser")
    rows: list[dict[str, str]] = []
    for result in soup.select(".result"):
        link_tag = result.select_one(".result__a")
        snippet_tag = result.select_one(".result__snippet")
        if not link_tag:
            continue
        real_link = ddg_unwrap(link_tag.get("href", ""))
        if domain_keyword not in real_link:
            continue
        title = link_tag.get_text(" ", strip=True)
        snippet = snippet_tag.get_text(" ", strip=True) if snippet_tag else ""
        rows.append({"标题原文": title, "摘要原文": snippet, "链接": real_link, "查询词": query})
        if len(rows) >= limit:
            break
    return rows


def clean_amazon_title(card: BeautifulSoup) -> str:
    selectors = [
        "div._cDEzb_p13n-sc-css-line-clamp-2_EWgCb",
        "div._cDEzb_p13n-sc-css-line-clamp-3_g3dy1",
    ]
    for selector in selectors:
        node = card.select_one(selector)
        if node:
            text = node.get_text(" ", strip=True)
            if text:
                return text
    img = card.select_one("img")
    return img.get("alt", "").strip() if img else ""


def fetch_amazon_bestsellers(limit: int = 15) -> pd.DataFrame:
    rows: list[dict[str, str | int]] = []
    for category in AMAZON_CATEGORIES:
        url = f"https://www.amazon.co.jp/gp/bestsellers/fashion/{category['node']}"
        html = ""
        cards: list[BeautifulSoup] = []
        for _ in range(3):
            html = requests.get(url, headers=HEADERS, timeout=30).text
            soup = BeautifulSoup(html, "html.parser")
            cards = soup.select("div.zg-grid-general-faceout")
            if cards and "opfcaptcha" not in html:
                break
        for rank, card in enumerate(cards[:limit], start=1):
            title = clean_amazon_title(card)
            rating_tag = card.select_one(".a-icon-alt")
            review_link = card.select_one('a[href*="/product-reviews/"]')
            price_tag = card.select_one("span._cDEzb_p13n-sc-price_3mJ9Z") or card.select_one(".p13n-sc-price")
            root = card.select_one("div.p13n-sc-uncoverable-faceout")
            review_count = ""
            if review_link:
                small = review_link.select_one("span.a-size-small")
                if small:
                    review_count = small.get_text(" ", strip=True)
            rows.append(
                {
                    "来源组": category["label"],
                    "目标适配性": "男装主来源",
                    "商品组": category["group"],
                    "当前排名": rank,
                    "商品标题原文": title,
                    "评分": rating_tag.get_text(" ", strip=True) if rating_tag else "",
                    "评论量": review_count,
                    "价格": price_tag.get_text(" ", strip=True) if price_tag else "",
                    "ASIN": root.get("id", "") if root else "",
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


def youtube_search(query: str, focus: str, limit: int = 10) -> list[dict[str, str]]:
    url = "https://www.youtube.com/results?search_query=" + quote(query)
    html = requests.get(url, headers=HEADERS, timeout=30).text
    matched = re.search(r"var ytInitialData = (\{.*?\});", html)
    if not matched:
        return []
    data = json.loads(matched.group(1))
    videos: list[dict[str, object]] = []
    find_video_renderers(data, videos)
    rows: list[dict[str, str]] = []
    for video in videos[:limit]:
        title = "".join(run.get("text", "") for run in video.get("title", {}).get("runs", []))
        channel = "".join(run.get("text", "") for run in video.get("ownerText", {}).get("runs", []))
        views = video.get("viewCountText", {}).get("simpleText", "")
        published = video.get("publishedTimeText", {}).get("simpleText", "")
        video_id = video.get("videoId", "")
        role = "男装主来源"
        blob = f"{title} {channel}"
        if any(token in blob for token in FEMALE_YOUTUBE_PATTERNS):
            role = "女装/女性视角辅助"
        rows.append(
            {
                "查询词": query,
                "关注方向": focus,
                "目标适配性": role,
                "视频标题原文": title,
                "频道": channel,
                "观看量": views,
                "发布时间": published,
                "视频ID": video_id,
                "视频链接": f"https://www.youtube.com/watch?v={video_id}",
            }
        )
    return rows


def fetch_youtube_searches() -> tuple[pd.DataFrame, pd.DataFrame]:
    main_rows: list[dict[str, str]] = []
    aux_rows: list[dict[str, str]] = []
    for item in YOUTUBE_MAIN_QUERIES:
        for row in youtube_search(item["query"], item["focus"], limit=12):
            if row["目标适配性"] == "男装主来源":
                main_rows.append(row)
            else:
                aux_rows.append(row)
    for item in YOUTUBE_AUX_QUERIES:
        for row in youtube_search(item["query"], item["focus"], limit=8):
            row["目标适配性"] = "女装/女性视角辅助"
            aux_rows.append(row)
    return pd.DataFrame(main_rows), pd.DataFrame(aux_rows)


def select_comment_videos(main_df: pd.DataFrame) -> list[dict[str, str]]:
    picks: list[dict[str, str]] = []
    rules = [
        ("办公室休闲", 3),
        ("春季主流", 3),
        ("优衣库男装", 2),
        ("牛仔穿搭", 2),
        ("美式休闲", 2),
        ("学院风", 1),
        ("30-40岁", 1),
        ("20代", 1),
    ]
    for focus, count in rules:
        sub = main_df.loc[main_df["关注方向"] == focus].head(count)
        picks.extend(sub.to_dict("records"))
    unique: dict[str, dict[str, str]] = {}
    for row in picks:
        if row["视频ID"] not in unique:
            unique[row["视频ID"]] = row
    return list(unique.values())


def fetch_youtube_comments(video_rows: list[dict[str, str]], limit_per_video: int = 8) -> pd.DataFrame:
    downloader = YoutubeCommentDownloader()
    rows: list[dict[str, str]] = []
    for video in video_rows:
        count = 0
        for comment in downloader.get_comments_from_url(video["视频链接"], sort_by=SORT_BY_POPULAR):
            if comment.get("reply"):
                continue
            text = str(comment.get("text", "")).replace("\n", " / ").strip()
            if not text or len(text) > 260:
                continue
            rows.append(
                {
                    "关注方向": video["关注方向"],
                    "目标适配性": video["目标适配性"],
                    "视频标题原文": video["视频标题原文"],
                    "频道": video["频道"],
                    "评论作者": comment.get("author", ""),
                    "点赞数": comment.get("votes", ""),
                    "评论原文": text,
                    "视频链接": video["视频链接"],
                }
            )
            count += 1
            if count >= limit_per_video:
                break
    return pd.DataFrame(rows)


def append_manual_rows(df: pd.DataFrame, manual_rows: list[dict[str, str]], subset: list[str]) -> pd.DataFrame:
    manual_df = pd.DataFrame(manual_rows)
    if df.empty:
        return manual_df.drop_duplicates(subset=subset, keep="first").reset_index(drop=True)
    return (
        pd.concat([df, manual_df], ignore_index=True)
        .drop_duplicates(subset=subset, keep="first")
        .reset_index(drop=True)
    )


def build_forum_df(limit_per_query: int = 2) -> pd.DataFrame:
    rows: list[dict[str, str]] = []
    seen: set[str] = set()
    for topic, query in FORUM_QUERIES:
        for result in ddg_search(query, "detail.chiebukuro.yahoo.co.jp", limit=limit_per_query):
            if result["链接"] in seen:
                continue
            seen.add(result["链接"])
            rows.append(
                {
                    "来源平台": "Yahoo!知恵袋",
                    "目标适配性": "男装主来源",
                    "主题": topic,
                    "帖子标题原文": result["标题原文"],
                    "帖子摘要原文": result["摘要原文"],
                    "链接": result["链接"],
                    "查询词": query,
                }
            )
    columns = ["来源平台", "目标适配性", "主题", "帖子标题原文", "帖子摘要原文", "链接", "查询词"]
    df = pd.DataFrame(rows, columns=columns)
    return append_manual_rows(df, MANUAL_FORUM_ROWS, subset=["链接", "帖子摘要原文"])


def build_review_df(limit_per_query: int = 2) -> pd.DataFrame:
    rows: list[dict[str, str]] = []
    seen: set[str] = set()
    for topic, query in REVIEW_QUERIES:
        for result in ddg_search(query, "review.rakuten.co.jp", limit=limit_per_query):
            if result["链接"] in seen:
                continue
            seen.add(result["链接"])
            rows.append(
                {
                    "来源平台": "楽天评论",
                    "目标适配性": "男装主来源",
                    "品类主题": topic,
                    "评论页标题原文": result["标题原文"],
                    "评论摘要原文": result["摘要原文"],
                    "链接": result["链接"],
                    "查询词": query,
                }
            )
    columns = ["来源平台", "目标适配性", "品类主题", "评论页标题原文", "评论摘要原文", "链接", "查询词"]
    df = pd.DataFrame(rows, columns=columns)
    return append_manual_rows(df, MANUAL_REVIEW_ROWS, subset=["链接", "评论摘要原文"])


def build_source_audit(
    amazon_df: pd.DataFrame,
    yt_main_df: pd.DataFrame,
    yt_aux_df: pd.DataFrame,
    forum_df: pd.DataFrame,
    review_df: pd.DataFrame,
) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "来源": "Amazon 日本",
                "当前样本量": len(amazon_df),
                "是否混入女装": "否",
                "处理结果": "全部来自男士分类榜单，可作为主来源。",
            },
            {
                "来源": "YouTube 主来源",
                "当前样本量": len(yt_main_df),
                "是否混入女装": "否",
                "处理结果": "只保留男装主来源视频进入主分析。",
            },
            {
                "来源": "YouTube 辅助来源",
                "当前样本量": len(yt_aux_df),
                "是否混入女装": "是",
                "处理结果": "单独放到辅助参考，不进入主结论。",
            },
            {
                "来源": "日本论坛问答",
                "当前样本量": len(forum_df),
                "是否混入女装": "基本否",
                "处理结果": "查询词限定为男装场景与男士问题页，保留为主来源。",
            },
            {
                "来源": "真实购买评论",
                "当前样本量": len(review_df),
                "是否混入女装": "基本否",
                "处理结果": "商品页限定男装品类；个别评论者可能是代买者，但不作为女装风格样本。",
            },
        ]
    )


def build_sample_assessment(
    amazon_df: pd.DataFrame,
    yt_main_df: pd.DataFrame,
    yt_comments_df: pd.DataFrame,
    forum_df: pd.DataFrame,
    review_df: pd.DataFrame,
) -> pd.DataFrame:
    rows = [
        ("Amazon 当前榜单商品", len(amazon_df), "60-100", "用于判断当前卖得最好的商品结构"),
        ("YouTube 男装视频", len(yt_main_df), "50-80", "用于判断内容热度和男装表达方式"),
        ("YouTube 男装评论", len(yt_comments_df), "80-120", "用于判断真实痛点和接受度"),
        ("日本论坛问答", len(forum_df), "20-30", "用于判断用户顾虑和决策障碍"),
        ("真实购买评论", len(review_df), "30-50", "用于判断卖点、差评点和尺寸/面料痛点"),
    ]
    data = []
    for name, actual, target, use_case in rows:
        low, high = [int(x) for x in target.split("-")]
        status = "达标" if actual >= low else "不足"
        gap = max(low - actual, 0)
        expand_ratio = 1.0 if actual >= low else round(low / max(actual, 1), 2)
        data.append(
            {
                "来源": name,
                "当前样本量": actual,
                "建议样本量": target,
                "最低建议量": low,
                "还需补充": gap,
                "最低建议扩充倍数": expand_ratio,
                "是否足够": status,
                "为什么需要这个量": use_case,
            }
        )
    return pd.DataFrame(data)


def build_first_step_reference(
    style_df: pd.DataFrame,
    item_df: pd.DataFrame,
    function_df: pd.DataFrame,
    scene_df: pd.DataFrame,
) -> pd.DataFrame:
    targets = [
        ("风格", "办公室休闲男装", style_df),
        ("风格", "美式休闲男装", style_df),
        ("风格", "学院派男装", style_df),
        ("商品", "男士衬衫", item_df),
        ("商品", "男士牛仔裤", item_df),
        ("商品", "男士西裤", item_df),
        ("商品", "男士套装", item_df),
        ("商品", "男士夹克", item_df),
        ("功能", "弹力男装", function_df),
        ("功能", "防泼水男装", function_df),
        ("场景", "办公室休闲男装", scene_df),
        ("场景", "通勤男装", scene_df),
    ]
    rows = []
    for cat, keyword, df in targets:
        row = pick_row(df, keyword)
        rows.append(
            {
                "来源层级": "第一步（低优先级参考）",
                "分类": cat,
                "中文关键词": keyword,
                "第一步判断": row["判断"],
                "同比变化（%）": row["同比变化（%）"],
                "4月季节指数": row["4月季节指数"],
                "说明": "这张表只作为趋势底图，不参与第二步评论和帖子统计。",
            }
        )
    return pd.DataFrame(rows)


def build_integrated_summary(
    style_df: pd.DataFrame,
    item_df: pd.DataFrame,
    function_df: pd.DataFrame,
    amazon_df: pd.DataFrame,
    yt_comments_df: pd.DataFrame,
    forum_df: pd.DataFrame,
    review_df: pd.DataFrame,
) -> pd.DataFrame:
    office = pick_row(style_df, "办公室休闲男装")
    amekaji = pick_row(style_df, "美式休闲男装")
    preppy = pick_row(style_df, "学院派男装")
    shirt = pick_row(item_df, "男士衬衫")
    denim = pick_row(item_df, "男士牛仔裤")
    slacks = pick_row(item_df, "男士西裤")
    setup = pick_row(item_df, "男士套装")
    blouson = pick_row(item_df, "男士夹克")
    stretch = pick_row(function_df, "弹力男装")

    def count_contains(df: pd.DataFrame, col: str, pattern: str) -> int:
        if df.empty:
            return 0
        return int(df[col].fillna("").str.contains(pattern, case=False, regex=True).sum())

    return pd.DataFrame(
        [
            {
                "方向": "办公室休闲",
                "第一步（低优先级）": f"{office['判断']}；同比 {office['同比变化（%）']}%",
                "第二步（高优先级）": f"YouTube 办公室休闲评论 {count_contains(yt_comments_df, '关注方向', '办公室休闲')} 条；论坛相关帖子 {count_contains(forum_df, '主题', '办公室休闲')} 条。",
                "最终动作": "维持主线，并重点做衬衫、西裤、轻西装外套。",
            },
            {
                "方向": "美式休闲",
                "第一步（低优先级）": f"{amekaji['判断']}；同比 {amekaji['同比变化（%）']}%",
                "第二步（高优先级）": f"YouTube 美式休闲/牛仔相关评论 {count_contains(yt_comments_df, '关注方向', '美式休闲|牛仔穿搭')} 条；Amazon 牛仔榜单 {len(amazon_df.loc[amazon_df['商品组']=='牛仔裤'])} 条。",
                "最终动作": "做清爽美式，不做厚重复古，重点放牛仔和短外套。",
            },
            {
                "方向": "衬衫",
                "第一步（低优先级）": f"{shirt['判断']}；四月季节指数 {shirt['4月季节指数']}",
                "第二步（高优先级）": f"Amazon 衬衫榜单 {len(amazon_df.loc[amazon_df['商品组']=='衬衫'])} 条；真实购买评论 {count_contains(review_df, '品类主题', '衬衫')} 条。",
                "最终动作": "主推免烫、快干、常规领、蓝白条和白衬衫。",
            },
            {
                "方向": "牛仔裤",
                "第一步（低优先级）": f"{denim['判断']}；四月季节指数 {denim['4月季节指数']}",
                "第二步（高优先级）": f"Amazon 牛仔榜单 {len(amazon_df.loc[amazon_df['商品组']=='牛仔裤'])} 条；评论强调直筒、舒适、长度。",
                "最终动作": "同时保留正统直筒和轻宽松，不做单一极宽版型。",
            },
            {
                "方向": "西裤",
                "第一步（低优先级）": f"{slacks['判断']}；同比 {slacks['同比变化（%）']}%",
                "第二步（高优先级）": f"论坛和评论都在强调办公室休闲、通勤、弹力、好走动。",
                "最终动作": "做双褶、轻弹、可通勤、可成套。",
            },
            {
                "方向": "轻套装",
                "第一步（低优先级）": f"{setup['判断']}；四月季节指数 {setup['4月季节指数']}",
                "第二步（高优先级）": f"论坛里面试和上班场景反复提到感动ジャケット/パンツ一类轻正式组合。",
                "最终动作": "做轻正式和轻机能两条线，不做纯运动套装。",
            },
            {
                "方向": "短夹克",
                "第一步（低优先级）": f"{blouson['判断']}；四月季节指数 {blouson['4月季节指数']}",
                "第二步（高优先级）": f"YouTube 春季主流与美式内容都反复出现短外套、短夹克、牛仔夹克。",
                "最终动作": "做短、利落、易搭配的夹克，不做厚重长款。",
            },
            {
                "方向": "学院风",
                "第一步（低优先级）": f"{preppy['判断']}；同比 {preppy['同比变化（%）']}%",
                "第二步（高优先级）": f"Amazon 当前榜单有 school cardigan 与 school blazer；YouTube 学院风内容热度低于主流。",
                "最终动作": "做轻学院胶囊，不做品牌主线。",
            },
            {
                "方向": "功能卖点",
                "第一步（低优先级）": f"{stretch['判断']}；四月季节指数 {stretch['4月季节指数']}",
                "第二步（高优先级）": "Amazon 衬衫与长裤榜单、楽天评论共同支持免烫、快干、弹力、易护理。",
                "最终动作": "卖点优先写免烫、易护理、弹力、防泼水。",
            },
        ]
    )


def build_academy_analysis(
    style_df: pd.DataFrame,
    amazon_df: pd.DataFrame,
    yt_main_df: pd.DataFrame,
    forum_df: pd.DataFrame,
    review_df: pd.DataFrame,
) -> pd.DataFrame:
    preppy = pick_row(style_df, "学院派男装")
    cardigan_df = amazon_df.loc[amazon_df["商品组"] == "开衫"]
    blazer_df = amazon_df.loc[amazon_df["商品组"] == "西装外套"]
    school_cardigan = int(cardigan_df["商品标题原文"].fillna("").str.contains("スクール|school", case=False, regex=True).sum())
    school_blazer = int(blazer_df["商品标题原文"].fillna("").str.contains("スクール|school", case=False, regex=True).sum())
    preppy_videos = int(yt_main_df["关注方向"].fillna("").str.contains("学院风").sum())
    preppy_forum = int(forum_df["主题"].fillna("").str.contains("学院风").sum())
    review_series = review_df["品类主题"] if "品类主题" in review_df.columns else pd.Series(dtype=str)
    preppy_review = int(review_series.fillna("").astype(str).str.contains("学院风|开衫").sum())
    return pd.DataFrame(
        [
            {"指标": "第一步搜索体量", "结果": preppy["近52周均值"], "说明": "学院派同比涨，但总体搜索量仍小。"},
            {"指标": "第一步同比", "结果": preppy["同比变化（%）"], "说明": "说明它是上升中的小众方向，而不是大盘方向。"},
            {"指标": "Amazon 开衫榜单", "结果": f"{school_cardigan} 个 school cardigan 相关商品", "说明": "学院风真实购买主要落在基础 V 领和 school cardigan。"},
            {"指标": "Amazon 西装外套榜单", "结果": f"{school_blazer} 个 school blazer 相关商品", "说明": "学院风的外套入口是紺ブレ和 school blazer。"},
            {"指标": "YouTube 男装视频", "结果": f"{preppy_videos} 条学院风男装视频进入主来源", "说明": "有内容，但量和热度明显低于办公室休闲、优衣库春装、牛仔。"},
            {"指标": "论坛问答", "结果": f"{preppy_forum} 条学院风相关帖子", "说明": "更多在问品牌归属、风格理解和学生预算。"},
            {"指标": "真实购买评论", "结果": f"{preppy_review} 条开衫/学院风相关评论", "说明": "购买需求真实存在，但集中在基础单品，不在整套造型。"},
            {"指标": "最终结论", "结果": "可做轻学院胶囊，不建议做主线", "说明": "建议用紺ブレ、V 领开衫、牛津衬衫、直筒牛仔 / 卡其裤去表达。"},
        ]
    )


def build_source_summary(amazon_df: pd.DataFrame, yt_main_df: pd.DataFrame, forum_df: pd.DataFrame, review_df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "来源": "Amazon 日本",
                "核心发现": "衬衫榜单被免烫、快干、商务衬衫占据；西装外套榜单前列是紺ブレ；牛仔同时存在直筒与宽松。",
                "商业含义": "四月卖点要偏功能和通勤，海军蓝西装外套与牛仔是高确定性单品。",
            },
            {
                "来源": "YouTube 男装主来源",
                "核心发现": f"主来源视频 {len(yt_main_df)} 条，评论区重点在白衬衫透感、内搭替代、裤型显矮、春季亮色和牛仔搭配。",
                "商业含义": "用户在意的是可穿性和解决方案，而不是抽象风格标签。",
            },
            {
                "来源": "日本论坛问答",
                "核心发现": f"有效帖子 {len(forum_df)} 条，话题集中在办公室休闲边界、阔腿裤身高适配、学生预算、学院风理解。",
                "商业含义": "四月要提前解决用户顾虑，不然搜索热度不会直接转成成交。",
            },
            {
                "来源": "真实购买评论",
                "核心发现": f"有效评论页摘要 {len(review_df)} 条，反复出现免烫、舒适、长度刚好、好动、价格合理。",
                "商业含义": "真实购买评价支持“功能、舒适、价格带”优先于纯风格叙事。",
            },
        ]
    )


def build_final_recommendation() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"层级": "主线风格", "结论": "办公室休闲 + 轻美式 + 简约利落", "动作": "继续作为品牌主线。"},
            {"层级": "核心商品", "结论": "衬衫 / 牛仔裤 / 西裤 / 轻套装 / 短夹克", "动作": "四月主力库存放在这 5 个方向。"},
            {"层级": "功能卖点", "结论": "免烫 / 易护理 / 弹力 / 防泼水", "动作": "详情页与广告优先写这些，而不是先写风格。"},
            {"层级": "裤型", "结论": "宽直筒、双褶、轻阔腿比极宽更稳", "动作": "控制裤宽和裤长，降低拖地风险。"},
            {"层级": "学院风", "结论": "适合做轻学院胶囊，不适合做主线", "动作": "用紺ブレ、V 领开衫、牛津衬衫做小系列。"},
            {"层级": "谨慎方向", "结论": "重街头 / 重古着 / 工装裤主推 / 过早凉感", "动作": "可以点缀，不适合四月主推。"},
        ]
    )


def write_markdown(
    audit_df: pd.DataFrame,
    sample_df: pd.DataFrame,
    academy_df: pd.DataFrame,
) -> None:
    lines = [
        "# 日本男装第二步：多来源验证 v3",
        "",
        f"- 分析日期：{EXPORT_DATE}",
        "- 本版新增：男装适配检查、样本充分性评估、评论/帖子中文翻译、第一步低优先级联动。",
        "",
        "## 男装适配检查",
        "",
    ]
    for _, row in audit_df.iterrows():
        lines.append(f"- {row['来源']}：{row['处理结果']}")

    lines.extend(["", "## 样本量评估", ""])
    for _, row in sample_df.iterrows():
        lines.append(
            f"- {row['来源']}：当前 {row['当前样本量']}，建议 {row['建议样本量']}，"
            f"最低还需补充 {row['还需补充']}，状态 {row['是否足够']}。"
        )

    lines.extend(["", "## 学院风结论", ""])
    for _, row in academy_df.iterrows():
        lines.append(f"- {row['指标']}：{row['结果']}。{row['说明']}")

    OUTPUT_MD.write_text("\n".join(lines), encoding="utf-8")


def autosize(worksheet, dataframe: pd.DataFrame) -> None:
    for idx, col in enumerate(dataframe.columns):
        values = dataframe[col].astype(str).tolist()
        max_len = max([len(str(col))] + [len(v) for v in values])
        worksheet.set_column(idx, idx, min(max_len + 2, 80))


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
    function_df = load_csv(FUNCTION_PATH)
    scene_df = load_csv(SCENE_PATH)

    amazon_df = fetch_amazon_bestsellers(limit=15)
    yt_main_df, yt_aux_df = fetch_youtube_searches()
    comment_videos = select_comment_videos(yt_main_df)
    yt_comments_df = fetch_youtube_comments(comment_videos, limit_per_video=10)
    forum_df = build_forum_df(limit_per_query=2)
    review_df = build_review_df(limit_per_query=2)

    yt_comments_df = add_translation_column(yt_comments_df, "评论原文", "评论中文")
    forum_df = add_translation_column(forum_df, "帖子标题原文", "帖子标题中文")
    forum_df = add_translation_column(forum_df, "帖子摘要原文", "帖子摘要中文")
    review_df = add_translation_column(review_df, "评论页标题原文", "评论页标题中文")
    review_df = add_translation_column(review_df, "评论摘要原文", "评论摘要中文")

    audit_df = build_source_audit(amazon_df, yt_main_df, yt_aux_df, forum_df, review_df)
    sample_df = build_sample_assessment(amazon_df, yt_main_df, yt_comments_df, forum_df, review_df)
    first_step_df = build_first_step_reference(style_df, item_df, function_df, scene_df)
    summary_df = build_integrated_summary(style_df, item_df, function_df, amazon_df, yt_comments_df, forum_df, review_df)
    academy_df = build_academy_analysis(style_df, amazon_df, yt_main_df, forum_df, review_df)
    source_summary_df = build_source_summary(amazon_df, yt_main_df, forum_df, review_df)
    final_df = build_final_recommendation()

    write_markdown(audit_df, sample_df, academy_df)

    overview_df = pd.DataFrame(
        [
            {"项目": "本版核心修正", "内容": "YouTube 改成男装主来源优先，女性视角单独降级为辅助；评论、帖子和评论页都加了中文翻译。"},
            {"项目": "本版扩样规模", "内容": f"Amazon 当前榜单 {len(amazon_df)} 条；YouTube 男装视频 {len(yt_main_df)} 条；YouTube 男装评论 {len(yt_comments_df)} 条；论坛问答 {len(forum_df)} 条；真实购买评论 {len(review_df)} 条。"},
            {"项目": "第一步与第二步关系", "内容": "第一步作为低优先级趋势底图保留；第二步作为高优先级多来源验证决定最终动作。"},
            {"项目": "Amazon 历史 4 月说明", "内容": "官方历史畅销榜无法公开导出，因此仍保留“历史 4 月代理”口径，重点看季节性与评论内容。"},
        ]
    )

    with pd.ExcelWriter(OUTPUT_XLSX, engine="xlsxwriter") as writer:
        write_sheet(writer, "概览", overview_df)
        write_sheet(writer, "来源适配检查", audit_df)
        write_sheet(writer, "样本充分性评估", sample_df)
        write_sheet(writer, "第一步趋势参考_低优先", first_step_df)
        write_sheet(writer, "综合判断", summary_df)
        write_sheet(writer, "来源总结", source_summary_df)
        write_sheet(writer, "Amazon当前榜单", amazon_df, link_columns=["榜单链接"])
        write_sheet(writer, "YouTube主来源", yt_main_df, link_columns=["视频链接"])
        write_sheet(writer, "YouTube辅助来源", yt_aux_df, link_columns=["视频链接"])
        write_sheet(writer, "YouTube评论", yt_comments_df, link_columns=["视频链接"])
        write_sheet(writer, "论坛与问答", forum_df, link_columns=["链接"])
        write_sheet(writer, "真实购买评论", review_df, link_columns=["链接"])
        write_sheet(writer, "学院风专项", academy_df)
        write_sheet(writer, "最终建议", final_df)


if __name__ == "__main__":
    main()
