from __future__ import annotations

import html
import json
import shutil
from pathlib import Path

import pandas as pd
from PIL import Image, ImageDraw
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
STEP6_DIR = ROOT / "step6_output"
EXPORT_DATE = "2026-03-23"

SOURCE_PACK_DIR = next(
    path for path in STEP6_DIR.iterdir() if path.is_dir() and path.name == "T01_白T素材包"
)
XIANG_DIR = next(path for path in SOURCE_PACK_DIR.rglob("*") if path.is_dir() and path.name == "象山阿姆达服饰有限公司")
XIANG_MAIN = XIANG_DIR / "主图细节图"
CHAO_DIR = next(path for path in SOURCE_PACK_DIR.rglob("*") if path.is_dir() and "潮克堡" in path.name)
CHAO_MAIN = CHAO_DIR / "主图细节图"

OUT_DIR = STEP6_DIR / f"T01_Oakln复刻版_{EXPORT_DATE}"
ASSET_DIR = OUT_DIR / "assets"
OUT_DIR.mkdir(parents=True, exist_ok=True)
ASSET_DIR.mkdir(parents=True, exist_ok=True)

HTML_PATH = OUT_DIR / "index.html"
SCREENSHOT_PATH = OUT_DIR / "preview_full.png"
SUMMARY_MD = STEP6_DIR / f"T01_Oakln复刻版说明_{EXPORT_DATE}.md"
SUMMARY_XLSX = STEP6_DIR / f"T01_Oakln复刻版素材_{EXPORT_DATE}.xlsx"
SUMMARY_JSON = STEP6_DIR / f"T01_Oakln复刻版素材_{EXPORT_DATE}.json"


def find_file(pattern: str) -> Path:
    matches = sorted(XIANG_MAIN.glob(pattern))
    if not matches:
        raise FileNotFoundError(pattern)
    return matches[0]


def find_chao_file(pattern: str) -> Path:
    matches = sorted(CHAO_MAIN.glob(pattern))
    if not matches:
        raise FileNotFoundError(pattern)
    return matches[0]


source_map = {
    "hero_white": find_file("主图细节图_12_*.jpg"),
    "hero_hanger": find_file("主图细节图_02_*.jpg"),
    "detail_collar": find_file("主图细节图_14_*.jpg"),
    "detail_fold": find_file("主图细节图_18_*.jpg"),
    "detail_fold_alt": find_file("主图细节图_10_*.jpg"),
    "detail_rack": find_file("主图细节图_09_*.jpg"),
    "color_multi": find_file("主图细节图_15_*.jpg"),
    "color_black": find_file("主图细节图_08_*.jpg"),
    "color_gray": find_file("主图细节图_05_*.jpg"),
    "color_olive": find_file("主图细节图_03_*.jpg"),
    "color_beige": find_file("主图细节图_20_*.jpg"),
    "model_selfie": find_chao_file("主图细节图_06_*.jpg"),
}


def save_clean_variants(asset_map: dict[str, str]) -> None:
    hero_clean = ASSET_DIR / "hero_hanger_clean.jpg"
    hero_img = Image.open(source_map["hero_hanger"]).convert("RGB")
    hero_crop = hero_img.crop((int(hero_img.width * 0.24), 0, hero_img.width, hero_img.height))
    hero_crop.save(hero_clean, quality=95)
    asset_map["hero_hanger_clean"] = f"assets/{hero_clean.name}"

    multi_clean = ASSET_DIR / "color_multi_clean.jpg"
    multi_img = Image.open(source_map["color_multi"]).convert("RGB")
    cleaned = multi_img.copy()
    ImageDraw.Draw(cleaned).rectangle((0, 0, cleaned.width, 150), fill=(255, 255, 255))
    cleaned.save(multi_clean, quality=95)
    asset_map["color_multi_clean"] = f"assets/{multi_clean.name}"

    model_clean = ASSET_DIR / "model_crop.jpg"
    model_img = Image.open(source_map["model_selfie"]).convert("RGB")
    # Keep only the torso area to preserve the garment while reducing自拍感.
    box = (
        int(model_img.width * 0.10),
        int(model_img.height * 0.12),
        int(model_img.width * 0.88),
        int(model_img.height * 0.92),
    )
    model_crop = model_img.crop(box)
    model_crop.save(model_clean, quality=95)
    asset_map["model_crop"] = f"assets/{model_clean.name}"


def copy_assets() -> dict[str, str]:
    asset_map: dict[str, str] = {}
    for key, source in source_map.items():
        out_name = f"{key}{source.suffix.lower()}"
        target = ASSET_DIR / out_name
        shutil.copy2(source, target)
        asset_map[key] = f"assets/{out_name}"
    save_clean_variants(asset_map)
    return asset_map


assets = copy_assets()

gallery_items = [
    ("hero_white", "正面主图"),
    ("model_crop", "上身展示"),
    ("hero_hanger_clean", "挂拍陈列"),
    ("detail_fold_alt", "面料折叠"),
    ("detail_collar", "领口细节"),
    ("detail_rack", "排挂近景"),
    ("detail_fold", "折叠近景"),
    ("color_multi_clean", "多色挂拍"),
    ("color_gray", "灰色平铺"),
]

recommend_items = [
    ("hero_white", "300g重磅コットン クルーネックTシャツ", "ホワイト", "¥6,980"),
    ("detail_fold_alt", "300g重磅コットン クルーネックTシャツ", "ブラック", "¥6,980"),
    ("color_gray", "300g重磅コットン クルーネックTシャツ", "グレー", "¥6,980"),
]

copy_rows = [
    {"字段": "日文标题", "内容": "300g重磅コットン クルーネックTシャツ メンズ 半袖 厚手 透けにくい ベーシック"},
    {"字段": "价格", "内容": "¥6,980"},
    {"字段": "副标题", "内容": "肉感のある300gコットンで、白でも透け感を抑えやすいベーシックTシャツ。"},
    {"字段": "卖点1标题", "内容": "輪郭がきれいに見える、ベーシックな主力白T"},
    {"字段": "卖点1正文", "内容": "一枚で着てもシルエットが崩れにくい、重みのあるコットンT。4月の主力カットソーとして提案しやすい。"},
    {"字段": "卖点2标题", "内容": "クルーネックの細部を整え、首回りがすっきり見える"},
    {"字段": "卖点2正文", "内容": "リブと縫製の見え方を重視し、ベーシックながら安っぽく見えにくいディテール感に寄せている。"},
    {"字段": "卖点3标题", "内容": "重みのある綿感で、白でも透けにくい見え方へ"},
    {"字段": "卖点3正文", "内容": "300gの厚手コットンをベースに、春の単品着用でも安心感が出やすい構成。"},
    {"字段": "描述", "内容": "肉感のある300gコットンを使用したクルーネックTシャツ。ゆったりとしたベーシックなシルエットで、春の立ち上がりから初夏まで長く活躍します。"},
    {"字段": "规格", "内容": "素材：コットン / 生地感：300g厚手 / サイズ：XS, S, M, L, XL, XXL, XXXL / カラー：ホワイト, ブラック, グレー, オリーブ, ベージュ"},
]

copy_map = {row["字段"]: row["内容"] for row in copy_rows}


def esc(text: str) -> str:
    return html.escape(text)


def build_html() -> str:
    gallery_html = "\n".join(
        f"""
        <figure class="gallery-card">
          <img src="{assets[key]}" alt="{esc(label)}">
        </figure>
        """
        for key, label in gallery_items
    )

    recommend_html = "\n".join(
        f"""
        <article class="recommend-card">
          <img src="{assets[key]}" alt="{esc(title)}">
          <div class="recommend-meta">
            <h4>{esc(title)}</h4>
            <p>{esc(color)}</p>
            <strong>{esc(price)}</strong>
          </div>
        </article>
        """
        for key, title, color, price in recommend_items
    )

    html_text = f"""<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{esc(copy_map['日文标题'])} – Oakln</title>
  <style>
    * {{ box-sizing: border-box; }}
    body {{ margin: 0; font-family: "Hiragino Sans", "Yu Gothic", "Microsoft YaHei", sans-serif; color: #111; background: #fff; }}
    a {{ color: inherit; text-decoration: none; }}
    header {{ border-bottom: 1px solid #e8e8e8; background: #fff; }}
    .topbar {{ max-width: 1200px; margin: 0 auto; padding: 18px 20px 14px; display: flex; align-items: center; gap: 28px; }}
    .logo {{ font-size: 40px; font-weight: 700; font-family: Georgia, serif; letter-spacing: .5px; }}
    .nav {{ display: flex; gap: 24px; font-size: 12px; color: #222; }}
    .icons {{ margin-left: auto; display: flex; gap: 16px; font-size: 22px; color: #333; }}
    .page {{ max-width: 1200px; margin: 0 auto; padding: 18px 20px 80px; }}
    .breadcrumb {{ margin: 0 0 18px; font-size: 12px; color: #666; }}
    .hero {{ display: grid; grid-template-columns: minmax(0, 730px) 350px; gap: 34px; align-items: start; }}
    .gallery-grid {{ column-count: 2; column-gap: 22px; }}
    .gallery-card {{ margin: 0 0 22px; background: #f7f7f7; break-inside: avoid; }}
    .gallery-card img {{ display: block; width: 100%; height: auto; }}
    .summary {{ position: sticky; top: 20px; }}
    .summary h1 {{ margin: 0 0 10px; font-size: 18px; line-height: 1.65; font-weight: 700; letter-spacing: .01em; }}
    .price-row {{ margin: 10px 0 6px; display: flex; gap: 10px; align-items: baseline; }}
    .price-row strong {{ font-size: 32px; letter-spacing: .01em; }}
    .price-row span {{ color: #666; text-decoration: line-through; }}
    .point-copy {{ font-size: 12px; color: #555; margin: 0 0 10px; }}
    .promo-list {{ display: grid; gap: 6px; margin: 0 0 16px; font-size: 12px; color: #444; }}
    .note {{ color: #666; font-size: 12px; line-height: 1.8; margin-bottom: 18px; }}
    .section-label {{ font-size: 12px; color: #555; margin: 18px 0 8px; }}
    .swatches, .sizes {{ display: flex; flex-wrap: wrap; gap: 8px; }}
    .swatch {{ width: 22px; height: 22px; border-radius: 50%; border: 1px solid #bdbdbd; }}
    .size {{ min-width: 42px; padding: 9px 12px; border: 1px solid #d7d7d7; font-size: 12px; text-align: center; background: #fff; }}
    .qty-row {{ margin-top: 16px; display: flex; align-items: center; gap: 8px; }}
    .qty-btn, .qty-value {{ width: 34px; height: 34px; border: 1px solid #d7d7d7; background: #fff; display: grid; place-items: center; font-size: 13px; }}
    .thumb-row {{ display: flex; gap: 8px; margin-top: 16px; }}
    .thumb-row img {{ width: 64px; height: 64px; object-fit: cover; border: 1px solid #e4e4e4; display: block; }}
    .cta-primary {{ width: 100%; background: #0c2d57; color: #fff; padding: 13px 16px; border: 0; font-size: 14px; font-weight: 700; margin-top: 18px; }}
    .cta-secondary {{ width: 100%; background: #fff; color: #111; padding: 12px 16px; border: 1px solid #cfcfcf; font-size: 14px; margin-top: 10px; }}
    .meta-box {{ margin-top: 18px; padding-top: 16px; border-top: 1px solid #ececec; color: #555; font-size: 12px; line-height: 1.9; }}
    .feature-triple {{ margin-top: 52px; display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 24px; }}
    .feature-triple article img {{ width: 100%; display: block; }}
    .feature-triple h3 {{ font-size: 15px; line-height: 1.55; margin: 14px 0 10px; font-weight: 700; }}
    .feature-triple p {{ margin: 0; font-size: 13px; line-height: 1.85; color: #333; }}
    .image-overlay {{ position: relative; }}
    .overlay-tags {{ position: absolute; left: 18px; bottom: 18px; display: flex; gap: 10px; flex-wrap: wrap; }}
    .overlay-tags span {{ background: rgba(255,255,255,.86); border: 1px solid rgba(255,255,255,.92); padding: 8px 12px; font-size: 12px; }}
    .tabs {{ margin-top: 40px; border-top: 1px solid #e5e5e5; }}
    .tab-head {{ display: flex; gap: 24px; padding: 14px 0; border-bottom: 1px solid #e5e5e5; font-size: 14px; }}
    .tab-head strong {{ font-weight: 700; }}
    .desc {{ padding-top: 18px; font-size: 13px; line-height: 1.95; color: #222; }}
    .specs {{ margin-top: 16px; }}
    .specs div {{ display: flex; gap: 16px; }}
    .specs dt {{ width: 84px; color: #666; flex: 0 0 84px; }}
    .recommend {{ margin-top: 54px; }}
    .recommend h2 {{ margin: 0 0 20px; font-size: 20px; }}
    .recommend-grid {{ display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 22px; }}
    .recommend-card {{ border: 1px solid #efefef; background: #fff; }}
    .recommend-card img {{ display: block; width: 100%; height: auto; }}
    .recommend-meta {{ padding: 14px; }}
    .recommend-meta h4 {{ margin: 0 0 8px; font-size: 15px; line-height: 1.5; }}
    .recommend-meta p {{ margin: 0 0 8px; font-size: 13px; color: #666; }}
    .recommend-meta strong {{ font-size: 18px; }}
    footer {{ margin-top: 80px; padding: 40px 0 60px; border-top: 1px solid #efefef; color: #666; font-size: 13px; }}
    @media (max-width: 980px) {{
      .hero {{ grid-template-columns: 1fr; }}
      .summary {{ position: static; }}
      .feature-triple, .recommend-grid {{ grid-template-columns: 1fr; }}
      .gallery-grid {{ column-count: 1; }}
    }}
  </style>
</head>
<body>
  <header>
    <div class="topbar">
      <div class="logo">Oakln</div>
      <nav class="nav">
        <a href="#">商品列表</a>
        <a href="#">热门商品</a>
        <a href="#">新品上市</a>
        <a href="#">特别专题</a>
        <a href="#">客户支持</a>
      </nav>
      <div class="icons"><span>⌕</span><span>◌</span><span>👜</span></div>
    </div>
  </header>

  <main class="page">
    <div class="breadcrumb">首页 / 男士T恤 / 300g重磅白T</div>
    <section class="hero">
      <div class="gallery-grid">
        {gallery_html}
      </div>

      <aside class="summary">
        <h1>{esc(copy_map['日文标题'])}</h1>
        <div class="price-row">
          <strong>{esc(copy_map['价格'])}</strong>
          <span>¥7,980</span>
        </div>
        <div class="point-copy">698ポイント獲得予定 / 会員登録で初回クーポン配布</div>
        <div class="promo-list">
          <div>全商品2点で15%OFF</div>
          <div>全商品3点で20%OFF</div>
          <div>全商品4点で25%OFF</div>
        </div>
        <div class="note">{esc(copy_map['副标题'])}</div>

        <div class="section-label">カラー</div>
        <div class="swatches">
          <span class="swatch" style="background:#ffffff"></span>
          <span class="swatch" style="background:#111111"></span>
          <span class="swatch" style="background:#b9b9b9"></span>
          <span class="swatch" style="background:#575f47"></span>
          <span class="swatch" style="background:#e9dec8"></span>
        </div>

        <div class="section-label">サイズ</div>
        <div class="sizes">
          <span class="size">XS</span><span class="size">S</span><span class="size">M</span><span class="size">L</span><span class="size">XL</span><span class="size">2XL</span><span class="size">3XL</span>
        </div>

        <div class="qty-row">
          <span class="qty-btn">-</span><span class="qty-value">1</span><span class="qty-btn">+</span>
        </div>

        <div class="thumb-row">
          <img src="{assets['hero_white']}" alt="thumb1">
          <img src="{assets['hero_hanger_clean']}" alt="thumb2">
          <img src="{assets['detail_collar']}" alt="thumb3">
        </div>

        <button class="cta-primary">カートに追加</button>
        <button class="cta-secondary">今すぐ購入</button>

        <div class="meta-box">
          <div>LINE友だち追加でクーポン配布</div>
          <div>30日間返品交換ポリシー</div>
          <div>11,000円以上で送料無料</div>
          <div>2-5営業日以内に発送</div>
        </div>
      </aside>
    </section>

    <section class="feature-triple">
      <article>
        <div class="image-overlay"><img src="{assets['model_crop']}" alt="feature1"></div>
        <h3>{esc(copy_map['卖点1标题'])}</h3>
        <p>{esc(copy_map['卖点1正文'])}</p>
      </article>
      <article>
        <div class="image-overlay">
          <img src="{assets['detail_collar']}" alt="feature2">
          <div class="overlay-tags"><span>クルーネック</span><span>首回りすっきり</span></div>
        </div>
        <h3>{esc(copy_map['卖点2标题'])}</h3>
        <p>{esc(copy_map['卖点2正文'])}</p>
      </article>
      <article>
        <div class="image-overlay">
          <img src="{assets['detail_fold']}" alt="feature3">
          <div class="overlay-tags"><span>300g</span><span>透けにくい見え方</span></div>
        </div>
        <h3>{esc(copy_map['卖点3标题'])}</h3>
        <p>{esc(copy_map['卖点3正文'])}</p>
      </article>
    </section>

    <section class="tabs">
      <div class="tab-head">
        <strong>商品説明</strong>
        <span>配送、返品と交換</span>
      </div>
      <div class="desc">
        <p>{esc(copy_map['描述'])}</p>
        <dl class="specs">
          <div><dt>産地</dt><dd>中国</dd></div>
          <div><dt>素材</dt><dd>コットン</dd></div>
          <div><dt>サイズ</dt><dd>XS, S, M, L, XL, 2XL, 3XL</dd></div>
          <div><dt>カラー</dt><dd>ホワイト, ブラック, グレー, オリーブ, ベージュ</dd></div>
          <div><dt>洗濯</dt><dd>手洗い推奨 / ネット使用での洗濯機洗いは要確認</dd></div>
        </dl>
      </div>
    </section>

    <section class="recommend">
      <h2>スタッフ服装</h2>
      <div class="recommend-grid">
        {recommend_html}
      </div>
    </section>

    <footer>
      <div>Oakln スタイルの復刻プレビュー。衣服だけを白Tに差し替えた検証用ローカルページです。</div>
    </footer>
  </main>
</body>
</html>
"""
    return html_text


def render_preview() -> dict[str, object]:
    html_text = build_html()
    HTML_PATH.write_text(html_text, encoding="utf-8")

    result: dict[str, object] = {"html": str(HTML_PATH), "screenshot": str(SCREENSHOT_PATH), "broken": []}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1440, "height": 2400})
        page.goto(HTML_PATH.resolve().as_uri(), wait_until="load", timeout=120000)
        page.wait_for_timeout(1500)
        broken: list[str] = []
        images = page.locator("img")
        for idx in range(images.count()):
            ok = images.nth(idx).evaluate("(img) => img.complete && img.naturalWidth > 0")
            src = images.nth(idx).get_attribute("src") or ""
            if not ok:
                broken.append(src)
        page.screenshot(path=str(SCREENSHOT_PATH), full_page=True)
        browser.close()
    result["broken"] = broken
    return result


render_result = render_preview()

asset_rows = [
    {
        "模块": "左侧图库",
        "素材键": key,
        "相对路径": assets[key],
        "来源文件": str(source_map[key]) if key in source_map else "由同款原图净化生成",
    }
    for key, _ in gallery_items
]
asset_rows.extend(
    [
        {"模块": "三联卖点", "素材键": "hero_white", "相对路径": assets["hero_white"], "来源文件": str(source_map["hero_white"])},
        {"模块": "三联卖点", "素材键": "detail_collar", "相对路径": assets["detail_collar"], "来源文件": str(source_map["detail_collar"])},
        {"模块": "三联卖点", "素材键": "detail_fold", "相对路径": assets["detail_fold"], "来源文件": str(source_map["detail_fold"])},
    ]
)

structure_rows = [
    {"区域": "顶部", "复刻内容": "Oakln式两栏商品详情页", "状态": "已完成"},
    {"区域": "左侧", "复刻内容": "2列商品图网格", "状态": "已完成"},
    {"区域": "右侧", "复刻内容": "标题、价格、颜色、尺码、按钮", "状态": "已完成"},
    {"区域": "中段", "复刻内容": "三联卖点图文模块", "状态": "已完成"},
    {"区域": "说明区", "复刻内容": "商品说明与规格区", "状态": "已完成"},
    {"区域": "底部", "复刻内容": "推荐商品区", "状态": "已完成"},
]

with pd.ExcelWriter(SUMMARY_XLSX, engine="openpyxl") as writer:
    pd.DataFrame(structure_rows).to_excel(writer, sheet_name="复刻结构", index=False)
    pd.DataFrame(asset_rows).to_excel(writer, sheet_name="素材映射", index=False)
    pd.DataFrame(copy_rows).to_excel(writer, sheet_name="日文文案", index=False)
    pd.DataFrame(
        [
            {"文件": str(HTML_PATH), "用途": "本地复刻版详情页"},
            {"文件": str(SCREENSHOT_PATH), "用途": "全页截图校验"},
        ]
    ).to_excel(writer, sheet_name="输出文件", index=False)

summary = {
    "out_dir": str(OUT_DIR),
    "html": str(HTML_PATH),
    "screenshot": str(SCREENSHOT_PATH),
    "broken_images": render_result["broken"],
    "asset_count": len(asset_rows),
}
SUMMARY_JSON.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

SUMMARY_MD.write_text(
    "\n".join(
        [
            "# T01 Oakln复刻版",
            "",
            f"- 目录：`{OUT_DIR}`",
            f"- 本地页面：`{HTML_PATH}`",
            f"- 全页截图：`{SCREENSHOT_PATH}`",
            f"- 坏图数量：`{len(render_result['broken'])}`",
            "- 复刻标准：页面结构、右侧购买区、三联卖点区、说明区、推荐区与 Oakln 参考页对齐，只更换为白T商品素材。",
        ]
    ),
    encoding="utf-8",
)
