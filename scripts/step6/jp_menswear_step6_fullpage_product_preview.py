from __future__ import annotations

from pathlib import Path
from typing import Iterable

from PIL import Image, ImageDraw, ImageEnhance, ImageFilter, ImageFont
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
STEP6_DIR = ROOT / "step6_output"
MATERIALS_ROOT = STEP6_DIR / "素材包"
BUNDLE_DIR = MATERIALS_ROOT / "T01_bundle_2026-03-24"
ROUND1_DIR = MATERIALS_ROOT / "T01_round1_2026-03-24" / "01_round1_generated"
ROUND3_DIR = MATERIALS_ROOT / "T01_round3_2026-03-24"
BUNDLE_UPLOAD_DIR = BUNDLE_DIR / "01_upload_pack_v2"
WEB_PREVIEW_DIRS = [
    ROUND3_DIR / "04_web_preview",
    BUNDLE_DIR / "04_web_preview",
]

FONT_PATHS = [
    Path(r"C:\Windows\Fonts\msyh.ttc"),
    Path(r"C:\Windows\Fonts\msyhbd.ttc"),
]


def pick(path: Path, pattern: str) -> Path:
    matches = sorted(path.glob(pattern))
    if not matches:
        raise FileNotFoundError(f"Missing file for pattern: {path}\\{pattern}")
    return matches[0]


def load_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    preferred = FONT_PATHS[1] if bold else FONT_PATHS[0]
    fallback = FONT_PATHS[0]
    for path in [preferred, fallback]:
        if path.exists():
            return ImageFont.truetype(str(path), size=size)
    return ImageFont.load_default()


def cover_resize(image: Image.Image, size: tuple[int, int]) -> Image.Image:
    target_w, target_h = size
    src_w, src_h = image.size
    scale = max(target_w / src_w, target_h / src_h)
    resized = image.resize((int(src_w * scale), int(src_h * scale)), Image.Resampling.LANCZOS)
    left = (resized.width - target_w) // 2
    top = (resized.height - target_h) // 2
    return resized.crop((left, top, left + target_w, top + target_h))


def add_bottom_gradient(image: Image.Image, strength: int = 120) -> Image.Image:
    base = image.convert("RGBA")
    overlay = Image.new("RGBA", base.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    width, height = base.size
    for idx in range(height):
        alpha = int(max(0, strength * ((idx / height) ** 2.2)))
        draw.line((0, idx, width, idx), fill=(0, 0, 0, alpha))
    return Image.alpha_composite(base, overlay)


def build_benefit_card(
    source: Path,
    dest: Path,
    badge: str,
    title: str,
    subtitle: str,
) -> None:
    base = Image.open(source).convert("RGB")
    canvas = cover_resize(base, (900, 900))
    canvas = ImageEnhance.Contrast(canvas).enhance(1.06)
    canvas = ImageEnhance.Brightness(canvas).enhance(0.96)
    canvas = add_bottom_gradient(canvas, strength=165).convert("RGBA")

    draw = ImageDraw.Draw(canvas)
    badge_font = load_font(34, bold=True)
    title_font = load_font(40, bold=True)
    sub_font = load_font(24)

    badge_box = (56, 70, 230, 126)
    draw.rounded_rectangle(badge_box, radius=22, fill=(255, 255, 255, 40), outline=(255, 255, 255, 90), width=2)
    draw.text((badge_box[0] + 24, badge_box[1] + 10), badge, font=badge_font, fill=(255, 255, 255, 230))

    draw.text((60, 680), title, font=title_font, fill=(255, 255, 255, 245))
    draw.text((60, 740), subtitle, font=sub_font, fill=(255, 255, 255, 220), spacing=8)

    canvas.convert("RGB").save(dest, quality=94)


def build_ref_icon_card(
    source: Path,
    dest: Path,
    headline: str,
    subline: str,
    icon_label: str,
) -> None:
    base = Image.open(source).convert("RGB")
    canvas = cover_resize(base.filter(ImageFilter.GaussianBlur(radius=0.2)), (900, 900))
    canvas = add_bottom_gradient(canvas, strength=180).convert("RGBA")
    draw = ImageDraw.Draw(canvas)
    headline_font = load_font(44, bold=True)
    sub_font = load_font(24)
    icon_font = load_font(26, bold=True)

    center_x = canvas.width // 2
    circle_box = (center_x - 56, 530, center_x + 56, 642)
    draw.ellipse(circle_box, outline=(255, 255, 255, 200), width=4)
    label_box = draw.textbbox((0, 0), icon_label, font=icon_font)
    label_w = label_box[2] - label_box[0]
    label_h = label_box[3] - label_box[1]
    draw.text((center_x - label_w / 2, 586 - label_h / 2), icon_label, font=icon_font, fill=(255, 255, 255, 235))

    headline_box = draw.textbbox((0, 0), headline, font=headline_font)
    headline_w = headline_box[2] - headline_box[0]
    draw.text((center_x - headline_w / 2, 680), headline, font=headline_font, fill=(255, 255, 255, 245))

    sub_box = draw.multiline_textbbox((0, 0), subline, font=sub_font, spacing=6, align="center")
    sub_w = sub_box[2] - sub_box[0]
    draw.multiline_text(
        (center_x - sub_w / 2, 748),
        subline,
        font=sub_font,
        fill=(255, 255, 255, 220),
        spacing=6,
        align="center",
    )
    canvas.convert("RGB").save(dest, quality=94)


def save_cover_copy(source: Path, dest: Path, size: tuple[int, int] = (900, 900)) -> None:
    image = Image.open(source).convert("RGB")
    cover_resize(image, size).save(dest, quality=94)


def crop_title_header(source: Path, dest: Path, trim_ratio: float = 0.18) -> None:
    image = Image.open(source).convert("RGB")
    trim_top = int(image.height * trim_ratio)
    cropped = image.crop((0, trim_top, image.width, image.height))
    cover_resize(cropped, (900, 900)).save(dest, quality=94)


def ensure_preview_assets(target_dir: Path) -> None:
    assets_dir = target_dir / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)

    save_cover_copy(pick(BUNDLE_UPLOAD_DIR / "01_商品主图_按上传顺序", "04_*"), assets_dir / "top_02_collar.jpg")
    save_cover_copy(pick(BUNDLE_UPLOAD_DIR / "01_商品主图_按上传顺序", "03_*"), assets_dir / "top_04_hanging.jpg")
    save_cover_copy(pick(BUNDLE_UPLOAD_DIR / "01_商品主图_按上传顺序", "05_*"), assets_dir / "top_06_fold.jpg")
    crop_title_header(pick(ROUND1_DIR, "G05_*.jpg"), assets_dir / "top_05_stack.jpg")
    crop_title_header(pick(ROUND1_DIR, "G07_*.jpg"), assets_dir / "top_07_scatter.jpg")
    crop_title_header(pick(ROUND1_DIR, "G09_*.jpg"), assets_dir / "tail_01_cuff.jpg")
    save_cover_copy(pick(BUNDLE_UPLOAD_DIR / "04_推荐商品图", "02_*"), assets_dir / "tail_02_black_flat.jpg")
    save_cover_copy(pick(BUNDLE_UPLOAD_DIR / "01_商品主图_按上传顺序", "01_*"), assets_dir / "tail_03_white_flat.jpg")
    save_cover_copy(pick(BUNDLE_UPLOAD_DIR / "01_商品主图_按上传顺序", "08_*"), assets_dir / "tail_04_gray_flat.jpg")
    save_cover_copy(pick(ROUND1_DIR, "G10_米白*.jpg"), assets_dir / "tail_05_ivory_flat.jpg")
    save_cover_copy(pick(ROUND1_DIR, "G10_军绿*.jpg"), assets_dir / "tail_06_olive_flat.jpg")
    save_cover_copy(pick(ROUND1_DIR, "G10_黑色*.jpg"), assets_dir / "tail_07_black_flat.jpg")
    save_cover_copy(pick(ROUND1_DIR, "G10_灰色*.jpg"), assets_dir / "tail_08_gray_flat.jpg")

    build_benefit_card(
        source=pick(ROUND1_DIR, "G13_*.jpg"),
        dest=assets_dir / "benefit_thick.jpg",
        badge="300G",
        title="厚实纯棉，不易透",
        subtitle="高密度面料更利落，单穿也能保持干净轮廓",
    )
    build_ref_icon_card(
        source=pick(ROUND1_DIR, "G08_*.jpg"),
        dest=assets_dir / "benefit_collar.jpg",
        headline="领口稳定",
        subline="洗后不易松垮\n日常机洗也更省心",
        icon_label="领",
    )
    build_ref_icon_card(
        source=pick(ROUND1_DIR, "G14_*.jpg"),
        dest=assets_dir / "benefit_texture.jpg",
        headline="柔软肌理",
        subline="贴肤不扎\n日常通勤和内搭都顺手",
        icon_label="柔",
    )


def top_gallery(relative_mode: str) -> list[str]:
    g = "../01_generated" if relative_mode == "round3" else "../../T01_round3_2026-03-24/01_generated"
    a = "./assets"
    return [
        f"{g}/G02.png",
        f"{a}/top_02_collar.jpg",
        f"{g}/G03.png",
        f"{a}/top_04_hanging.jpg",
        f"{a}/top_05_stack.jpg",
        f"{a}/top_06_fold.jpg",
        f"{a}/top_07_scatter.jpg",
        f"{a}/top_02_collar.jpg",
    ]


def tail_gallery(relative_mode: str) -> list[str]:
    a = "./assets"
    return [
        f"{a}/tail_01_cuff.jpg",
        f"{a}/tail_02_black_flat.jpg",
        f"{a}/tail_03_white_flat.jpg",
        f"{a}/tail_04_gray_flat.jpg",
        f"{a}/tail_05_ivory_flat.jpg",
        f"{a}/tail_06_olive_flat.jpg",
        f"{a}/tail_07_black_flat.jpg",
        f"{a}/tail_08_gray_flat.jpg",
    ]


def build_html(relative_mode: str) -> str:
    if relative_mode == "round3":
        g = "../01_generated"
        a = "./assets"
    else:
        g = "../../T01_round3_2026-03-24/01_generated"
        a = "./assets"

    gallery_items = top_gallery(relative_mode)
    tail_items = tail_gallery(relative_mode)

    def img_tags(items: Iterable[str], cls: str) -> str:
        return "\n".join(f'        <img class="{cls}" src="{src}" alt="preview">' for src in items)

    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>T01 300g 重磅白T 商品页预览</title>
  <style>
    * {{
      box-sizing: border-box;
    }}
    body {{
      margin: 0;
      font-family: "Microsoft YaHei", "PingFang SC", sans-serif;
      color: #1b1b1b;
      background: #f7f4ee;
    }}
    img {{
      display: block;
      width: 100%;
    }}
    .sitebar {{
      height: 8px;
      background: #102c4e;
    }}
    .header {{
      background: #fff;
      border-bottom: 1px solid #e8e1d6;
    }}
    .header-inner {{
      width: min(1200px, calc(100vw - 52px));
      margin: 0 auto;
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 14px 0 12px;
    }}
    .brand {{
      font-family: Georgia, "Times New Roman", serif;
      font-size: 24px;
      font-weight: 700;
      letter-spacing: -0.03em;
    }}
    .nav, .toolbar {{
      display: flex;
      gap: 18px;
      color: #6c665d;
      font-size: 11px;
    }}
    .wrap {{
      width: min(1200px, calc(100vw - 52px));
      margin: 0 auto;
      padding: 16px 0 80px;
    }}
    .breadcrumb {{
      margin-bottom: 14px;
      color: #91897f;
      font-size: 11px;
    }}
    .top {{
      display: grid;
      grid-template-columns: minmax(0, 1.72fr) minmax(340px, 0.88fr);
      gap: 26px;
      align-items: start;
    }}
    .gallery {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 16px;
    }}
    .gallery img,
    .tail-gallery img {{
      aspect-ratio: 1 / 1;
      object-fit: cover;
      background: #fff;
    }}
    .tail-wrap {{
      width: calc((100% - 26px) * 0.661);
      margin-top: 18px;
    }}
    .tail-gallery {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 16px;
    }}
    .buy {{
      position: sticky;
      top: 18px;
      background: #fbf8f3;
      border: 1px solid #e6dfd4;
      padding: 18px 18px 16px;
    }}
    h1 {{
      margin: 0 0 12px;
      font-size: 27px;
      line-height: 1.45;
      font-weight: 700;
    }}
    .price {{
      font-size: 42px;
      font-weight: 700;
      margin-bottom: 8px;
    }}
    .price span {{
      margin-left: 10px;
      color: #8e8579;
      text-decoration: line-through;
      font-size: 16px;
      font-weight: 400;
    }}
    .meta {{
      color: #8e8579;
      font-size: 11px;
      line-height: 1.8;
      margin-bottom: 14px;
    }}
    .option-title {{
      margin: 14px 0 8px;
      font-size: 12px;
      font-weight: 700;
    }}
    .swatches {{
      display: flex;
      gap: 7px;
      margin-bottom: 12px;
    }}
    .swatches span {{
      width: 14px;
      height: 14px;
      display: inline-block;
      border-radius: 50%;
      border: 1px solid #cfc6b9;
    }}
    .size-grid {{
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
      margin-bottom: 12px;
    }}
    .size-grid span {{
      min-width: 30px;
      padding: 6px 8px;
      font-size: 11px;
      text-align: center;
      border: 1px solid #d8cfbf;
      background: #fff;
    }}
    .thumbs {{
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 8px;
      margin: 12px 0 16px;
    }}
    .thumbs img {{
      aspect-ratio: 1 / 1;
      object-fit: cover;
      border: 1px solid #e6dfd4;
      background: #fff;
    }}
    .qty {{
      display: flex;
      align-items: center;
      gap: 8px;
      margin-bottom: 12px;
    }}
    .qty span {{
      width: 28px;
      height: 28px;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      border: 1px solid #d8cfbf;
      background: #fff;
      font-size: 12px;
    }}
    .btn {{
      display: block;
      width: 100%;
      padding: 11px 14px;
      margin-top: 10px;
      text-decoration: none;
      text-align: center;
      font-size: 13px;
      font-weight: 700;
      letter-spacing: 0.03em;
    }}
    .primary {{
      background: #17375d;
      color: #fff;
    }}
    .secondary {{
      border: 1px solid #17375d;
      color: #17375d;
      background: #fff;
    }}
    .section {{
      margin-top: 34px;
    }}
    .benefit-grid {{
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 16px;
    }}
    .benefit-card img {{
      aspect-ratio: 1 / 1;
      object-fit: cover;
      background: #fff;
    }}
    .benefit-title {{
      margin: 10px 0 6px;
      font-size: 14px;
      font-weight: 700;
      line-height: 1.45;
    }}
    .benefit-desc {{
      color: #514a42;
      font-size: 12px;
      line-height: 1.75;
    }}
    .desc-tabs {{
      display: flex;
      gap: 20px;
      padding-bottom: 10px;
      margin-bottom: 12px;
      border-bottom: 1px solid #ddd4c8;
      font-size: 13px;
    }}
    .desc-tabs span:first-child {{
      font-weight: 700;
    }}
    .specs {{
      color: #514a42;
      font-size: 12px;
      line-height: 1.95;
    }}
    .staff-title {{
      margin: 0 0 14px;
      font-size: 18px;
      font-weight: 700;
    }}
    .staff-grid {{
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 16px;
    }}
    .staff-card {{
      position: relative;
      background: #fff;
    }}
    .staff-card img {{
      aspect-ratio: 0.83 / 1;
      object-fit: cover;
    }}
    .staff-meta {{
      position: absolute;
      left: 14px;
      bottom: 12px;
      color: rgba(255,255,255,0.88);
      font-size: 11px;
      line-height: 1.6;
      text-shadow: 0 1px 2px rgba(0,0,0,0.18);
      white-space: pre-line;
    }}
    .footer {{
      margin-top: 56px;
      padding-top: 18px;
      border-top: 1px solid #ddd4c8;
      color: #857d71;
      font-size: 11px;
    }}
  </style>
</head>
<body>
  <div class="sitebar"></div>
  <header class="header">
    <div class="header-inner">
      <div class="brand">Oakln</div>
      <nav class="nav">
        <span>产品列表</span>
        <span>热门商品</span>
        <span>新品上市</span>
        <span>特别专题</span>
        <span>客户支持</span>
      </nav>
      <div class="toolbar">
        <span>搜索</span>
        <span>账户</span>
        <span>购物车</span>
      </div>
    </div>
  </header>

  <main class="wrap">
    <div class="breadcrumb">首页 / 男装 / T 恤 / 300g 重磅白T</div>

    <section class="top">
      <div class="gallery">
{img_tags(gallery_items, "gallery-slot")}
      </div>

      <aside class="buy">
        <h1>300g 重磅コットン クルーネックTシャツ<br>メンズ 半袖 厚手 透けにくい ベーシック</h1>
        <div class="price">¥6,980<span>¥7,980</span></div>
        <div class="meta">
          当前预览已按参照页重排主图区顺序。<br>
          核心卖点基于前面评论分析：厚实不透、领口稳定、可机洗好打理。
        </div>

        <div class="option-title">カラー</div>
        <div class="swatches">
          <span style="background:#f6f5f1"></span>
          <span style="background:#111111"></span>
          <span style="background:#b5b5b5"></span>
          <span style="background:#6f775e"></span>
          <span style="background:#dfd7c9"></span>
        </div>

        <div class="option-title">サイズ</div>
        <div class="size-grid">
          <span>S</span><span>M</span><span>L</span><span>XL</span><span>2XL</span>
        </div>

        <div class="option-title">数量</div>
        <div class="qty">
          <span>-</span><span>1</span><span>+</span>
        </div>

        <div class="option-title">缩略图</div>
        <div class="thumbs">
          <img src="{g}/G02.png" alt="thumb1">
          <img src="{g}/G03.png" alt="thumb2">
          <img src="{tail_items[0]}" alt="thumb3">
        </div>

        <a class="btn primary" href="javascript:void(0)">加入购物车</a>
        <a class="btn secondary" href="javascript:void(0)">立即购买</a>
      </aside>
    </section>

    <section class="tail-wrap">
      <div class="tail-gallery">
{img_tags(tail_items, "tail-slot")}
      </div>
    </section>

    <section class="section">
      <div class="benefit-grid">
        <article class="benefit-card">
          <img src="{g}/G12.png" alt="卖点1">
          <div class="benefit-title">厚实不透，单穿也有干净轮廓</div>
          <div class="benefit-desc">用户最在意的是白T会不会透、会不会像打底衫。这张位就按参照页用生活方式图，强调“单穿就成立”。</div>
        </article>
        <article class="benefit-card">
          <img src="{a}/benefit_collar.jpg" alt="卖点2">
          <div class="benefit-title">领口稳定，不易松垮</div>
          <div class="benefit-desc">用户评论里反复提到领口洗后容易变形，所以中间这张改成更接近参照页的结构特写卖点卡。</div>
        </article>
        <article class="benefit-card">
          <img src="{a}/benefit_texture.jpg" alt="卖点3">
          <div class="benefit-title">可机洗，柔软肌理更省心</div>
          <div class="benefit-desc">第三张位承接“好打理”和“贴肤舒适”，对应参照页里那种带文字覆盖的功能图。</div>
        </article>
      </div>
    </section>

    <section class="section">
      <div class="desc-tabs">
        <span>产品描述</span>
        <span>配送、退货和换货</span>
      </div>
      <div class="specs">
        产地：中国<br>
        材质：300g 纯棉（Cotton 100%）<br>
        尺码：S、M、L、XL、2XL<br>
        季节：春季、夏季、秋季<br>
        洗涤说明：可机洗<br>
        重点卖点：厚实不透、领口稳定、日常通勤可直接单穿
      </div>
    </section>

    <section class="section">
      <h2 class="staff-title">员工服装</h2>
      <div class="staff-grid">
        <article class="staff-card">
          <img src="{g}/G15.png" alt="员工服装1">
          <div class="staff-meta">身高180CM\n体重70KG\nパンツサイズXL</div>
        </article>
        <article class="staff-card">
          <img src="{g}/G16.png" alt="员工服装2">
          <div class="staff-meta">身高180CM\n体重70KG\nパンツサイズXL</div>
        </article>
        <article class="staff-card">
          <img src="{g}/G17.png" alt="员工服装3">
          <div class="staff-meta">身高180CM\n体重70KG\nパンツサイズXL</div>
        </article>
      </div>
    </section>

    <div class="footer">Oakln 商品页结构复排预览 · T01 白T · 仅用于素材槽位确认</div>
  </main>
</body>
</html>
"""


def write_preview_html() -> None:
    for preview_dir in WEB_PREVIEW_DIRS:
        preview_dir.mkdir(parents=True, exist_ok=True)
        ensure_preview_assets(preview_dir)

    (ROUND3_DIR / "04_web_preview" / "index.html").write_text(build_html("round3"), encoding="utf-8")
    (BUNDLE_DIR / "04_web_preview" / "index.html").write_text(build_html("bundle"), encoding="utf-8")


def screenshot(html_path: Path) -> None:
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", headless=True)
        page = browser.new_page(viewport={"width": 1600, "height": 2400}, device_scale_factor=1)
        page.goto(html_path.resolve().as_uri(), wait_until="load")
        page.wait_for_timeout(2200)
        page.screenshot(path=str(html_path.with_name("preview_check.png")), full_page=False)
        page.screenshot(path=str(html_path.with_name("preview_full.png")), full_page=True)
        page.screenshot(path=str(html_path.with_name("整页长图.png")), full_page=True)
        browser.close()


def main() -> None:
    write_preview_html()
    screenshot(ROUND3_DIR / "04_web_preview" / "index.html")
    screenshot(BUNDLE_DIR / "04_web_preview" / "index.html")


if __name__ == "__main__":
    main()
