from __future__ import annotations

import shutil
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageEnhance, ImageFilter, ImageFont, ImageOps, ImageStat
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
STEP6_DIR = ROOT / "step6_output"
MATERIALS_ROOT = STEP6_DIR / "素材包"
ROUND4_DIR = MATERIALS_ROOT / "T01_round4_2026-03-24"
ROUND5_DIR = MATERIALS_ROOT / "T01_round5_2026-03-24"

FONT_PATHS = [
    Path(r"C:\Windows\Fonts\msyh.ttc"),
    Path(r"C:\Windows\Fonts\msyhbd.ttc"),
]


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
    left = max(0, (resized.width - target_w) // 2)
    top = max(0, (resized.height - target_h) // 2)
    return resized.crop((left, top, left + target_w, top + target_h))


def sample_background(image: Image.Image) -> tuple[int, int, int]:
    patches = [
        image.crop((0, 0, 24, 24)),
        image.crop((image.width - 24, 0, image.width, 24)),
        image.crop((0, image.height - 24, 24, image.height)),
        image.crop((image.width - 24, image.height - 24, image.width, image.height)),
    ]
    merged = Image.new("RGB", (48, 48), (255, 255, 255))
    for idx, patch in enumerate(patches):
        merged.paste(patch.convert("RGB"), ((idx % 2) * 24, (idx // 2) * 24))
    stat = ImageStat.Stat(merged)
    return tuple(int(v) for v in stat.mean[:3])


def extract_subject(path: Path, threshold: int = 18) -> Image.Image:
    image = Image.open(path).convert("RGB")
    bg_color = sample_background(image)
    bg = Image.new("RGB", image.size, bg_color)
    diff = ImageChops.difference(image, bg)
    gray = ImageOps.grayscale(diff)
    mask = gray.point(lambda p: 0 if p < threshold else min(255, (p - threshold) * 10))
    mask = mask.filter(ImageFilter.GaussianBlur(1.8))
    rgba = image.convert("RGBA")
    rgba.putalpha(mask)
    return rgba


def tint_subject(image: Image.Image, dark: tuple[int, int, int], light: tuple[int, int, int]) -> Image.Image:
    alpha = image.getchannel("A")
    gray = ImageOps.grayscale(image.convert("RGB"))
    rgb = ImageOps.colorize(gray, black=dark, white=light).convert("RGBA")
    rgb.putalpha(alpha)
    return rgb


def paste_with_shadow(canvas: Image.Image, image: Image.Image, x: int, y: int, blur: int = 12, alpha: float = 0.22) -> None:
    shadow = Image.new("RGBA", image.size, (0, 0, 0, 0))
    shadow.putalpha(image.getchannel("A").point(lambda p: int(p * alpha)))
    shadow = shadow.filter(ImageFilter.GaussianBlur(blur))
    canvas.alpha_composite(shadow, (x + 12, y + 15))
    canvas.alpha_composite(image, (x, y))


def compose_fold_stack(dest: Path, assets_dir: Path) -> None:
    fold = extract_subject(assets_dir / "top_06_fold.jpg")
    variants = [
        tint_subject(fold, (198, 188, 174), (246, 240, 232)),
        tint_subject(fold, (232, 232, 232), (252, 252, 252)),
        tint_subject(fold, (156, 156, 156), (210, 210, 210)),
        tint_subject(fold, (92, 100, 78), (148, 156, 122)),
        tint_subject(fold, (18, 18, 18), (70, 70, 70)),
    ]

    canvas = Image.new("RGBA", (900, 900), (214, 178, 149, 255))
    positions = [(74, 250), (155, 220), (248, 193), (343, 166), (445, 136)]
    scales = [0.88, 0.91, 0.93, 0.95, 0.98]
    for variant, pos, scale in zip(variants, positions, scales):
        item = variant.resize((int(variant.width * scale), int(variant.height * scale)), Image.Resampling.LANCZOS)
        paste_with_shadow(canvas, item, pos[0], pos[1], blur=12, alpha=0.20)
    canvas.convert("RGB").save(dest, quality=95)


def compose_fold_scatter(dest: Path, assets_dir: Path) -> None:
    fold = extract_subject(assets_dir / "top_06_fold.jpg")
    variants = [
        tint_subject(fold, (198, 188, 174), (246, 240, 232)),
        tint_subject(fold, (232, 232, 232), (252, 252, 252)),
        tint_subject(fold, (156, 156, 156), (210, 210, 210)),
        tint_subject(fold, (18, 18, 18), (70, 70, 70)),
    ]

    canvas = Image.new("RGBA", (900, 900), (214, 178, 149, 255))
    rotations = [-11, -7, -3, 8]
    positions = [(84, 228), (178, 178), (286, 144), (402, 126)]
    scales = [0.82, 0.86, 0.90, 0.93]
    for variant, angle, pos, scale in zip(variants, rotations, positions, scales):
        item = variant.resize((int(variant.width * scale), int(variant.height * scale)), Image.Resampling.LANCZOS)
        item = item.rotate(angle, expand=True, resample=Image.Resampling.BICUBIC)
        paste_with_shadow(canvas, item, pos[0], pos[1], blur=15, alpha=0.20)
    canvas.convert("RGB").save(dest, quality=95)


def add_bottom_gradient(image: Image.Image, strength: int = 145) -> Image.Image:
    base = image.convert("RGBA")
    overlay = Image.new("RGBA", base.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    width, height = base.size
    for idx in range(height):
        alpha = int(max(0, strength * ((idx / height) ** 2.15)))
        draw.line((0, idx, width, idx), fill=(0, 0, 0, alpha))
    return Image.alpha_composite(base, overlay)


def build_ref_icon_card(source: Path, dest: Path, headline: str, subline: str, icon_label: str) -> None:
    base = Image.open(source).convert("RGB")
    canvas = cover_resize(base, (900, 900))
    canvas = ImageEnhance.Brightness(canvas).enhance(0.96)
    canvas = add_bottom_gradient(canvas, strength=175).convert("RGBA")
    draw = ImageDraw.Draw(canvas)
    headline_font = load_font(42, bold=True)
    sub_font = load_font(24)
    icon_font = load_font(26, bold=True)
    center_x = canvas.width // 2
    circle_box = (center_x - 56, 530, center_x + 56, 642)
    draw.ellipse(circle_box, outline=(255, 255, 255, 205), width=4)
    label_box = draw.textbbox((0, 0), icon_label, font=icon_font)
    label_w = label_box[2] - label_box[0]
    label_h = label_box[3] - label_box[1]
    draw.text((center_x - label_w / 2, 586 - label_h / 2), icon_label, font=icon_font, fill=(255, 255, 255, 235))
    headline_box = draw.textbbox((0, 0), headline, font=headline_font)
    headline_w = headline_box[2] - headline_box[0]
    draw.text((center_x - headline_w / 2, 680), headline, font=headline_font, fill=(255, 255, 255, 245))
    sub_box = draw.multiline_textbbox((0, 0), subline, font=sub_font, spacing=6, align="center")
    sub_w = sub_box[2] - sub_box[0]
    draw.multiline_text((center_x - sub_w / 2, 746), subline, font=sub_font, fill=(255, 255, 255, 220), spacing=6, align="center")
    canvas.convert("RGB").save(dest, quality=94)


def prepare_round5() -> None:
    web_preview = ROUND5_DIR / "04_web_preview"
    assets_dir = web_preview / "assets"
    generated_dir = ROUND5_DIR / "01_generated"
    generated_dir.mkdir(parents=True, exist_ok=True)
    assets_dir.mkdir(parents=True, exist_ok=True)

    for path in (ROUND4_DIR / "01_generated").glob("*.png"):
        shutil.copy2(path, generated_dir / path.name)
    for path in (ROUND4_DIR / "04_web_preview" / "assets").glob("*"):
        if path.is_file():
            shutil.copy2(path, assets_dir / path.name)

    compose_fold_stack(assets_dir / "top_05_stack_round5.jpg", assets_dir)
    compose_fold_scatter(assets_dir / "top_07_scatter_round5.jpg", assets_dir)

    build_ref_icon_card(
        source=assets_dir / "top_02_collar.jpg",
        dest=assets_dir / "benefit_neck_round5.jpg",
        headline="领口稳定",
        subline="洗后不易松垮\n日常机洗也更省心",
        icon_label="领",
    )
    build_ref_icon_card(
        source=assets_dir / "top_06_fold.jpg",
        dest=assets_dir / "benefit_soft_round5.jpg",
        headline="柔软肌理",
        subline="贴肤不扎\n通勤和内搭都顺手",
        icon_label="柔",
    )


def build_html() -> str:
    g = "../01_generated"
    a = "./assets"
    top_gallery = [
        f"{g}/G02.png",
        f"{a}/top_02_collar.jpg",
        f"{g}/G03.png",
        f"{a}/top_04_hanging.jpg",
        f"{a}/top_05_stack_round5.jpg",
        f"{a}/top_06_fold.jpg",
        f"{a}/top_07_scatter_round5.jpg",
        f"{a}/top_02_collar.jpg",
    ]
    tail_gallery = [
        f"{a}/tail_01_cuff.jpg",
        f"{a}/tail_02_black_flat.jpg",
        f"{a}/tail_03_white_flat.jpg",
        f"{a}/tail_04_gray_flat.jpg",
        f"{a}/tail_05_ivory_flat.jpg",
        f"{a}/tail_06_olive_flat.jpg",
        f"{a}/tail_07_black_flat.jpg",
        f"{a}/tail_08_gray_flat.jpg",
    ]

    def image_tags(items: list[str]) -> str:
        return "\n".join(f'        <img src="{src}" alt="preview">' for src in items)

    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>T01 Round5 商品页预览</title>
  <style>
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Microsoft YaHei", "PingFang SC", sans-serif;
      color: #1b1b1b;
      background: #f7f4ee;
    }}
    img {{ display: block; width: 100%; }}
    .sitebar {{ height: 8px; background: #102c4e; }}
    .header {{ background: #fff; border-bottom: 1px solid #e8e1d6; }}
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
    .breadcrumb {{ margin-bottom: 14px; color: #91897f; font-size: 11px; }}
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
    .gallery img, .tail-gallery img {{
      aspect-ratio: 1 / 1;
      object-fit: cover;
      background: #fff;
    }}
    .tail-wrap {{ width: calc((100% - 26px) * 0.661); margin-top: 18px; }}
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
    h1 {{ margin: 0 0 12px; font-size: 27px; line-height: 1.45; font-weight: 700; }}
    .price {{ font-size: 42px; font-weight: 700; margin-bottom: 8px; }}
    .price span {{ margin-left: 10px; color: #8e8579; text-decoration: line-through; font-size: 16px; font-weight: 400; }}
    .meta {{ color: #8e8579; font-size: 11px; line-height: 1.8; margin-bottom: 14px; }}
    .option-title {{ margin: 14px 0 8px; font-size: 12px; font-weight: 700; }}
    .swatches {{ display: flex; gap: 7px; margin-bottom: 12px; }}
    .swatches span {{ width: 14px; height: 14px; display: inline-block; border-radius: 50%; border: 1px solid #cfc6b9; }}
    .size-grid {{ display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 12px; }}
    .size-grid span {{ min-width: 30px; padding: 6px 8px; font-size: 11px; text-align: center; border: 1px solid #d8cfbf; background: #fff; }}
    .thumbs {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; margin: 12px 0 16px; }}
    .thumbs img {{ aspect-ratio: 1 / 1; object-fit: cover; border: 1px solid #e6dfd4; background: #fff; }}
    .qty {{ display: flex; align-items: center; gap: 8px; margin-bottom: 12px; }}
    .qty span {{ width: 28px; height: 28px; display: inline-flex; align-items: center; justify-content: center; border: 1px solid #d8cfbf; background: #fff; font-size: 12px; }}
    .btn {{ display: block; width: 100%; padding: 11px 14px; margin-top: 10px; text-decoration: none; text-align: center; font-size: 13px; font-weight: 700; letter-spacing: 0.03em; }}
    .primary {{ background: #17375d; color: #fff; }}
    .secondary {{ border: 1px solid #17375d; color: #17375d; background: #fff; }}
    .section {{ margin-top: 34px; }}
    .benefit-grid {{ display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 16px; }}
    .benefit-card img {{ aspect-ratio: 1 / 1; object-fit: cover; background: #fff; }}
    .benefit-title {{ margin: 10px 0 6px; font-size: 14px; font-weight: 700; line-height: 1.45; }}
    .benefit-desc {{ color: #514a42; font-size: 12px; line-height: 1.75; }}
    .desc-tabs {{ display: flex; gap: 20px; padding-bottom: 10px; margin-bottom: 12px; border-bottom: 1px solid #ddd4c8; font-size: 13px; }}
    .desc-tabs span:first-child {{ font-weight: 700; }}
    .specs {{ color: #514a42; font-size: 12px; line-height: 1.95; }}
    .staff-title {{ margin: 0 0 14px; font-size: 18px; font-weight: 700; }}
    .staff-grid {{ display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 16px; }}
    .staff-card {{
      position: relative;
      background: #fff;
      padding: 4px;
      box-shadow: 0 0 0 1px #ebe3d7 inset;
    }}
    .staff-card img {{ aspect-ratio: 0.83 / 1; object-fit: cover; }}
    .staff-meta {{
      position: absolute;
      left: 16px;
      bottom: 16px;
      color: rgba(255,255,255,0.88);
      font-size: 11px;
      line-height: 1.6;
      text-shadow: 0 1px 2px rgba(0,0,0,0.18);
      white-space: pre-line;
    }}
    .footer {{ margin-top: 56px; padding-top: 18px; border-top: 1px solid #ddd4c8; color: #857d71; font-size: 11px; }}
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
{image_tags(top_gallery)}
      </div>
      <aside class="buy">
        <h1>300g 重磅コットン クルーネックTシャツ<br>メンズ 半袖 厚手 透けにくい ベーシック</h1>
        <div class="price">¥6,980<span>¥7,980</span></div>
        <div class="meta">
          Round5 继续收顶部镜头。<br>
          这轮把第 3/4 排改成“折叠衣物镜头”而不是“平铺 T 恤叠放”。
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
          <img src="{a}/top_05_stack_round5.jpg" alt="thumb3">
        </div>
        <a class="btn primary" href="javascript:void(0)">加入购物车</a>
        <a class="btn secondary" href="javascript:void(0)">立即购买</a>
      </aside>
    </section>
    <section class="tail-wrap">
      <div class="tail-gallery">
{image_tags(tail_gallery)}
      </div>
    </section>
    <section class="section">
      <div class="benefit-grid">
        <article class="benefit-card">
          <img src="{g}/G12.png" alt="卖点1">
          <div class="benefit-title">厚实不透，单穿也有干净轮廓</div>
          <div class="benefit-desc">左侧保持生活方式图，先讲白T单穿时最重要的轮廓和氛围。</div>
        </article>
        <article class="benefit-card">
          <img src="{a}/benefit_neck_round5.jpg" alt="卖点2">
          <div class="benefit-title">领口稳定，不易松垮</div>
          <div class="benefit-desc">对应参照页的结构卖点位，继续强调用户最关心的洗后领口稳定性。</div>
        </article>
        <article class="benefit-card">
          <img src="{a}/benefit_soft_round5.jpg" alt="卖点3">
          <div class="benefit-title">柔软肌理，机洗也省心</div>
          <div class="benefit-desc">右侧位继续承接好打理和舒适度，让三联区更像原页的信息节奏。</div>
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
        重点卖点：厚实不透、领口稳定、通勤可单穿、日常打理省心
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
    <div class="footer">Oakln 商品页结构复排预览 · T01 白T · Round5</div>
  </main>
</body>
</html>
"""


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


def write_readme() -> None:
    text = """# T01 Round5

本轮更新重点：

- 顶部第 3/4 排改成了折叠衣物镜头，不再是平铺 T 恤叠放。
- 员工服装卡片加了更接近参照页的白底边框感。
- 继续保留整页网页预览、检查图和长图。
"""
    (ROUND5_DIR / "README.md").write_text(text, encoding="utf-8")


def main() -> None:
    prepare_round5()
    web_preview = ROUND5_DIR / "04_web_preview"
    (web_preview / "index.html").write_text(build_html(), encoding="utf-8")
    screenshot(web_preview / "index.html")
    write_readme()


if __name__ == "__main__":
    main()
