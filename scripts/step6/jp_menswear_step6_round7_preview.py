from __future__ import annotations

import shutil
from pathlib import Path

from PIL import Image, ImageColor, ImageDraw, ImageEnhance, ImageFilter, ImageFont, ImageOps
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
STEP6_DIR = ROOT / "step6_output"
MATERIALS_ROOT = STEP6_DIR / "素材包"
ROUND6_DIR = MATERIALS_ROOT / "T01_round6_2026-03-24"
ROUND7_DIR = MATERIALS_ROOT / "T01_round7_2026-03-24"

FONT_PATHS = [
    Path(r"C:\Windows\Fonts\msyh.ttc"),
    Path(r"C:\Windows\Fonts\msyhbd.ttc"),
]


def load_font(size: int, bold: bool = False):
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


def tint_texture(texture: Image.Image, color_hex: str) -> Image.Image:
    gray = ImageOps.grayscale(texture)
    base = ImageColor.getrgb(color_hex)
    dark = tuple(max(0, int(c * 0.78)) for c in base)
    light = tuple(min(255, int(c * 1.08)) for c in base)
    img = ImageOps.colorize(gray, black=dark, white=light)
    return ImageEnhance.Contrast(img).enhance(1.05)


def rounded_mask(size: tuple[int, int], radius: int) -> Image.Image:
    mask = Image.new("L", size, 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle((0, 0, size[0], size[1]), radius=radius, fill=255)
    return mask


def make_folded_tile(texture: Image.Image, color_hex: str, size: tuple[int, int]) -> Image.Image:
    tile = tint_texture(texture, color_hex)
    tile = cover_resize(tile, size).convert("RGBA")
    mask = rounded_mask(size, radius=20)
    tile.putalpha(mask)

    # Add a subtle fold line and highlight so it reads as folded cloth.
    overlay = Image.new("RGBA", size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    draw.line((32, 54, size[0] - 26, 54), fill=(255, 255, 255, 120), width=3)
    draw.line((32, 60, size[0] - 26, 60), fill=(0, 0, 0, 28), width=2)
    draw.rounded_rectangle((0, 0, size[0], size[1]), radius=20, outline=(255, 255, 255, 28), width=2)
    return Image.alpha_composite(tile, overlay)


def paste_shadowed(canvas: Image.Image, item: Image.Image, pos: tuple[int, int], angle: float = 0.0) -> None:
    if angle:
        item = item.rotate(angle, expand=True, resample=Image.Resampling.BICUBIC)
    shadow = Image.new("RGBA", item.size, (0, 0, 0, 0))
    shadow.putalpha(item.getchannel("A").point(lambda p: int(p * 0.18)))
    shadow = shadow.filter(ImageFilter.GaussianBlur(12))
    canvas.alpha_composite(shadow, (pos[0] + 10, pos[1] + 12))
    canvas.alpha_composite(item, pos)


def build_stack_image(dest: Path, textures: dict[str, Image.Image]) -> None:
    canvas = Image.new("RGBA", (900, 900), (214, 178, 149, 255))
    specs = [
        ("ivory", "#EFE4D2", (86, 324), 0),
        ("white", "#F8F8F6", (156, 268), 0),
        ("gray", "#B8B8B8", (238, 220), 0),
        ("olive", "#7E8565", (332, 182), 0),
        ("black", "#232323", (440, 146), 0),
    ]
    for key, color, pos, angle in specs:
        tile = make_folded_tile(textures[key], color, (338, 220))
        paste_shadowed(canvas, tile, pos, angle=angle)
    canvas.convert("RGB").save(dest, quality=96)


def build_scatter_image(dest: Path, textures: dict[str, Image.Image]) -> None:
    canvas = Image.new("RGBA", (900, 900), (214, 178, 149, 255))
    specs = [
        ("ivory", "#EFE4D2", (82, 330), -10),
        ("white", "#F8F8F6", (176, 278), -6),
        ("gray", "#B8B8B8", (292, 232), -2),
        ("black", "#232323", (420, 198), 7),
    ]
    for key, color, pos, angle in specs:
        tile = make_folded_tile(textures[key], color, (330, 214))
        paste_shadowed(canvas, tile, pos, angle=angle)
    canvas.convert("RGB").save(dest, quality=96)


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


def prepare_round7() -> None:
    web_preview = ROUND7_DIR / "04_web_preview"
    assets_dir = web_preview / "assets"
    generated_dir = ROUND7_DIR / "01_generated"
    generated_dir.mkdir(parents=True, exist_ok=True)
    assets_dir.mkdir(parents=True, exist_ok=True)

    for path in (ROUND6_DIR / "01_generated").glob("*.png"):
        shutil.copy2(path, generated_dir / path.name)
    for path in (ROUND6_DIR / "04_web_preview" / "assets").glob("*"):
        if path.is_file():
            shutil.copy2(path, assets_dir / path.name)

    textures = {
        "white": Image.open(assets_dir / "tail_03_white_flat.jpg").convert("RGB"),
        "gray": Image.open(assets_dir / "tail_04_gray_flat.jpg").convert("RGB"),
        "black": Image.open(assets_dir / "tail_02_black_flat.jpg").convert("RGB"),
        "ivory": Image.open(assets_dir / "tail_05_ivory_flat.jpg").convert("RGB"),
        "olive": Image.open(assets_dir / "tail_06_olive_flat.jpg").convert("RGB"),
    }
    build_stack_image(assets_dir / "top_05_stack_round7.jpg", textures)
    build_scatter_image(assets_dir / "top_07_scatter_round7.jpg", textures)

    build_ref_icon_card(
        source=assets_dir / "top_02_collar.jpg",
        dest=assets_dir / "benefit_neck_round7.jpg",
        headline="领口稳定",
        subline="洗后不易松垮\n日常机洗也更省心",
        icon_label="领",
    )
    build_ref_icon_card(
        source=assets_dir / "top_06_fold.jpg",
        dest=assets_dir / "benefit_soft_round7.jpg",
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
        f"{a}/top_05_stack_round7.jpg",
        f"{a}/top_06_fold.jpg",
        f"{a}/top_07_scatter_round7.jpg",
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
  <title>T01 Round7 商品页预览</title>
  <style>
    * {{ box-sizing: border-box; }}
    body {{ margin: 0; font-family: "Microsoft YaHei", "PingFang SC", sans-serif; color: #1b1b1b; background: #f7f4ee; }}
    img {{ display: block; width: 100%; }}
    .sitebar {{ height: 8px; background: #102c4e; }}
    .header {{ background: #fff; border-bottom: 1px solid #e8e1d6; }}
    .header-inner {{ width: min(1200px, calc(100vw - 52px)); margin: 0 auto; display: flex; align-items: center; justify-content: space-between; padding: 14px 0 12px; }}
    .brand {{ font-family: Georgia, "Times New Roman", serif; font-size: 24px; font-weight: 700; letter-spacing: -0.03em; }}
    .nav, .toolbar {{ display: flex; gap: 18px; color: #6c665d; font-size: 11px; }}
    .wrap {{ width: min(1200px, calc(100vw - 52px)); margin: 0 auto; padding: 16px 0 80px; }}
    .breadcrumb {{ margin-bottom: 14px; color: #91897f; font-size: 11px; }}
    .top {{ display: grid; grid-template-columns: minmax(0, 1.72fr) minmax(340px, 0.88fr); gap: 26px; align-items: start; }}
    .gallery {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; }}
    .gallery img, .tail-gallery img {{ aspect-ratio: 1 / 1; object-fit: cover; background: #fff; }}
    .tail-wrap {{ width: calc((100% - 26px) * 0.661); margin-top: 18px; }}
    .tail-gallery {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; }}
    .buy {{ position: sticky; top: 18px; background: #fbf8f3; border: 1px solid #e6dfd4; padding: 18px 18px 16px; }}
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
    .staff-card {{ position: relative; background: #fff; padding: 4px; box-shadow: 0 0 0 1px #ebe3d7 inset; }}
    .staff-card img {{ aspect-ratio: 0.83 / 1; object-fit: cover; }}
    .staff-meta {{ position: absolute; left: 16px; bottom: 16px; color: rgba(255,255,255,0.88); font-size: 11px; line-height: 1.6; text-shadow: 0 1px 2px rgba(0,0,0,0.18); white-space: pre-line; }}
    .footer {{ margin-top: 56px; padding-top: 18px; border-top: 1px solid #ddd4c8; color: #857d71; font-size: 11px; }}
  </style>
</head>
<body>
  <div class="sitebar"></div>
  <header class="header">
    <div class="header-inner">
      <div class="brand">Oakln</div>
      <nav class="nav"><span>产品列表</span><span>热门商品</span><span>新品上市</span><span>特别专题</span><span>客户支持</span></nav>
      <div class="toolbar"><span>搜索</span><span>账户</span><span>购物车</span></div>
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
          Round7 改成了更干净的折叠衣物块构图。<br>
          这一轮重点是让第 3/4 排先在“镜头类型”上贴近参照页。
        </div>
        <div class="option-title">カラー</div>
        <div class="swatches"><span style="background:#f6f5f1"></span><span style="background:#111111"></span><span style="background:#b5b5b5"></span><span style="background:#6f775e"></span><span style="background:#dfd7c9"></span></div>
        <div class="option-title">サイズ</div>
        <div class="size-grid"><span>S</span><span>M</span><span>L</span><span>XL</span><span>2XL</span></div>
        <div class="option-title">数量</div>
        <div class="qty"><span>-</span><span>1</span><span>+</span></div>
        <div class="option-title">缩略图</div>
        <div class="thumbs"><img src="{g}/G02.png" alt="thumb1"><img src="{g}/G03.png" alt="thumb2"><img src="{a}/top_05_stack_round7.jpg" alt="thumb3"></div>
        <a class="btn primary" href="javascript:void(0)">加入购物车</a>
        <a class="btn secondary" href="javascript:void(0)">立即购买</a>
      </aside>
    </section>
    <section class="tail-wrap"><div class="tail-gallery">
{image_tags(tail_gallery)}
    </div></section>
    <section class="section">
      <div class="benefit-grid">
        <article class="benefit-card"><img src="{g}/G12.png" alt="卖点1"><div class="benefit-title">厚实不透，单穿也有干净轮廓</div><div class="benefit-desc">左侧保持生活方式图，先讲白T单穿时最重要的轮廓和氛围。</div></article>
        <article class="benefit-card"><img src="{a}/benefit_neck_round7.jpg" alt="卖点2"><div class="benefit-title">领口稳定，不易松垮</div><div class="benefit-desc">对应参照页的结构卖点位，继续强调用户最关心的洗后领口稳定性。</div></article>
        <article class="benefit-card"><img src="{a}/benefit_soft_round7.jpg" alt="卖点3"><div class="benefit-title">柔软肌理，机洗也省心</div><div class="benefit-desc">右侧位继续承接好打理和舒适度，让三联区更像原页的信息节奏。</div></article>
      </div>
    </section>
    <section class="section"><div class="desc-tabs"><span>产品描述</span><span>配送、退货和换货</span></div><div class="specs">
      产地：中国<br>
      材质：300g 纯棉（Cotton 100%）<br>
      尺码：S、M、L、XL、2XL<br>
      季节：春季、夏季、秋季<br>
      洗涤说明：可机洗<br>
      重点卖点：厚实不透、领口稳定、通勤可单穿、日常打理省心
    </div></section>
    <section class="section">
      <h2 class="staff-title">员工服装</h2>
      <div class="staff-grid">
        <article class="staff-card"><img src="{g}/G15.png" alt="员工服装1"><div class="staff-meta">身高180CM\n体重70KG\nパンツサイズXL</div></article>
        <article class="staff-card"><img src="{g}/G16.png" alt="员工服装2"><div class="staff-meta">身高180CM\n体重70KG\nパンツサイズXL</div></article>
        <article class="staff-card"><img src="{g}/G17.png" alt="员工服装3"><div class="staff-meta">身高180CM\n体重70KG\nパンツサイズXL</div></article>
      </div>
    </section>
    <div class="footer">Oakln 商品页结构复排预览 · T01 白T · Round7</div>
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
    text = """# T01 Round7

本轮更新重点：

- 顶部第 3/4 排改成了更干净的折叠衣物块构图，避免上一轮的发糊问题。
- 继续保留整页网页预览、检查图和长图。
- 下一轮如果还要继续提高，会优先转到更接近参照页的生活方式图和员工服装区裁切。
"""
    (ROUND7_DIR / "README.md").write_text(text, encoding="utf-8")


def main() -> None:
    prepare_round7()
    web_preview = ROUND7_DIR / "04_web_preview"
    (web_preview / "index.html").write_text(build_html(), encoding="utf-8")
    screenshot(web_preview / "index.html")
    write_readme()


if __name__ == "__main__":
    main()
