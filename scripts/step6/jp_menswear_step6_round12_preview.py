from __future__ import annotations

import shutil
from pathlib import Path

from PIL import Image
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
STEP6 = ROOT / "step6_output"
ROUND7 = next((STEP6).glob("*/T01_round7_2026-03-24"))
ROUND10 = next((STEP6).glob("*/T01_round10_2026-03-24"))
ROUND12 = next((STEP6).glob("*/T01_round12_2026-03-25"))


def ensure_dirs() -> tuple[Path, Path]:
    preview = ROUND12 / "04_web_preview"
    assets = preview / "assets"
    preview.mkdir(parents=True, exist_ok=True)
    assets.mkdir(parents=True, exist_ok=True)
    return preview, assets


def copy_base_assets(assets_dir: Path) -> None:
    for path in (ROUND7 / "04_web_preview" / "assets").glob("*"):
        if path.is_file():
            shutil.copy2(path, assets_dir / path.name)


def save_as_jpg(src: Path, dest: Path) -> None:
    img = Image.open(src).convert("RGB")
    img.save(dest, quality=96)


def prepare_round12_assets(assets_dir: Path) -> None:
    harvest = ROUND12 / "04_harvested"
    stack = harvest / "lovart_stack_multicolor_exact.jpg"
    fold = harvest / "lovart_fold_multicolor_exact.jpg"
    if stack.exists():
        shutil.copy2(stack, assets_dir / "top_05_stack_round12.jpg")
    if fold.exists():
        shutil.copy2(fold, assets_dir / "top_07_fold_multicolor_round12.jpg")


def build_html() -> str:
    g = "../../T01_round10_2026-03-24/01_generated"
    a = "./assets"
    top_gallery = [
        f"{g}/G02.png",
        f"{a}/top_02_collar.jpg",
        f"{g}/G03.png",
        f"{a}/top_04_hanging.jpg",
        f"{a}/top_05_stack_round12.jpg",
        f"{a}/top_06_fold.jpg",
        f"{a}/top_07_fold_multicolor_round12.jpg",
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
  <title>T01 Round12 商品页预览</title>
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
    <div class="breadcrumb">首页 / 男装 / T 恤 / 300g 重磅白 T</div>
    <section class="top">
      <div class="gallery">
{image_tags(top_gallery)}
      </div>
      <aside class="buy">
        <h1>300g 重磅棉质 圆领基础白T<br>男士短袖 厚实 不易透视 基础款</h1>
        <div class="price">¥6,980<span>¥7,980</span></div>
        <div class="meta">
          Round12 当前重点是把整页镜头逻辑往参照页继续收紧。<br>
          本版已替换成新的多色堆叠图和多色层叠折叠近景。
        </div>
        <div class="option-title">颜色</div>
        <div class="swatches"><span style="background:#f6f5f1"></span><span style="background:#111111"></span><span style="background:#b5b5b5"></span><span style="background:#6f775e"></span><span style="background:#dfd7c9"></span></div>
        <div class="option-title">尺码</div>
        <div class="size-grid"><span>S</span><span>M</span><span>L</span><span>XL</span><span>2XL</span></div>
        <div class="option-title">数量</div>
        <div class="qty"><span>-</span><span>1</span><span>+</span></div>
        <div class="option-title">缩略图</div>
        <div class="thumbs"><img src="{g}/G02.png" alt="thumb1"><img src="{g}/G03.png" alt="thumb2"><img src="{a}/top_05_stack_round12.jpg" alt="thumb3"></div>
        <a class="btn primary" href="javascript:void(0)">加入购物车</a>
        <a class="btn secondary" href="javascript:void(0)">立即购买</a>
      </aside>
    </section>
    <section class="tail-wrap"><div class="tail-gallery">
{image_tags(tail_gallery)}
    </div></section>
    <section class="section">
      <div class="benefit-grid">
        <article class="benefit-card"><img src="{g}/G12.png" alt="卖点1"><div class="benefit-title">厚实不透，单穿也有干净轮廓</div><div class="benefit-desc">左侧继续保留生活方式图，强调白T 单穿时最重要的轮廓和氛围。</div></article>
        <article class="benefit-card"><img src="{a}/benefit_neck_round7.jpg" alt="卖点2"><div class="benefit-title">领口稳定，不易松垮</div><div class="benefit-desc">对应参照页的结构卖点位，持续强化用户最关心的洗后领口稳定性。</div></article>
        <article class="benefit-card"><img src="{a}/benefit_soft_round7.jpg" alt="卖点3"><div class="benefit-title">柔软肌理，机洗也省心</div><div class="benefit-desc">右侧承接第二卖点，保持好打理和舒适度。</div></article>
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
        <article class="staff-card"><img src="{g}/G15.png" alt="员工服装1"><div class="staff-meta">身长180CM\n体重70KG\n穿着尺码XL</div></article>
        <article class="staff-card"><img src="{g}/G16.png" alt="员工服装2"><div class="staff-meta">身长180CM\n体重70KG\n穿着尺码XL</div></article>
        <article class="staff-card"><img src="{g}/G17.png" alt="员工服装3"><div class="staff-meta">身长180CM\n体重70KG\n穿着尺码XL</div></article>
      </div>
    </section>
    <div class="footer">Oakln 商品页结构预览 · T01 白T · Round12</div>
  </main>
</body>
</html>
"""


def screenshot(html_path: Path) -> None:
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", headless=True)
        page = browser.new_page(viewport={"width": 1600, "height": 2400}, device_scale_factor=1)
        page.goto(html_path.resolve().as_uri(), wait_until="load")
        page.wait_for_timeout(1800)
        page.screenshot(path=str(html_path.with_name("preview_check.png")), full_page=False)
        page.screenshot(path=str(html_path.with_name("preview_full.png")), full_page=True)
        page.screenshot(path=str(html_path.with_name("整页长图.png")), full_page=True)
        browser.close()


def main() -> None:
    preview, assets = ensure_dirs()
    copy_base_assets(assets)
    prepare_round12_assets(assets)
    html = preview / "index.html"
    html.write_text(build_html(), encoding="utf-8")
    screenshot(html)
    print(html)


if __name__ == "__main__":
    main()
