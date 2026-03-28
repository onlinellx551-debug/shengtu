from __future__ import annotations

import shutil
from pathlib import Path

from PIL import Image
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
STEP6 = ROOT / "step6_output" / "material_bundle"
ROUND20 = STEP6 / "T01_round20_2026-03-26"
PREVIEW = ROUND20 / "04_web_preview"
ASSETS = PREVIEW / "assets"
ORIG = ROOT / "step6_output" / "T01_案例版详情素材_2026-03-20" / "01_同款精选原图"


def reset_dir(path: Path) -> None:
    if path.exists():
        shutil.rmtree(path)
    path.mkdir(parents=True, exist_ok=True)


def copy(src: Path, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dest)


def scale_product(src: Path, dest: Path, scale: float = 1.12) -> None:
    im = Image.open(src).convert("RGB")
    canvas = Image.new("RGB", im.size, "white")
    nw = int(im.width * scale)
    nh = int(im.height * scale)
    resized = im.resize((nw, nh), Image.LANCZOS)
    x = (canvas.width - nw) // 2
    y = (canvas.height - nh) // 2 + 6
    canvas.paste(resized, (x, y))
    canvas.save(dest, quality=96)


def stage_assets() -> None:
    reset_dir(PREVIEW)
    ASSETS.mkdir(parents=True, exist_ok=True)
    src_assets = ROUND20 / "03_assets"
    for p in src_assets.glob("*"):
        copy(p, ASSETS / p.name)

    # top01: tighten product ratio so the flat tee fills the square more like the reference.
    scale_product(ORIG / "白色平铺.jpg", ASSETS / "top01_real_front_whitespace.jpg", scale=1.18)

    copy(ROUND20 / "02_generated" / "top02_onbody_round20_clean.jpg", ASSETS / "top02_onbody_round20_clean.jpg")
    copy(ROUND20 / "02_generated" / "top03_onbody_round20_clean.jpg", ASSETS / "top03_onbody_round20_clean.jpg")
    copy(ROUND20 / "02_generated" / "top04_onbody_round20_clean.jpg", ASSETS / "top04_onbody_round20_clean.jpg")

    copy(ROUND20 / "02_generated" / "benefit01_lifestyle_round18.png", ASSETS / "benefit01_lifestyle_round18.png")
    copy(ROUND20 / "02_generated" / "staff_01_round18.png", ASSETS / "staff_01_round18.png")
    copy(ROUND20 / "02_generated" / "staff_02_round18.png", ASSETS / "staff_02_round18.png")
    copy(ROUND20 / "02_generated" / "staff_03_round18.png", ASSETS / "staff_03_round18.png")


def build_html() -> str:
    gallery = [
        "./assets/top01_real_front_whitespace.jpg",
        "./assets/top02_onbody_round20_clean.jpg",
        "./assets/top03_onbody_round20_clean.jpg",
        "./assets/top04_onbody_round20_clean.jpg",
        "./assets/top_05_stack_round14.jpg",
        "./assets/top_06_fold.jpg",
        "./assets/top_07_fold_round14.jpg",
        "./assets/top_08_collar.jpg",
    ]
    tails = [
        "./assets/tail_01_hand_detail.jpg",
        "./assets/tail_02_black_flat.jpg",
        "./assets/tail_03_white_flat.jpg",
        "./assets/tail_04_gray_flat.jpg",
        "./assets/tail_05_ivory_flat.jpg",
        "./assets/tail_06_olive_flat.jpg",
        "./assets/tail_07_black_flat.jpg",
        "./assets/tail_08_gray_flat.jpg",
    ]

    def imgs(items: list[str]) -> str:
        return "\n".join(f'        <img src="{item}" alt="preview">' for item in items)

    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>T01 商品页预览</title>
  <style>
    * {{ box-sizing: border-box; }}
    body {{ margin: 0; background: #f7f4ee; color: #1b1b1b; font-family: "Microsoft YaHei", "PingFang SC", sans-serif; }}
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
    <div class="breadcrumb">首页 / 男装 / T 恤 / 300g 重磅基础白T</div>
    <section class="top">
      <div class="gallery">
{imgs(gallery)}
      </div>
      <aside class="buy">
        <h1>300g 重磅棉质 圆领基础白T<br>男士短袖 厚实 不易透视 基础款</h1>
        <div class="price">¥6,980<span>¥7,980</span></div>
        <div class="meta">
          当前版本已按参考页的主槽位顺序重排。<br>
          top_01 使用真实单品白底图，top_02 到 top_04 正在继续往“同构图 + 不同颜色”推进，下面 4 张则继续往真实短袖 T 质感收敛。
        </div>
        <div class="option-title">颜色</div>
        <div class="swatches"><span style="background:#f6f5f1"></span><span style="background:#111111"></span><span style="background:#b5b5b5"></span><span style="background:#6f775e"></span><span style="background:#dfd7c9"></span></div>
        <div class="option-title">尺码</div>
        <div class="size-grid"><span>S</span><span>M</span><span>L</span><span>XL</span><span>2XL</span></div>
        <div class="option-title">数量</div>
        <div class="qty"><span>-</span><span>1</span><span>+</span></div>
        <div class="option-title">缩略图</div>
        <div class="thumbs"><img src="./assets/top01_real_front_whitespace.jpg" alt="thumb1"><img src="./assets/top02_onbody_round20_clean.jpg" alt="thumb2"><img src="./assets/top_05_stack_round14.jpg" alt="thumb3"></div>
        <a class="btn primary" href="javascript:void(0)">加入购物车</a>
        <a class="btn secondary" href="javascript:void(0)">立即购买</a>
      </aside>
    </section>
    <section class="tail-wrap"><div class="tail-gallery">
{imgs(tails)}
    </div></section>
    <section class="section">
      <div class="benefit-grid">
        <article class="benefit-card"><img src="./assets/benefit01_lifestyle_round18.png" alt="卖点1"><div class="benefit-title">厚实不透，单穿也有干净轮廓</div><div class="benefit-desc">先用真实生活方式图回答白 T 最核心的问题：单穿是否有分量感，是否足够干净利落。</div></article>
        <article class="benefit-card"><img src="./assets/benefit_02_clean.jpg" alt="卖点2"><div class="benefit-title">领口稳定，反复穿洗也更省心</div><div class="benefit-desc">用结构近景回应用户最在意的领口问题，强调圆领稳定、走线干净和日常打理友好。</div></article>
        <article class="benefit-card"><img src="./assets/benefit_03_clean.jpg" alt="卖点3"><div class="benefit-title">300g 厚实棉感，单穿更安心</div><div class="benefit-desc">用折叠近景说明厚度、触感和层次，把“厚实不透、可机洗”直接落在画面里。</div></article>
      </div>
    </section>
    <section class="section">
      <div class="desc-tabs"><span>产品描述</span><span>配送、退货和换货</span></div>
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
        <article class="staff-card"><img src="./assets/staff_01_round18.png" alt="员工服装1"><div class="staff-meta">身高180CM\n体重70KG\n穿着尺码XL</div></article>
        <article class="staff-card"><img src="./assets/staff_02_round18.png" alt="员工服装2"><div class="staff-meta">身高180CM\n体重70KG\n穿着尺码XL</div></article>
        <article class="staff-card"><img src="./assets/staff_03_round18.png" alt="员工服装3"><div class="staff-meta">身高180CM\n体重70KG\n穿着尺码XL</div></article>
      </div>
    </section>
    <div class="footer">Oakln 商品页结构预览 · T01 白T · Round20</div>
  </main>
</body>
</html>
"""


def screenshot(html: Path) -> None:
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", headless=True)
        page = browser.new_page(viewport={"width": 1600, "height": 2400}, device_scale_factor=1)
        page.goto(html.resolve().as_uri(), wait_until="load")
        page.wait_for_timeout(1800)
        page.evaluate("window.scrollTo(0, 0)")
        page.wait_for_timeout(500)
        page.screenshot(path=str(PREVIEW / "preview_check.png"), full_page=False)
        page.screenshot(path=str(PREVIEW / "preview_full.png"), full_page=True)
        page.screenshot(path=str(PREVIEW / "整页长图.png"), full_page=True)
        browser.close()


def main() -> None:
    stage_assets()
    html = PREVIEW / "index.html"
    html.write_text(build_html(), encoding="utf-8")
    screenshot(html)
    print(html)


if __name__ == "__main__":
    main()
