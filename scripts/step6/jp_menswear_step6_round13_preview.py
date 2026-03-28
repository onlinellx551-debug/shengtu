from __future__ import annotations

import shutil
from pathlib import Path

from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
STEP6 = ROOT / "step6_output" / "material_bundle"
ROUND12 = STEP6 / "T01_round12_2026-03-25"
ROUND13 = STEP6 / "T01_round13_2026-03-25"
PREVIEW = ROUND13 / "04_web_preview"
ASSETS = PREVIEW / "assets"


def ensure_dirs() -> None:
    ASSETS.mkdir(parents=True, exist_ok=True)


def copy_base_assets() -> None:
    src = ROUND12 / "04_web_preview" / "assets"
    for path in src.glob("*"):
        if path.is_file():
            shutil.copy2(path, ASSETS / path.name)


def build_html() -> str:
    g = "../../T01_round10_2026-03-24/01_generated"
    a = "./assets"
    top_gallery = [
        f"{g}/G02.png",
        f"{a}/top_02_collar.jpg",
        f"{g}/G03.png",
        f"{a}/top_04_hanging.jpg",
        f"{a}/top_05_stack_round13.jpg",
        f"{a}/top_06_fold.jpg",
        f"{a}/top_07_fold_round13.jpg",
        f"{a}/top_08_collar.jpg",
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

    def imgs(items: list[str]) -> str:
        return "\n".join(f'        <img src="{item}" alt="preview">' for item in items)

    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>T01 Round13 鍟嗗搧椤甸瑙?/title>
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
      <nav class="nav"><span>浜у搧鍒楄〃</span><span>鐑棬鍟嗗搧</span><span>鏂板搧涓婂競</span><span>鐗瑰埆涓撻</span><span>瀹㈡埛鏀寔</span></nav>
      <div class="toolbar"><span>鎼滅储</span><span>璐︽埛</span><span>璐墿杞?/span></div>
    </div>
  </header>
  <main class="wrap">
    <div class="breadcrumb">棣栭〉 / 鐢疯 / T鎭?/ 300g 閲嶇鍦嗛鐧絋</div>
    <section class="top">
      <div class="gallery">
{imgs(top_gallery)}
      </div>
      <aside class="buy">
        <h1>300g 閲嶇妫夎川 鍦嗛鍩虹鐧絋<br>鐢峰＋鐭 鍘氬疄 涓嶆槗閫忚 鍩虹娆?/h1>
        <div class="price">楼6,980<span>楼7,980</span></div>
        <div class="meta">
          Round13 褰撳墠閲嶇偣鏄妸鍙傜収椤甸《閮ㄩ潤鐗╂Ы浣嶇户缁敹鍑嗐€?br>
          杩欑増椤堕儴宸茬粡鏇挎崲鎴愭柊鐨勫鑹插爢鍙犲浘鍜屽鑹叉姌鍙犺繎鏅浘銆?        </div>
        <div class="option-title">棰滆壊</div>
        <div class="swatches"><span style="background:#f6f5f1"></span><span style="background:#111111"></span><span style="background:#b5b5b5"></span><span style="background:#6f775e"></span><span style="background:#dfd7c9"></span></div>
        <div class="option-title">灏虹爜</div>
        <div class="size-grid"><span>S</span><span>M</span><span>L</span><span>XL</span><span>2XL</span></div>
        <div class="option-title">鏁伴噺</div>
        <div class="qty"><span>-</span><span>1</span><span>+</span></div>
        <div class="option-title">缂╃暐鍥?/div>
        <div class="thumbs"><img src="{g}/G02.png" alt="thumb1"><img src="{g}/G03.png" alt="thumb2"><img src="{a}/top_05_stack_round13.jpg" alt="thumb3"></div>
        <a class="btn primary" href="javascript:void(0)">鍔犲叆璐墿杞?/a>
        <a class="btn secondary" href="javascript:void(0)">绔嬪嵆璐拱</a>
      </aside>
    </section>
    <section class="tail-wrap"><div class="tail-gallery">
{imgs(tail_gallery)}
    </div></section>
    <section class="section">
      <div class="benefit-grid">
        <article class="benefit-card"><img src="{g}/G12.png" alt="鍗栫偣1"><div class="benefit-title">鍘氬疄涓嶉€忥紝鍗曠┛涔熸湁骞插噣杞粨</div><div class="benefit-desc">宸︿晶缁х画淇濈暀鐢熸椿鏂瑰紡鍥撅紝寮鸿皟鐧絋鍗曠┛鏃舵渶閲嶈鐨勮疆寤撳拰姘涘洿銆?/div></article>
        <article class="benefit-card"><img src="{a}/benefit_neck_round7.jpg" alt="鍗栫偣2"><div class="benefit-title">棰嗗彛绋冲畾锛屼笉鏄撴澗鍨?/div><div class="benefit-desc">瀵瑰簲鍙傜収椤电殑缁撴瀯鍗栫偣浣嶏紝鎸佺画寮哄寲鐢ㄦ埛鏈€鍏冲績鐨勬礂鍚庨鍙ｇǔ瀹氭€с€?/div></article>
        <article class="benefit-card"><img src="{a}/benefit_soft_round7.jpg" alt="鍗栫偣3"><div class="benefit-title">鏌旇蒋鑲岀悊锛屾満娲椾篃鐪佸績</div><div class="benefit-desc">鍙充晶鎵挎帴绗簩鍗栫偣锛屼繚鎸佸ソ鎵撶悊鍜岃垝閫傚害銆?/div></article>
      </div>
    </section>
    <section class="section"><div class="desc-tabs"><span>浜у搧鎻忚堪</span><span>閰嶉€併€侀€€璐у拰鎹㈣揣</span></div><div class="specs">
      浜у湴锛氫腑鍥?br>
      鏉愯川锛?00g 绾锛圕otton 100%锛?br>
      灏虹爜锛歋銆丮銆丩銆乆L銆?XL<br>
      瀛ｈ妭锛氭槬瀛ｃ€佸瀛ｃ€佺瀛?br>
      娲楁钉璇存槑锛氬彲鏈烘礂<br>
      閲嶇偣鍗栫偣锛氬帤瀹炰笉閫忋€侀鍙ｇǔ瀹氥€侀€氬嫟鍙崟绌裤€佹棩甯告墦鐞嗙渷蹇?    </div></section>
    <section class="section">
      <h2 class="staff-title">鍛樺伐鏈嶈</h2>
      <div class="staff-grid">
        <article class="staff-card"><img src="{g}/G15.png" alt="鍛樺伐鏈嶈1"><div class="staff-meta">韬珮180CM\n浣撻噸70KG\n绌跨潃灏虹爜XL</div></article>
        <article class="staff-card"><img src="{g}/G16.png" alt="鍛樺伐鏈嶈2"><div class="staff-meta">韬珮180CM\n浣撻噸70KG\n绌跨潃灏虹爜XL</div></article>
        <article class="staff-card"><img src="{g}/G17.png" alt="鍛樺伐鏈嶈3"><div class="staff-meta">韬珮180CM\n浣撻噸70KG\n绌跨潃灏虹爜XL</div></article>
      </div>
    </section>
    <div class="footer">Oakln 鍟嗗搧椤电粨鏋勯瑙?路 T01 鐧絋 路 Round13</div>
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
        page.screenshot(path=str(PREVIEW / "preview_check.png"), full_page=False)
        page.screenshot(path=str(PREVIEW / "preview_full.png"), full_page=True)
        page.screenshot(path=str(PREVIEW / "鏁撮〉闀垮浘.png"), full_page=True)
        browser.close()


def main() -> None:
    ensure_dirs()
    copy_base_assets()
    shutil.copy2(ROUND13 / "04_harvested" / "lovart_stack_retry_simple.jpg", ASSETS / "top_05_stack_round13.jpg")
    shutil.copy2(ROUND13 / "04_harvested" / "lovart_fold_retry_simple.jpg", ASSETS / "top_07_fold_round13.jpg")
    shutil.copy2(ASSETS / "top_02_collar.jpg", ASSETS / "top_08_collar.jpg")
    html = PREVIEW / "index.html"
    html.write_text(build_html(), encoding="utf-8")
    screenshot(html)
    print(html)


if __name__ == "__main__":
    main()

