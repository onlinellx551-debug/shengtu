from __future__ import annotations

import json
import time
from dataclasses import dataclass
import os
from pathlib import Path

import requests
from openpyxl import Workbook
from PIL import Image, ImageDraw, ImageFont, ImageOps
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
STEP6_DIR = ROOT / "step6_output"
MATERIALS_ROOT = next(p for p in STEP6_DIR.iterdir() if p.is_dir() and (p / "T01_bundle_2026-03-24").exists())
BUNDLE_DIR = MATERIALS_ROOT / "T01_bundle_2026-03-24"
ROUND3_DIR = MATERIALS_ROOT / "T01_round3_2026-03-24"
GEN_DIR = ROUND3_DIR / "01_generated"
COMPARE_DIR = ROUND3_DIR / "02_compare"
DEBUG_DIR = ROUND3_DIR / "03_debug"
PREVIEW_DIR = ROUND3_DIR / "04_web_preview"
STAGE_DIR = STEP6_DIR / "gen_stage"
PROFILE_DIR = ROOT / "browser_profiles" / "taobao_1688_manual"
RETRY_DIR = STEP6_DIR / "alphashop_persistent_test"
EXPORT_XLSX = ROUND3_DIR / "T01_round3_alphashop_generate_2026-03-24.xlsx"
EXPORT_MD = ROUND3_DIR / "T01_round3_alphashop_generate_2026-03-24.md"

for folder in [ROUND3_DIR, GEN_DIR, COMPARE_DIR, DEBUG_DIR, PREVIEW_DIR, STAGE_DIR]:
    folder.mkdir(parents=True, exist_ok=True)


REQUEST_HEADERS = {"User-Agent": "Mozilla/5.0"}
STYLE_TRANSFER_TEXT = "风格模仿"


@dataclass(frozen=True)
class Job:
    gid: str
    title: str
    ref_index: int
    note: str


JOBS: list[Job] = [
    Job("G02", "上半身正面模特图", 3, "主图模特图 1，优先看是否保住参考页的成熟感和构图裁切。"),
    Job("G03", "上半身三分之二模特图", 4, "主图模特图 2，优先看是否保住参考页的裁切方式。"),
    Job("G12", "坐姿生活方式图", 39, "中段生活方式图，优先看氛围和坐姿场景。"),
    Job("G15", "全身搭配图 1", 42, "浅色裤装全身 look。"),
    Job("G16", "全身搭配图 2", 43, "外套层次全身 look。"),
    Job("G17", "全身搭配图 3", 44, "白 T 作为内搭的层次图。"),
]

REPLACEMENTS = {
    "G12": {
        "image": "g12_retry_en_result.png",
        "json": "g12_retry_en.json",
        "page": "g12_retry_en.png",
    },
    "G16": {
        "image": "G16_retry_en_result.png",
        "json": "G16_retry_en.json",
        "page": "G16_retry_en.png",
    },
    "G17": {
        "image": "G17_retry_en_result.png",
        "json": "G17_retry_en.json",
        "page": "G17_retry_en.png",
    },
}


def load_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    candidates = [
        Path(r"C:\Windows\Fonts\msyhbd.ttc") if bold else Path(r"C:\Windows\Fonts\msyh.ttc"),
        Path(r"C:\Windows\Fonts\simhei.ttf"),
        Path(r"C:\Windows\Fonts\arial.ttf"),
    ]
    for path in candidates:
        if path.exists():
            try:
                return ImageFont.truetype(str(path), size=size)
            except OSError:
                continue
    return ImageFont.load_default()


FONT_H1 = load_font(38, bold=True)
FONT_H2 = load_font(26, bold=True)
FONT_BODY = load_font(20)


def bundle_inputs() -> tuple[Path, dict[int, Path]]:
    pack_dir = BUNDLE_DIR / "01_upload_pack_v2"
    main_dir = next(p for p in pack_dir.iterdir() if p.is_dir() and p.name.startswith("01_"))
    hero = sorted(p for p in main_dir.iterdir() if p.is_file())[0]
    ref_dir = BUNDLE_DIR / "03_reference_source_images"
    ref_imgs = [p for p in sorted(ref_dir.iterdir()) if p.is_file() and p.suffix.lower() in {".jpg", ".jpeg", ".png", ".webp", ".avif"}]
    ref_map = {i + 1: p for i, p in enumerate(ref_imgs)}
    return hero, ref_map


def stage_files(hero_src: Path, ref_map: dict[int, Path]) -> tuple[Path, dict[str, Path]]:
    product_dst = STAGE_DIR / "product_white_tee.jpg"
    if not product_dst.exists():
        product_dst.write_bytes(hero_src.read_bytes())
    ref_stage: dict[str, Path] = {}
    for job in JOBS:
        dst = STAGE_DIR / f"{job.gid}_ref{ref_map[job.ref_index].suffix.lower()}"
        if not dst.exists():
            dst.write_bytes(ref_map[job.ref_index].read_bytes())
        ref_stage[job.gid] = dst
    return product_dst, ref_stage


def apply_manual_replacements(ref_map: dict[int, Path]) -> None:
    for job in JOBS:
        replacement = REPLACEMENTS.get(job.gid)
        if not replacement:
            continue
        image_src = RETRY_DIR / replacement["image"]
        json_src = RETRY_DIR / replacement["json"]
        page_src = RETRY_DIR / replacement["page"]
        image_dst = GEN_DIR / f"{job.gid}.png"
        json_dst = DEBUG_DIR / f"{job.gid}_debug.json"
        page_dst = DEBUG_DIR / f"{job.gid}_page.png"
        if image_src.exists():
            image_dst.write_bytes(image_src.read_bytes())
            compare_board(job, ref_map[job.ref_index], image_dst)
        if json_src.exists():
            json_dst.write_bytes(json_src.read_bytes())
        if page_src.exists():
            page_dst.write_bytes(page_src.read_bytes())


def contain(img: Image.Image, size: tuple[int, int]) -> Image.Image:
    return ImageOps.contain(img.convert("RGB"), size, method=Image.Resampling.LANCZOS)


def compare_board(job: Job, ref_path: Path, gen_path: Path) -> Path:
    canvas = Image.new("RGB", (1800, 1100), "#f7f4ef")
    draw = ImageDraw.Draw(canvas)
    draw.text((60, 40), f"{job.gid} 对比图", fill="#1d1d1f", font=FONT_H1)
    draw.text((60, 95), job.title, fill="#6f675f", font=FONT_BODY)
    draw.text((60, 130), job.note, fill="#6f675f", font=FONT_BODY)

    left = contain(Image.open(ref_path), (780, 820))
    right = contain(Image.open(gen_path), (780, 820))
    canvas.paste(left, (80 + (780 - left.width) // 2, 200))
    canvas.paste(right, (940 + (780 - right.width) // 2, 200))
    draw.text((80, 1040), f"左：参考图 #{job.ref_index}", fill="#1d1d1f", font=FONT_H2)
    draw.text((940, 1040), "右：AlphaShop 生成结果", fill="#1d1d1f", font=FONT_H2)

    out = COMPARE_DIR / f"{job.gid}_compare.png"
    canvas.save(out)
    return out


def write_preview(rows: list[dict[str, str]]) -> None:
    blocks = []
    for row in rows:
        gen_rel = os.path.relpath(Path(row["result_file"]).resolve(), PREVIEW_DIR.resolve()).replace("\\", "/")
        cmp_rel = os.path.relpath(Path(row["compare_file"]).resolve(), PREVIEW_DIR.resolve()).replace("\\", "/")
        blocks.append(
            f"""
            <section class="card">
              <div class="meta">
                <h2>{row['gid']} {row['title']}</h2>
                <p>{row['note']}</p>
                <p class="small">状态：{row['status']}</p>
                <p class="small">会话：<a href="{row['session_url']}">{row['session_url']}</a></p>
              </div>
              <div class="grid">
                <figure>
                  <img src="{gen_rel}" alt="{row['gid']} generated">
                  <figcaption>生成结果</figcaption>
                </figure>
                <figure>
                  <img src="{cmp_rel}" alt="{row['gid']} compare">
                  <figcaption>参考对比</figcaption>
                </figure>
              </div>
            </section>
            """
        )

    html = f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>T01 Round3 AlphaShop 结果预览</title>
  <style>
    body {{
      margin: 0;
      font-family: "Microsoft YaHei", "PingFang SC", sans-serif;
      background: #f6f4ef;
      color: #1d1d1f;
    }}
    .wrap {{
      width: min(1400px, calc(100vw - 48px));
      margin: 0 auto;
      padding: 32px 0 56px;
    }}
    h1 {{
      margin: 0 0 10px;
      font-size: 34px;
    }}
    .lead {{
      color: #6e665e;
      margin: 0 0 24px;
      line-height: 1.7;
    }}
    .card {{
      background: #fff;
      border-radius: 20px;
      padding: 24px;
      margin: 0 0 24px;
      box-shadow: 0 10px 30px rgba(0,0,0,.05);
    }}
    .meta h2 {{
      margin: 0 0 8px;
      font-size: 24px;
    }}
    .meta p {{
      margin: 0 0 8px;
      line-height: 1.7;
    }}
    .small {{
      font-size: 14px;
      color: #6e665e;
      word-break: break-all;
    }}
    .grid {{
      margin-top: 16px;
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 18px;
    }}
    figure {{
      margin: 0;
      background: #faf8f3;
      padding: 12px;
      border-radius: 16px;
    }}
    img {{
      width: 100%;
      display: block;
      border-radius: 12px;
      background: #fff;
    }}
    figcaption {{
      padding-top: 10px;
      color: #6e665e;
      font-size: 14px;
    }}
    a {{
      color: #16365d;
    }}
    @media (max-width: 900px) {{
      .grid {{
        grid-template-columns: 1fr;
      }}
    }}
  </style>
</head>
<body>
  <div class="wrap">
    <h1>T01 Round3 AlphaShop 高优先级镜头</h1>
    <p class="lead">这一页只展示本轮 6 张同模特/同镜头语言的高优先级结果。左侧信息保留会话链接，右侧同时放生成结果和参考对比，方便你直接判断哪几张已经能上架。</p>
    {''.join(blocks)}
  </div>
</body>
</html>
"""
    (PREVIEW_DIR / "index.html").write_text(html, encoding="utf-8")


def capture_preview_screenshot() -> None:
    preview = PREVIEW_DIR / "index.html"
    screenshot = PREVIEW_DIR / "preview_check.png"
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", headless=True)
        page = browser.new_page(viewport={"width": 1440, "height": 2200})
        page.goto(preview.resolve().as_uri(), wait_until="load")
        page.wait_for_timeout(2000)
        page.screenshot(path=str(screenshot), full_page=True)
        browser.close()


def write_excel(rows: list[dict[str, str]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "结果汇总"
    headers = ["gid", "title", "ref_index", "status", "note", "session_url", "result_url", "result_file", "compare_file", "page_debug"]
    ws.append(headers)
    for row in rows:
        ws.append([row.get(h, "") for h in headers])
    for col in ws.columns:
        width = max(len(str(cell.value or "")) for cell in col) + 2
        ws.column_dimensions[col[0].column_letter].width = min(max(width, 12), 60)
    wb.save(EXPORT_XLSX)


def write_markdown(rows: list[dict[str, str]]) -> None:
    lines = [
        "# T01 Round3 AlphaShop 生成结果",
        "",
        "- 时间：2026-03-24",
        "- 工具：AlphaShop 风格模仿",
        "- 商品：T01 厚实不透白 T",
        "- 目标镜头：G02 / G03 / G12 / G15 / G16 / G17",
        "",
        "## 结果",
        "",
    ]
    for row in rows:
        lines.append(f"- {row['gid']} {row['title']}：{row['status']}，结果图 `{row['result_file']}`，对比图 `{row['compare_file']}`")
    EXPORT_MD.write_text("\n".join(lines), encoding="utf-8")


def click_style_transfer(page) -> None:
    page.evaluate(
        """() => {
          const target = '风格模仿';
          const els = Array.from(document.querySelectorAll('*'));
          const el = els.find(node => ((node.innerText || '').trim()) === target);
          if (el) el.click();
        }"""
    )


def upload_two_files(page, product: Path, refimg: Path) -> None:
    page.wait_for_timeout(3000)
    buttons = page.locator("button")
    with page.expect_file_chooser() as fc1:
        buttons.nth(2).click()
    fc1.value.set_files(str(product))
    page.wait_for_timeout(4000)
    buttons = page.locator("button")
    with page.expect_file_chooser() as fc2:
        buttons.nth(2).click()
    fc2.value.set_files(str(refimg))
    page.wait_for_timeout(8000)
    buttons = page.locator("button")
    buttons.nth(2).click()


def wait_for_finish(page, timeout_s: int = 180) -> dict:
    for _ in range(timeout_s // 5):
        page.wait_for_timeout(5000)
        body = page.locator("body").inner_text()
        if "任务已结束" in body:
            break
    return page.evaluate(
        """() => ({
          url: location.href,
          body: document.body.innerText,
          imgs: Array.from(document.querySelectorAll('img')).map((el,i)=>({i,src:(el.src||''),w:el.naturalWidth,h:el.naturalHeight})),
          buttons: Array.from(document.querySelectorAll('button')).map((el,i)=>({i,text:(el.innerText||'').trim(),disabled:el.disabled,cls:el.className})).filter(x=>x.text)
        })"""
    )


def select_result_url(debug: dict) -> str:
    candidates = [
        img
        for img in debug["imgs"]
        if "cbu_global_ai_agent" in img["src"] and img["w"] >= 700 and img["h"] >= 700
    ]
    if not candidates:
        candidates = [img for img in debug["imgs"] if img["w"] >= 700 and img["h"] >= 700]
    if not candidates:
        return ""
    best = max(candidates, key=lambda x: x["w"] * x["h"])
    return best["src"]


def download(url: str, out_path: Path) -> None:
    response = requests.get(url, headers=REQUEST_HEADERS, timeout=60)
    response.raise_for_status()
    out_path.write_bytes(response.content)


def run_batch(product: Path, ref_stage: dict[str, Path], ref_map: dict[int, Path]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            channel="chrome",
            headless=True,
            viewport={"width": 1500, "height": 2200},
            args=["--disable-blink-features=AutomationControlled"],
        )
        try:
            for job in JOBS:
                result_path = GEN_DIR / f"{job.gid}.png"
                debug_path = DEBUG_DIR / f"{job.gid}_debug.json"
                page_path = DEBUG_DIR / f"{job.gid}_page.png"
                compare_path = COMPARE_DIR / f"{job.gid}_compare.png"

                if result_path.exists() and debug_path.exists() and compare_path.exists():
                    debug = json.loads(debug_path.read_text(encoding="utf-8"))
                    rows.append(
                        {
                            "gid": job.gid,
                            "title": job.title,
                            "ref_index": str(job.ref_index),
                            "status": "已生成（复用已有结果）",
                            "note": job.note,
                            "session_url": debug.get("url", ""),
                            "result_url": select_result_url(debug),
                            "result_file": str(result_path),
                            "compare_file": str(compare_path),
                            "page_debug": str(page_path),
                        }
                    )
                    continue

                page = context.new_page()
                page.goto("https://create.alphashop.cn/", wait_until="domcontentloaded", timeout=120000)
                page.wait_for_timeout(6000)
                click_style_transfer(page)
                upload_two_files(page, product, ref_stage[job.gid])
                debug = wait_for_finish(page)
                debug_path.write_text(json.dumps(debug, ensure_ascii=False, indent=2), encoding="utf-8")
                page.screenshot(path=str(page_path), full_page=True)
                session_url = debug["url"]
                result_url = select_result_url(debug)

                if result_url:
                    download(result_url, result_path)
                    compare_board(job, ref_map[job.ref_index], result_path)
                    status = "已生成"
                else:
                    status = "失败：未抓到结果图直链"

                rows.append(
                    {
                        "gid": job.gid,
                        "title": job.title,
                        "ref_index": str(job.ref_index),
                        "status": status,
                        "note": job.note,
                        "session_url": session_url,
                        "result_url": result_url,
                        "result_file": str(result_path),
                        "compare_file": str(compare_path),
                        "page_debug": str(page_path),
                    }
                )
                page.close()
                time.sleep(5)
        finally:
            context.close()
    return rows


def main() -> None:
    hero_src, ref_map = bundle_inputs()
    product, ref_stage = stage_files(hero_src, ref_map)
    apply_manual_replacements(ref_map)
    rows = run_batch(product, ref_stage, ref_map)
    write_preview(rows)
    capture_preview_screenshot()
    write_excel(rows)
    write_markdown(rows)


if __name__ == "__main__":
    main()
