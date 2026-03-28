from __future__ import annotations

import json
import os
import time
from dataclasses import dataclass
from pathlib import Path

import requests
from openpyxl import Workbook
from PIL import Image, ImageDraw, ImageFont, ImageOps
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
STEP6_DIR = ROOT / "step6_output"
MATERIALS_ROOT = STEP6_DIR / "素材包"
BUNDLE_DIR = MATERIALS_ROOT / "T01_bundle_2026-03-24"
ROUND7_DIR = MATERIALS_ROOT / "T01_round7_2026-03-24"
ROUND8_DIR = MATERIALS_ROOT / "T01_round8_2026-03-24"
ROUND9_DIR = MATERIALS_ROOT / "T01_round9_2026-03-24"
GEN_DIR = ROUND9_DIR / "01_generated"
COMPARE_DIR = ROUND9_DIR / "02_compare"
DEBUG_DIR = ROUND9_DIR / "03_debug"
PREVIEW_DIR = ROUND9_DIR / "04_web_preview"
STAGE_DIR = ROUND9_DIR / "05_stage"
PROFILE_DIR = ROOT / "browser_profiles" / "taobao_1688_manual"
EXPORT_XLSX = ROUND9_DIR / "T01_round9_alphashop_stack_retry_2026-03-24.xlsx"
EXPORT_MD = ROUND9_DIR / "T01_round9_alphashop_stack_retry_2026-03-24.md"

for folder in [ROUND9_DIR, GEN_DIR, COMPARE_DIR, DEBUG_DIR, PREVIEW_DIR, STAGE_DIR]:
    folder.mkdir(parents=True, exist_ok=True)

REQUEST_HEADERS = {"User-Agent": "Mozilla/5.0"}


@dataclass(frozen=True)
class Job:
    gid: str
    title: str
    ref_index: int
    product_override: Path
    note: str


JOBS: list[Job] = [
    Job(
        "G05B",
        "顶部第三排第一张 堆叠图 重试",
        6,
        ROUND7_DIR / "04_web_preview" / "assets" / "top_05_stack_round7.jpg",
        "用本地搭好的堆叠基底去锁结构，再让 AlphaShop 学参考页质感。",
    ),
    Job(
        "G07B",
        "顶部第四排第一张 散放图 重试",
        8,
        ROUND7_DIR / "04_web_preview" / "assets" / "top_07_scatter_round7.jpg",
        "用本地散放基底去锁结构，再让 AlphaShop 学参考页布料层次。",
    ),
]


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


def bundle_refs() -> dict[int, Path]:
    ref_dir = BUNDLE_DIR / "03_reference_source_images"
    ref_imgs = [p for p in sorted(ref_dir.iterdir()) if p.is_file() and p.suffix.lower() in {".jpg", ".jpeg", ".png", ".webp", ".avif"}]
    return {i + 1: p for i, p in enumerate(ref_imgs)}


def stage_files(ref_map: dict[int, Path]) -> dict[str, tuple[Path, Path]]:
    staged: dict[str, tuple[Path, Path]] = {}
    for job in JOBS:
        product_dst = STAGE_DIR / f"{job.gid}_product{job.product_override.suffix.lower()}"
        product_dst.write_bytes(job.product_override.read_bytes())
        ref_src = ref_map[job.ref_index]
        ref_dst = STAGE_DIR / f"{job.gid}_ref{ref_src.suffix.lower()}"
        ref_dst.write_bytes(ref_src.read_bytes())
        staged[job.gid] = (product_dst, ref_dst)
    return staged


def contain(img: Image.Image, size: tuple[int, int]) -> Image.Image:
    return ImageOps.contain(img.convert("RGB"), size, method=Image.Resampling.LANCZOS)


def compare_board(job: Job, ref_path: Path, gen_path: Path, product_path: Path) -> Path:
    canvas = Image.new("RGB", (2700, 1100), "#f7f4ef")
    draw = ImageDraw.Draw(canvas)
    draw.text((60, 40), f"{job.gid} 对比图", fill="#1d1d1f", font=FONT_H1)
    draw.text((60, 95), job.title, fill="#6f675f", font=FONT_BODY)
    draw.text((60, 130), job.note, fill="#6f675f", font=FONT_BODY)

    left = contain(Image.open(ref_path), (780, 820))
    mid = contain(Image.open(product_path), (780, 820))
    right = contain(Image.open(gen_path), (780, 820))
    canvas.paste(left, (80 + (780 - left.width) // 2, 200))
    canvas.paste(mid, (960 + (780 - mid.width) // 2, 200))
    canvas.paste(right, (1840 + (780 - right.width) // 2, 200))
    draw.text((80, 1040), f"左：参考图 #{job.ref_index}", fill="#1d1d1f", font=FONT_H2)
    draw.text((960, 1040), "中：本地结构基底", fill="#1d1d1f", font=FONT_H2)
    draw.text((1840, 1040), "右：AlphaShop 生成结果", fill="#1d1d1f", font=FONT_H2)

    out = COMPARE_DIR / f"{job.gid}_compare.png"
    canvas.save(out)
    return out


def write_preview(rows: list[dict[str, str]]) -> None:
    cards = []
    for row in rows:
        gen_rel = os.path.relpath(Path(row["result_file"]).resolve(), PREVIEW_DIR.resolve()).replace("\\", "/")
        cmp_rel = os.path.relpath(Path(row["compare_file"]).resolve(), PREVIEW_DIR.resolve()).replace("\\", "/")
        cards.append(
            f"""
            <section class="card">
              <h2>{row['gid']} {row['title']}</h2>
              <p>{row['note']}</p>
              <p class="small">状态：{row['status']}</p>
              <p class="small">会话：<a href="{row['session_url']}">{row['session_url']}</a></p>
              <div class="grid">
                <figure><img src="{gen_rel}" alt="{row['gid']}"><figcaption>生成结果</figcaption></figure>
                <figure><img src="{cmp_rel}" alt="{row['gid']} compare"><figcaption>参考 / 基底 / 结果 对比</figcaption></figure>
              </div>
            </section>
            """
        )
    html = f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>T01 Round9 AlphaShop 堆叠镜头重试</title>
  <style>
    body {{ margin:0; font-family:"Microsoft YaHei","PingFang SC",sans-serif; background:#f6f4ef; color:#1d1d1f; }}
    .wrap {{ width:min(1500px, calc(100vw - 48px)); margin:0 auto; padding:32px 0 56px; }}
    h1 {{ margin:0 0 12px; font-size:34px; }}
    .lead {{ color:#6e665e; line-height:1.7; margin:0 0 24px; }}
    .card {{ background:#fff; border-radius:20px; padding:24px; margin-bottom:24px; box-shadow:0 10px 30px rgba(0,0,0,.05); }}
    h2 {{ margin:0 0 8px; font-size:24px; }}
    .small {{ font-size:14px; color:#6e665e; word-break:break-all; }}
    .grid {{ margin-top:16px; display:grid; grid-template-columns:minmax(0, 0.8fr) minmax(0, 1.2fr); gap:18px; }}
    figure {{ margin:0; background:#faf8f3; padding:12px; border-radius:16px; }}
    img {{ width:100%; display:block; border-radius:12px; background:#fff; }}
    figcaption {{ padding-top:10px; color:#6e665e; font-size:14px; }}
    a {{ color:#16365d; }}
  </style>
</head>
<body>
  <div class="wrap">
    <h1>T01 Round9 AlphaShop 堆叠/散放镜头重试</h1>
    <p class="lead">这轮把本地已经搭好的结构图当作商品基底，再用参考页做风格模仿，目标是把镜头结构锁住，而不是让模型重新理解成单件白T。</p>
    {''.join(cards)}
  </div>
</body>
</html>
"""
    (PREVIEW_DIR / "index.html").write_text(html, encoding="utf-8")


def capture_preview_screenshot() -> None:
    preview = PREVIEW_DIR / "index.html"
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", headless=True)
        page = browser.new_page(viewport={"width": 1500, "height": 2400})
        page.goto(preview.resolve().as_uri(), wait_until="load")
        page.wait_for_timeout(2000)
        page.screenshot(path=str(PREVIEW_DIR / "preview_check.png"), full_page=True)
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
        "# T01 Round9 AlphaShop 堆叠/散放镜头重试",
        "",
        "- 时间：2026-03-24",
        "- 工具：AlphaShop 风格模仿",
        "- 重试策略：用本地结构基底锁定构图，再让平台学参考页质感",
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
          const targets = ['风格模仿', '椋庢牸妯′豢'];
          const els = Array.from(document.querySelectorAll('*'));
          const el = els.find(node => targets.includes((node.innerText || '').trim()));
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
        if "任务已结束" in body or "浠诲姟宸茬粨鏉?" in body:
            break
    return page.evaluate(
        """() => ({
          url: location.href,
          body: document.body.innerText,
          imgs: Array.from(document.querySelectorAll('img')).map((el,i)=>({i,src:(el.src||''),w:el.naturalWidth,h:el.naturalHeight}))
        })"""
    )


def select_result_url(debug: dict) -> str:
    candidates = [
        img for img in debug["imgs"]
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


def run_batch(staged: dict[str, tuple[Path, Path]], ref_map: dict[int, Path]) -> list[dict[str, str]]:
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
                product, ref_img = staged[job.gid]
                result_path = GEN_DIR / f"{job.gid}.png"
                debug_path = DEBUG_DIR / f"{job.gid}_debug.json"
                page_path = DEBUG_DIR / f"{job.gid}_page.png"
                compare_path = COMPARE_DIR / f"{job.gid}_compare.png"
                page = context.new_page()
                page.goto("https://create.alphashop.cn/", wait_until="domcontentloaded", timeout=120000)
                page.wait_for_timeout(6000)
                click_style_transfer(page)
                upload_two_files(page, product, ref_img)
                debug = wait_for_finish(page)
                debug_path.write_text(json.dumps(debug, ensure_ascii=False, indent=2), encoding="utf-8")
                page.screenshot(path=str(page_path), full_page=True)
                session_url = debug["url"]
                result_url = select_result_url(debug)
                if result_url:
                    download(result_url, result_path)
                    compare_board(job, ref_map[job.ref_index], result_path, product)
                    status = "已生成"
                else:
                    status = "失败：未抓到结果图"
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
                time.sleep(6)
        finally:
            context.close()
    return rows


def main() -> None:
    ref_map = bundle_refs()
    staged = stage_files(ref_map)
    rows = run_batch(staged, ref_map)
    write_preview(rows)
    capture_preview_screenshot()
    write_excel(rows)
    write_markdown(rows)


if __name__ == "__main__":
    main()
