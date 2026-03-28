from __future__ import annotations

import json
import shutil
from pathlib import Path

from openpyxl import Workbook
from PIL import Image, ImageChops, ImageDraw, ImageEnhance, ImageFilter, ImageFont, ImageOps


ROOT = Path(__file__).resolve().parent
STEP6_DIR = ROOT / "step6_output"
MATERIALS_ROOT = next(
    p for p in STEP6_DIR.iterdir() if p.is_dir() and not p.name.startswith(("T01_", "tmp", "oakln", "supplier"))
)
BUNDLE_DIR = MATERIALS_ROOT / "T01_bundle_2026-03-24"
ROUND1_DIR = MATERIALS_ROOT / "T01_round1_2026-03-24"
ROUND2_DIR = MATERIALS_ROOT / "T01_round2_2026-03-24"
GEN_DIR = ROUND2_DIR / "01_round2_generated"
COMPARE_DIR = ROUND2_DIR / "02_compare_after_opt"
PREVIEW_DIR = ROUND2_DIR / "03_web_preview"

for folder in [ROUND2_DIR, GEN_DIR, COMPARE_DIR, PREVIEW_DIR]:
    folder.mkdir(parents=True, exist_ok=True)


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


FONT_H1 = load_font(42, bold=True)
FONT_H2 = load_font(24, bold=True)
FONT_BODY = load_font(20)

BG = "#f7f4ef"
CARD = "#ffffff"
TEXT = "#1d1d1f"
MUTED = "#6f675e"
LINE = "#e9dece"


def contain(image: Image.Image, size: tuple[int, int]) -> Image.Image:
    return ImageOps.contain(image.convert("RGB"), size, method=Image.Resampling.LANCZOS)


def cover(image: Image.Image, size: tuple[int, int]) -> Image.Image:
    return ImageOps.fit(image.convert("RGB"), size, method=Image.Resampling.LANCZOS)


def canvas() -> Image.Image:
    return Image.new("RGB", (1080, 1080), BG)


def draw_head(draw: ImageDraw.ImageDraw, title: str, subtitle: str) -> None:
    draw.text((70, 58), title, fill=TEXT, font=FONT_H1)
    draw.text((70, 114), subtitle, fill=MUTED, font=FONT_BODY)


def add_shadow(base: Image.Image, img: Image.Image, xy: tuple[int, int]) -> None:
    shadow = Image.new("RGBA", (img.width + 26, img.height + 26), (0, 0, 0, 0))
    blk = Image.new("RGBA", img.size, (0, 0, 0, 55))
    shadow.paste(blk, (13, 13))
    shadow = shadow.filter(ImageFilter.GaussianBlur(12))
    base.paste(shadow, (xy[0] - 13, xy[1] - 13), shadow)
    base.paste(img, xy)


def save(img: Image.Image, name: str) -> Path:
    path = GEN_DIR / name
    img.save(path, quality=95)
    return path


pack_dir = BUNDLE_DIR / "01_upload_pack_v2"
main_dir = next(p for p in pack_dir.iterdir() if p.is_dir() and p.name.startswith("01_"))
color_dir = next(p for p in pack_dir.iterdir() if p.is_dir() and p.name.startswith("03_"))
ref_dir = BUNDLE_DIR / "03_reference_source_images"

main_imgs = sorted(p for p in main_dir.iterdir() if p.is_file())
color_imgs = sorted(p for p in color_dir.iterdir() if p.is_file())
ref_imgs = [p for p in sorted(ref_dir.iterdir()) if p.is_file() and p.suffix.lower() in {".jpg", ".jpeg", ".png", ".webp", ".avif"}]
ref_map = {i + 1: p for i, p in enumerate(ref_imgs)}

hero_white = main_imgs[0]
collar = main_imgs[3]
fold = main_imgs[4]
rack = main_imgs[5]
white = next(p for p in color_imgs if "白色" in p.name)
black = next(p for p in color_imgs if "黑色" in p.name)
gray = next(p for p in color_imgs if "灰色" in p.name)
beige = next(p for p in color_imgs if "米白" in p.name)


def crop_background(image: Image.Image) -> Image.Image:
    bg = Image.new("RGB", image.size, (255, 255, 255))
    diff = ImageChops.difference(image.convert("RGB"), bg)
    box = diff.getbbox()
    return image.crop(box) if box else image


def make_g05_v2() -> Path:
    out = canvas()
    draw = ImageDraw.Draw(out)
    draw_head(draw, "面料与色组图 v2", "改成更接近参考页的层叠排布和阴影关系")
    draw.rectangle((90, 220, 990, 940), fill="#d8b396")
    stack_specs = [
        (beige, (140, 500), 0.97),
        (white, (250, 440), 0.99),
        (gray, (380, 380), 1.00),
        (black, (520, 320), 1.02),
    ]
    for src, (x, y), factor in stack_specs:
        img = crop_background(Image.open(src))
        img = contain(img, (350, 350))
        img = ImageEnhance.Brightness(img).enhance(factor)
        add_shadow(out, img, (x, y))
    return save(out, "G05_v2_面料与色组图_折叠堆叠.jpg")


def make_g09_v2() -> Path:
    out = canvas()
    draw = ImageDraw.Draw(out)
    draw_head(draw, "袖口与下摆近景 v2", "从主图原图里裁出更像局部结构图的区域")
    src = Image.open(hero_white)
    crop = src.crop((40, 240, 560, 900))
    crop = cover(crop, (860, 760))
    crop = ImageEnhance.Contrast(crop).enhance(1.06)
    out.paste(crop, ((1080 - crop.width) // 2, 220))
    return save(out, "G09_v2_袖口与下摆近景.jpg")


def make_g13_v2() -> Path:
    out = canvas()
    draw = ImageDraw.Draw(out)
    draw_head(draw, "面料表面近景 v2", "进一步压近构图，强化表面纹理和柔光感")
    src = Image.open(fold)
    crop = src.crop((180, 180, src.width - 100, src.height - 40))
    crop = cover(crop, (860, 760))
    crop = ImageEnhance.Contrast(crop).enhance(1.12)
    crop = ImageEnhance.Sharpness(crop).enhance(1.2)
    out.paste(crop, ((1080 - crop.width) // 2, 220))
    return save(out, "G13_v2_面料表面近景.jpg")


def make_g14_v2() -> Path:
    out = canvas()
    draw = ImageDraw.Draw(out)
    draw_head(draw, "肌理微距图 v2", "从领口区域裁出更像参考页的结构和肌理细节")
    src = Image.open(collar)
    crop = src.crop((180, 180, src.width - 140, src.height - 50))
    crop = cover(crop, (860, 760))
    crop = ImageEnhance.Contrast(crop).enhance(1.14)
    crop = ImageEnhance.Sharpness(crop).enhance(1.18)
    out.paste(crop, ((1080 - crop.width) // 2, 220))
    return save(out, "G14_v2_肌理微距图.jpg")


generated = {
    "G05": make_g05_v2(),
    "G09": make_g09_v2(),
    "G13": make_g13_v2(),
    "G14": make_g14_v2(),
}


compare_specs = {
    "G05": (51, "第二轮后更接近参考页的层叠和暖底氛围，可直接替换第一轮"),
    "G09": (55, "第二轮比第一轮更像局部结构图，但仍然没有手部入镜"),
    "G13": (40, "第二轮已经比第一轮更接近面料表面特写"),
    "G14": (41, "第二轮已经比第一轮更接近微距肌理，但还不是真正织物微距"),
}


def make_compare(gid: str, ref_idx: int, note: str) -> Path:
    out = Image.new("RGB", (1800, 1080), BG)
    draw = ImageDraw.Draw(out)
    draw.text((60, 40), f"{gid} 第二轮优化对比", fill=TEXT, font=FONT_H1)
    draw.text((60, 100), note, fill=MUTED, font=FONT_BODY)
    ref_img = contain(Image.open(ref_map[ref_idx]), (780, 780))
    out.paste(ref_img, (80 + (780 - ref_img.width) // 2, 180))
    draw.text((80, 970), f"左：参考图 #{ref_idx}", fill=TEXT, font=FONT_H2)
    cur_img = contain(Image.open(generated[gid]), (780, 780))
    out.paste(cur_img, (940 + (780 - cur_img.width) // 2, 180))
    draw.text((940, 970), "右：第二轮优化图", fill=TEXT, font=FONT_H2)
    path = COMPARE_DIR / f"{gid}_opt_compare.jpg"
    out.save(path, quality=95)
    return path


compare_paths = {gid: make_compare(gid, ref_idx, note) for gid, (ref_idx, note) in compare_specs.items()}


preview_html = f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>T01 Round2 优化预览</title>
  <style>
    body {{ margin: 0; font-family: Arial, "Microsoft YaHei", sans-serif; background: #f7f4ef; color: #1d1d1f; }}
    .wrap {{ max-width: 1380px; margin: 0 auto; padding: 30px; }}
    h1 {{ font-size: 34px; margin: 0 0 8px; }}
    p {{ color: #6f675e; line-height: 1.7; }}
    .grid {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 18px; }}
    .card {{ background: white; border: 1px solid #e9dece; padding: 12px; }}
    img {{ width: 100%; display: block; }}
    .cap {{ margin-top: 8px; font-size: 12px; color: #6f675e; }}
  </style>
</head>
<body>
  <div class="wrap">
    <h1>T01 Round2 自动优化预览</h1>
    <p>这一轮只继续优化本地还能提升的细节和结构镜头。高优先级缺口仍然是同模特上身图、坐姿图和全身搭配图，这部分已经到需要平台生成的阶段。</p>
    <div class="grid">
      {''.join(f'<div class="card"><img src="../01_round2_generated/{path.name}" alt=""><div class="cap">{gid}</div></div>' for gid, path in generated.items())}
    </div>
  </div>
</body>
</html>
"""
(PREVIEW_DIR / "index.html").write_text(preview_html, encoding="utf-8")

wb = Workbook()
ws = wb.active
ws.title = "optimized"
ws.append(["清单ID", "第二轮输出文件", "说明"])
for gid, path in generated.items():
    ws.append([gid, str(path), compare_specs[gid][1]])

ws2 = wb.create_sheet("compare")
ws2.append(["清单ID", "对比图", "说明"])
for gid, path in compare_paths.items():
    ws2.append([gid, str(path), compare_specs[gid][1]])

workbook_path = ROUND2_DIR / "T01_round2_opt_2026-03-24.xlsx"
wb.save(workbook_path)

summary = {
    "generated": {gid: str(path) for gid, path in generated.items()},
    "compare": {gid: str(path) for gid, path in compare_paths.items()},
    "preview_html": str(PREVIEW_DIR / "index.html"),
    "workbook": str(workbook_path),
    "still_blocked": ["G02", "G03", "G12", "G15", "G16", "G17"],
}
(ROUND2_DIR / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

readme = "\n".join(
    [
        "# T01 Round2 自动优化结果",
        "",
        f"- 第二轮优化图目录：`{GEN_DIR}`",
        f"- 第二轮对比目录：`{COMPARE_DIR}`",
        f"- 预览页：`{PREVIEW_DIR / 'index.html'}`",
        f"- 工作簿：`{workbook_path}`",
        "",
        "这轮优化完成后，仍然无法在本地继续推进的镜头：",
        "- G02",
        "- G03",
        "- G12",
        "- G15",
        "- G16",
        "- G17",
        "",
        "这些镜头已经到需要平台生成或重绘的阶段。",
    ]
)
(ROUND2_DIR / "README.md").write_text(readme, encoding="utf-8")

print(workbook_path)
print(PREVIEW_DIR / "index.html")
