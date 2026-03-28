from __future__ import annotations

import json
import shutil
from io import BytesIO
from pathlib import Path

import pandas as pd
import requests
from PIL import Image, ImageColor, ImageDraw, ImageFilter, ImageFont, ImageOps


ROOT = Path(__file__).resolve().parent
STEP6_DIR = ROOT / "step6_output"
EXPORT_DATE = "2026-03-28"
CASE_DIR = STEP6_DIR / f"S01_case_materials_{EXPORT_DATE}"
RAW_DIR = CASE_DIR / "01_raw"
CASE_IMAGE_DIR = CASE_DIR / "02_case_images"
DOC_DIR = CASE_DIR / "03_docs"
PREVIEW_HTML = CASE_DIR / "04_preview.html"
JSON_PATH = STEP6_DIR / f"S01_case_materials_{EXPORT_DATE}.json"
MD_PATH = STEP6_DIR / f"S01_case_materials_{EXPORT_DATE}.md"
XLSX_PATH = STEP6_DIR / f"S01_case_materials_{EXPORT_DATE}.xlsx"

SOURCE_IMAGE_URL = "https://cbu01.alicdn.com/img/ibank/O1CN01XDXUss1CHEp7L6lK5_!!2458340055-0-cib.jpg"
SOURCE_LINK = "https://detail.1688.com/offer/771706278919.html"
BACKUP_LINK = "https://item.taobao.com/item.htm?id=826439401178"


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


for folder in [CASE_DIR, RAW_DIR, CASE_IMAGE_DIR, DOC_DIR]:
    ensure_dir(folder)


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


FONT_TITLE = load_font(46, bold=True)
FONT_SUB = load_font(24)
FONT_BODY = load_font(26)
FONT_SMALL = load_font(20)

BG = ImageColor.getrgb("#f5f3ef")
CARD = ImageColor.getrgb("#ffffff")
TEXT = ImageColor.getrgb("#1f1f1f")
MUTED = ImageColor.getrgb("#666666")
ACCENT = ImageColor.getrgb("#d7d0c5")
LINE = ImageColor.getrgb("#e7e1d8")


def wrap_text(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, width: int) -> list[str]:
    chars = list(text)
    lines: list[str] = []
    current = ""
    for ch in chars:
        probe = current + ch
        if draw.textlength(probe, font=font) <= width:
            current = probe
        else:
            if current:
                lines.append(current)
            current = ch
    if current:
        lines.append(current)
    return lines


def contain(image: Image.Image, size: tuple[int, int]) -> Image.Image:
    return ImageOps.contain(image.convert("RGB"), size, method=Image.Resampling.LANCZOS)


def fit(image: Image.Image, size: tuple[int, int]) -> Image.Image:
    return ImageOps.fit(image.convert("RGB"), size, method=Image.Resampling.LANCZOS)


def add_shadow(canvas: Image.Image, box: tuple[int, int, int, int]) -> None:
    x0, y0, x1, y1 = box
    shadow = Image.new("RGBA", (x1 - x0 + 28, y1 - y0 + 28), (0, 0, 0, 0))
    shadow_box = Image.new("RGBA", (x1 - x0, y1 - y0), (0, 0, 0, 55))
    shadow.paste(shadow_box, (14, 14))
    shadow = shadow.filter(ImageFilter.GaussianBlur(16))
    canvas.paste(shadow, (x0 - 14, y0 - 14), shadow)


def save_card(name: str, image: Image.Image) -> Path:
    path = CASE_IMAGE_DIR / f"{name}.png"
    image.save(path, quality=96)
    return path


def download_source_image() -> Path:
    out = RAW_DIR / "s01_main_source.jpg"
    r = requests.get(SOURCE_IMAGE_URL, headers={"User-Agent": "Mozilla/5.0"}, timeout=30)
    r.raise_for_status()
    out.write_bytes(r.content)
    return out


def build_hero_card(source: Image.Image) -> dict[str, str]:
    canvas = Image.new("RGB", (1080, 1080), BG)
    draw = ImageDraw.Draw(canvas)
    art = contain(source, (840, 760))
    add_shadow(canvas, (120, 120, 960, 860))
    canvas.paste(Image.new("RGB", (840, 740), CARD), (120, 120))
    canvas.paste(art, (120 + (840 - art.width) // 2, 120 + (740 - art.height) // 2))
    draw.text((80, 910), "S01 White Shirt", fill=TEXT, font=FONT_TITLE)
    draw.text((80, 972), "DP no-iron regular-collar white shirt for office-casual and interview scenes.", fill=MUTED, font=FONT_SUB)
    path = save_card("01_hero_white_shirt", canvas)
    return {
        "成品文件": str(path),
        "类型": "纯图",
        "模块": "首屏主图",
        "来源": "1688主图原图",
        "生成方式": "原图二次排版",
        "说明": "用于首屏和SKU白衬衫主视觉。",
    }


def build_crop_card(name: str, title: str, subtitle: str, source: Image.Image, box: tuple[int, int, int, int]) -> dict[str, str]:
    crop = source.crop(box)
    canvas = Image.new("RGB", (1080, 1080), BG)
    draw = ImageDraw.Draw(canvas)
    art = contain(crop, (860, 760))
    add_shadow(canvas, (110, 120, 970, 860))
    canvas.paste(Image.new("RGB", (860, 740), CARD), (110, 120))
    canvas.paste(art, (110 + (860 - art.width) // 2, 120 + (740 - art.height) // 2))
    draw.text((80, 910), title, fill=TEXT, font=FONT_TITLE)
    draw.text((80, 972), subtitle, fill=MUTED, font=FONT_SUB)
    path = save_card(name, canvas)
    return {
        "成品文件": str(path),
        "类型": "纯图",
        "模块": title,
        "来源": "1688主图原图裁切",
        "生成方式": "局部放大",
        "说明": subtitle,
    }


def build_feature_card(name: str, kicker: str, title: str, bullets: list[str], source: Image.Image) -> dict[str, str]:
    canvas = Image.new("RGB", (1080, 1080), BG)
    draw = ImageDraw.Draw(canvas)
    draw.rounded_rectangle((58, 58, 1022, 1022), radius=28, fill=CARD)
    thumb = fit(source, (360, 480))
    canvas.paste(thumb, (96, 180))
    draw.text((510, 146), kicker, fill=MUTED, font=FONT_SMALL)
    draw.text((510, 182), title, fill=TEXT, font=FONT_TITLE)
    draw.line((510, 256, 938, 256), fill=ACCENT, width=2)
    y = 318
    for bullet in bullets:
        lines = wrap_text(draw, bullet, FONT_BODY, 420)
        draw.text((520, y), "-", fill=TEXT, font=FONT_BODY)
        inner_y = y
        for line in lines:
            draw.text((550, inner_y), line, fill=TEXT, font=FONT_BODY)
            inner_y += 40
        y = inner_y + 18
    path = save_card(name, canvas)
    return {
        "成品文件": str(path),
        "类型": "营销图",
        "模块": kicker,
        "来源": "1688主图原图+结构化文案",
        "生成方式": "本地营销卡片",
        "说明": title,
    }


def build_spec_card(name: str, spec_rows: list[tuple[str, str]]) -> dict[str, str]:
    canvas = Image.new("RGB", (1080, 1080), BG)
    draw = ImageDraw.Draw(canvas)
    draw.rounded_rectangle((70, 70, 1010, 1010), radius=30, fill=CARD)
    draw.text((110, 120), "Specs & Selling Notes", fill=TEXT, font=FONT_TITLE)
    draw.text((110, 180), "Built from step-3/step-4 sourcing data for this exact user-requested link.", fill=MUTED, font=FONT_SUB)
    y = 270
    for label, value in spec_rows:
        draw.rounded_rectangle((110, y - 12, 970, y + 72), radius=18, fill=BG, outline=LINE)
        draw.text((140, y), label, fill=MUTED, font=FONT_BODY)
        lines = wrap_text(draw, value, FONT_BODY, 520)
        inner_y = y
        for line in lines:
            draw.text((380, inner_y), line, fill=TEXT, font=FONT_BODY)
            inner_y += 36
        y = inner_y + 26
    path = save_card(name, canvas)
    return {
        "成品文件": str(path),
        "类型": "信息图",
        "模块": "规格说明",
        "来源": "第3步选品+第4步货源信息",
        "生成方式": "本地信息图",
        "说明": "用于详情页底部规格与风险提示承接。",
    }


def build_html(image_rows: list[dict[str, str]]) -> str:
    imgs = "\n".join(
        f'      <section class="card"><img src="./02_case_images/{Path(row["成品文件"]).name}" alt="{row["模块"]}"><p>{row["模块"]}</p></section>'
        for row in image_rows
    )
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>S01 Step6 Preview</title>
  <style>
    * {{ box-sizing: border-box; }}
    body {{ margin: 0; font-family: "Microsoft YaHei", sans-serif; background: linear-gradient(180deg, #f6f4ef 0%, #ece7de 100%); color: #1f1f1f; }}
    header {{ padding: 32px 20px 12px; text-align: center; }}
    header h1 {{ margin: 0; font-size: 30px; }}
    header p {{ margin: 10px auto 0; max-width: 760px; color: #666; line-height: 1.7; }}
    .grid {{ max-width: 1180px; margin: 0 auto; padding: 24px 16px 48px; display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 18px; }}
    .card {{ background: rgba(255,255,255,0.88); border-radius: 20px; overflow: hidden; box-shadow: 0 18px 42px rgba(0,0,0,0.08); }}
    img {{ width: 100%; display: block; }}
    p {{ margin: 0; padding: 14px 16px 18px; font-size: 14px; color: #4c4c4c; }}
  </style>
</head>
<body>
  <header>
    <h1>S01 第六步详情素材预览</h1>
    <p>基于用户指定的 1688 链接 {SOURCE_LINK} 生成。页面原始详情受防爬限制，本次采用可直接下载的主图原图、局部裁切和结构化卖点卡完成可上手的详情页素材包。</p>
  </header>
  <main class="grid">
{imgs}
  </main>
</body>
</html>"""


def main() -> None:
    raw_image_path = download_source_image()
    source = Image.open(BytesIO(raw_image_path.read_bytes())).convert("RGB")
    shutil.copy2(ROOT / "step4_output" / "alphashop_images" / "S01_main.png", RAW_DIR / "s01_main_thumb.png")
    shutil.copy2(ROOT / "step4_output" / "alphashop_images" / "S01_backup.png", RAW_DIR / "s01_backup_thumb.png")

    image_rows: list[dict[str, str]] = []
    image_rows.append(build_hero_card(source))
    image_rows.append(build_crop_card("02_collar_detail", "Collar Detail", "强调常规领线条和商务通勤识别度。", source, (220, 70, 610, 360)))
    image_rows.append(build_crop_card("03_placket_buttons", "Placket & Buttons", "用于承接门襟、扣位和前幅整洁感。", source, (250, 180, 560, 760)))
    image_rows.append(build_crop_card("04_cuff_sleeve", "Sleeve & Cuff", "补足长袖衬衫在细节页里需要的袖口信息。", source, (120, 240, 360, 760)))
    image_rows.append(build_crop_card("05_fabric_texture", "Fabric Texture", "放大布面与褶皱控制，配合免烫卖点说明。", source, (300, 120, 700, 500)))
    image_rows.append(
        build_feature_card(
            "06_non_iron_card",
            "No-Iron DP",
            "面向日本通勤场景的省心卖点",
            [
                "DP免烫处理，适合出勤、入职、面试等需要整洁观感的场景。",
                "抗皱和易打理是这条链接在第4步命中的核心词。",
                "建议详情页文案避免夸大，保留“建议打样复核”的提示。",
            ],
            source,
        )
    )
    image_rows.append(
        build_feature_card(
            "07_business_scene_card",
            "Office Casual",
            "白衬衫作为稳定基础款的搭配逻辑",
            [
                "白色、常规领、长袖三点组合，适合办公室休闲风格线。",
                "适配西裤、针织开衫、轻西装等日本男装常见上班搭配。",
                "比纯正式正装衬衫更适合做跨场景基础款。",
            ],
            source,
        )
    )
    image_rows.append(
        build_feature_card(
            "08_risk_note_card",
            "Sampling Note",
            "第六步保留的风控与验证建议",
            [
                "关键风险是面料是否偏薄、是否有过强光泽、版型是否过于商务化。",
                "正式上架前建议优先打样确认透感、领型和洗后平整度。",
                "若后续拿到更多原始详情图，可替换本次裁切细节图。",
            ],
            source,
        )
    )
    image_rows.append(
        build_spec_card(
            "09_specs_card",
            [
                ("中文品名", "免烫常规领白衬衫"),
                ("商品标题", "免烫DP成衣免烫男士白衬衫男抗皱易打理长袖衬衫春秋纯色全棉衬衣"),
                ("价格", "75.9 元"),
                ("销量信号", "全网 2.0万+ 件"),
                ("店铺", "义乌市顺顺服饰有限公司"),
                ("核心命中词", "纯棉 / DP / 易打理 / 抗皱 / 长袖"),
                ("日文搜索词", "ノーアイロン シャツ メンズ / 白シャツ メンズ / オフィスカジュアル シャツ メンズ"),
            ],
        )
    )

    PREVIEW_HTML.write_text(build_html(image_rows), encoding="utf-8")

    copy_rows = [
        {"字段": "日文标题", "内容": "ノーアイロン メンズ 白シャツ 長袖 レギュラーカラー 通勤向け ベーシックシャツ"},
        {"字段": "短卖点", "内容": "DP加工でシワになりにくく、毎日の通勤で扱いやすい白シャツ。"},
        {"字段": "短卖点", "内容": "レギュラーカラーで面接からオフィスカジュアルまで使いやすい。"},
        {"字段": "短卖点", "内容": "白シャツの基本需要に合わせた、清潔感重視の定番提案。"},
        {"字段": "商品说明", "内容": "中国側ソーシング情報ではDP・抗皺・長袖・純綿が主要訴求。上架前は透け感と光沢感のサンプル確認を推奨。"},
    ]

    gap_rows = [
        {"缺口": "原始详情长图", "处理方式": "1688 页面存在防爬限制，本次以主图原图裁切和本地营销卡片先完成可用包。"},
        {"缺口": "更多细节原图", "处理方式": "后续若能拿到袖口、后背、面料近拍，可直接替换 02-05 号图位。"},
        {"缺口": "真人上身图", "处理方式": "当前不强行生成，建议拿样后再补拍或用受控AI出图。"},
    ]

    payload = {
        "sku_id": "S01",
        "case_dir": str(CASE_DIR),
        "source_link": SOURCE_LINK,
        "backup_link": BACKUP_LINK,
        "raw_dir": str(RAW_DIR),
        "case_image_dir": str(CASE_IMAGE_DIR),
        "html_preview": str(PREVIEW_HTML),
        "selected_source_images": {
            "main_source": str(raw_image_path),
            "alphashop_main_thumb": str(RAW_DIR / "s01_main_thumb.png"),
            "alphashop_backup_thumb": str(RAW_DIR / "s01_backup_thumb.png"),
        },
        "case_images": image_rows,
        "copy_rows": copy_rows,
        "gap_rows": gap_rows,
    }
    JSON_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    md_lines = [
        "# S01 第六步详情素材包",
        "",
        f"- 输出目录：`{CASE_DIR}`",
        f"- 主商品链接：{SOURCE_LINK}",
        f"- 备选链接：{BACKUP_LINK}",
        "- 成品图数量：9",
        "- 生成策略：主图原图裁切 + 卖点卡 + 规格信息卡",
        "",
        "## 日文文案建议",
    ]
    for row in copy_rows:
        md_lines.append(f"- `{row['字段']}`：{row['内容']}")
    md_lines.extend(
        [
            "",
            "## 缺口与处理",
        ]
    )
    for row in gap_rows:
        md_lines.append(f"- `{row['缺口']}`：{row['处理方式']}")
    md_text = "\n".join(md_lines)
    MD_PATH.write_text(md_text, encoding="utf-8")
    (DOC_DIR / "summary.md").write_text(md_text, encoding="utf-8")
    (DOC_DIR / "payload.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    with pd.ExcelWriter(XLSX_PATH) as writer:
        pd.DataFrame(image_rows).to_excel(writer, sheet_name="case_images", index=False)
        pd.DataFrame(copy_rows).to_excel(writer, sheet_name="copy_rows", index=False)
        pd.DataFrame(gap_rows).to_excel(writer, sheet_name="gap_rows", index=False)
        pd.DataFrame(
            [
                {
                    "sku_id": "S01",
                    "product_name": "免烫常规领白衬衫",
                    "source_link": SOURCE_LINK,
                    "backup_link": BACKUP_LINK,
                    "source_image_url": SOURCE_IMAGE_URL,
                }
            ]
        ).to_excel(writer, sheet_name="overview", index=False)


if __name__ == "__main__":
    main()
