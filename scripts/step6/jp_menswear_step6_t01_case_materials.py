from __future__ import annotations

import json
import shutil
from pathlib import Path

import pandas as pd
from PIL import Image, ImageColor, ImageDraw, ImageFilter, ImageFont, ImageOps


ROOT = Path(__file__).resolve().parent
STEP6_DIR = ROOT / "step6_output"
EXPORT_DATE = "2026-03-20"

SOURCE_PACK_DIR = next(
    path
    for path in STEP6_DIR.iterdir()
    if path.is_dir() and path.name == "T01_白T素材包"
)
REFERENCE_DIR = STEP6_DIR / "oakln_reference"
CASE_DIR = STEP6_DIR / f"T01_案例版详情素材_{EXPORT_DATE}"
RAW_DIR = CASE_DIR / "01_同款精选原图"
CASE_IMAGE_DIR = CASE_DIR / "02_案例版成品图"
REFERENCE_OUT_DIR = CASE_DIR / "03_参考页拆解"
DOC_DIR = CASE_DIR / "04_文案与说明"
XLSX_PATH = STEP6_DIR / f"T01_案例版详情素材_{EXPORT_DATE}.xlsx"
MD_PATH = STEP6_DIR / f"T01_案例版详情素材说明_{EXPORT_DATE}.md"
JSON_PATH = STEP6_DIR / f"T01_案例版详情素材清单_{EXPORT_DATE}.json"
HTML_PREVIEW = CASE_DIR / "05_详情页预览.html"

for folder in [CASE_DIR, RAW_DIR, CASE_IMAGE_DIR, REFERENCE_OUT_DIR, DOC_DIR]:
    folder.mkdir(parents=True, exist_ok=True)


def find_file(folder: Path, pattern: str) -> Path:
    matches = sorted(folder.glob(pattern))
    if not matches:
        raise FileNotFoundError(f"Missing file: {folder} / {pattern}")
    return matches[0]


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


FONT_TITLE = load_font(48, bold=True)
FONT_SUBTITLE = load_font(24)
FONT_BODY = load_font(28)
FONT_SMALL = load_font(20)

BG = ImageColor.getrgb("#f6f4ef")
CARD = ImageColor.getrgb("#ffffff")
TEXT = ImageColor.getrgb("#222222")
MUTED = ImageColor.getrgb("#666666")
LINE = ImageColor.getrgb("#dddddd")
ACCENT = ImageColor.getrgb("#d5d0c5")

XIANG_DIR = next(path for path in SOURCE_PACK_DIR.rglob("*") if path.is_dir() and path.name == "象山阿姆达服饰有限公司")
XIANG_MAIN = XIANG_DIR / "主图细节图"

SELECTED_SOURCE_IMAGES = {
    "白T挂拍主图": find_file(XIANG_MAIN, "主图细节图_02_*.jpg"),
    "军绿色平铺": find_file(XIANG_MAIN, "主图细节图_03_*.jpg"),
    "灰色平铺": find_file(XIANG_MAIN, "主图细节图_05_*.jpg"),
    "领口近景": find_file(XIANG_MAIN, "主图细节图_07_*.jpg"),
    "黑色平铺": find_file(XIANG_MAIN, "主图细节图_08_*.jpg"),
    "排挂近景": find_file(XIANG_MAIN, "主图细节图_09_*.jpg"),
    "折叠领口": find_file(XIANG_MAIN, "主图细节图_10_*.jpg"),
    "白色平铺": find_file(XIANG_MAIN, "主图细节图_12_*.jpg"),
    "白色领口特写": find_file(XIANG_MAIN, "主图细节图_14_*.jpg"),
    "多色挂拍": find_file(XIANG_MAIN, "主图细节图_15_*.jpg"),
    "白色折叠特写": find_file(XIANG_MAIN, "主图细节图_18_*.jpg"),
    "米白平铺": find_file(XIANG_MAIN, "主图细节图_20_*.jpg"),
}

STEP6_XLSX = next(
    path
    for path in STEP6_DIR.glob("T01_*第6步素材包_*.xlsx")
    if not path.name.startswith("~$")
)
SUPPLIER_DF = pd.read_excel(STEP6_XLSX, sheet_name=0)
BASE_ROW = SUPPLIER_DF.loc[SUPPLIER_DF["供应商"].astype(str).str.contains("象山阿姆达")].iloc[0]


def sanitize_name(text: str) -> str:
    blocked = '<>:"/\\|?*'
    safe = "".join("_" if ch in blocked else ch for ch in text)
    return safe.replace(" ", "_")


def wrap_text(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, width: int) -> list[str]:
    words = list(text)
    lines: list[str] = []
    current = ""
    for word in words:
        attempt = current + word
        if draw.textlength(attempt, font=font) <= width:
            current = attempt
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines


def make_canvas() -> Image.Image:
    return Image.new("RGB", (1080, 1080), BG)


def contain_image(image: Image.Image, box: tuple[int, int]) -> Image.Image:
    return ImageOps.contain(image.convert("RGB"), box, method=Image.Resampling.LANCZOS)


def crop_remove_badge(image: Image.Image) -> Image.Image:
    width, height = image.size
    left = int(width * 0.24)
    return image.crop((left, 0, width, height))


def paste_card(canvas: Image.Image, image: Image.Image, xy: tuple[int, int]) -> None:
    shadow = Image.new("RGBA", (image.width + 24, image.height + 24), (0, 0, 0, 0))
    shadow_box = Image.new("RGBA", image.size, (0, 0, 0, 60))
    shadow.paste(shadow_box, (12, 12))
    shadow = shadow.filter(ImageFilter.GaussianBlur(14))
    canvas.paste(shadow, (xy[0] - 12, xy[1] - 12), shadow)

    card = Image.new("RGB", image.size, CARD)
    card.paste(image, (0, 0))
    canvas.paste(card, xy)


def save_image_card(name: str, image: Image.Image) -> Path:
    out_path = CASE_IMAGE_DIR / f"{name}.png"
    image.save(out_path, quality=95)
    return out_path


def copy_selected_originals() -> list[dict[str, str]]:
    copied_rows: list[dict[str, str]] = []
    for label, source in SELECTED_SOURCE_IMAGES.items():
        out_path = RAW_DIR / f"{sanitize_name(label)}{source.suffix.lower()}"
        shutil.copy2(source, out_path)
        copied_rows.append(
            {
                "素材名称": label,
                "来源供应商": "象山阿姆达服饰有限公司",
                "原始路径": str(source),
                "复制路径": str(out_path),
                "用途": "案例版详情素材基底",
            }
        )
    return copied_rows


def make_single_image_card(file_label: str, title: str, subtitle: str, source_key: str, crop_badge: bool = False) -> dict[str, str]:
    canvas = make_canvas()
    draw = ImageDraw.Draw(canvas)

    image = Image.open(SELECTED_SOURCE_IMAGES[source_key]).convert("RGB")
    if crop_badge:
        image = crop_remove_badge(image)
    art = contain_image(image, (860, 760))
    paste_card(canvas, art, ((1080 - art.width) // 2, 120))

    draw.text((80, 920), title, fill=TEXT, font=FONT_TITLE)
    draw.text((80, 980), subtitle, fill=MUTED, font=FONT_SUBTITLE)

    out_path = save_image_card(file_label, canvas)
    return {
        "成品文件": str(out_path),
        "类型": "纯图",
        "参考模块": title,
        "来源": source_key,
        "生成方式": "原图二次排版",
        "说明": subtitle,
    }


def make_feature_card(file_label: str, kicker: str, title: str, bullets: list[str], source_key: str) -> dict[str, str]:
    canvas = make_canvas()
    draw = ImageDraw.Draw(canvas)
    draw.rounded_rectangle((60, 70, 1020, 1010), radius=32, fill=CARD)

    image = Image.open(SELECTED_SOURCE_IMAGES[source_key]).convert("RGB")
    art = contain_image(image, (430, 740))
    paste_card(canvas, art, (90, 170))

    draw.text((580, 160), kicker, fill=MUTED, font=FONT_SMALL)
    draw.text((580, 200), title, fill=TEXT, font=FONT_TITLE)
    draw.line((580, 275, 940, 275), fill=ACCENT, width=2)

    y = 330
    for bullet in bullets:
        lines = wrap_text(draw, bullet, FONT_BODY, 360)
        draw.text((590, y), "•", fill=TEXT, font=FONT_BODY)
        inner_y = y
        for line in lines:
            draw.text((625, inner_y), line, fill=TEXT, font=FONT_BODY)
            inner_y += 42
        y = inner_y + 24

    out_path = save_image_card(file_label, canvas)
    return {
        "成品文件": str(out_path),
        "类型": "营销图",
        "参考模块": kicker,
        "来源": source_key,
        "生成方式": "本地营销图生成",
        "说明": title,
    }


def make_color_grid_card(file_label: str) -> dict[str, str]:
    canvas = make_canvas()
    draw = ImageDraw.Draw(canvas)
    draw.text((80, 70), "颜色一览", fill=TEXT, font=FONT_TITLE)
    draw.text((80, 130), "保留同款 300g 重磅纯棉版型的可用色卡。", fill=MUTED, font=FONT_SUBTITLE)

    grid_items = [
        ("白色", "白色平铺"),
        ("黑色", "黑色平铺"),
        ("灰色", "灰色平铺"),
        ("军绿", "军绿色平铺"),
        ("米白", "米白平铺"),
    ]

    x_positions = [80, 420, 760]
    y_positions = [230, 620]
    for index, (label, key) in enumerate(grid_items):
        image = Image.open(SELECTED_SOURCE_IMAGES[key]).convert("RGB")
        art = contain_image(image, (260, 260))
        x = x_positions[index % 3]
        y = y_positions[index // 3]
        draw.rounded_rectangle((x - 10, y - 10, x + 290, y + 330), radius=24, fill=CARD, outline=LINE)
        canvas.paste(art, (x + (270 - art.width) // 2, y + 10))
        draw.text((x, y + 282), label, fill=TEXT, font=FONT_BODY)
    out_path = save_image_card(file_label, canvas)
    return {
        "成品文件": str(out_path),
        "类型": "纯图",
        "参考模块": "多色展示",
        "来源": "白色平铺/黑色平铺/灰色平铺/军绿色平铺/米白平铺",
        "生成方式": "本地色卡拼图",
        "说明": "适合详情页后段颜色展示。",
    }


def make_spec_card(file_label: str) -> dict[str, str]:
    canvas = make_canvas()
    draw = ImageDraw.Draw(canvas)
    draw.rounded_rectangle((60, 60, 1020, 1020), radius=32, fill=CARD)

    draw.text((90, 90), "规格信息", fill=TEXT, font=FONT_TITLE)
    draw.text((90, 150), "按同款主素材供应商信息整理，可直接转成商品详情页文案。", fill=MUTED, font=FONT_SUBTITLE)
    draw.line((90, 205, 990, 205), fill=LINE, width=2)

    specs = [
        ("品类", "男士重磅圆领短袖 T 恤"),
        ("面料", f"{BASE_ROW['面料名称']} / 主面料成分：{BASE_ROW['主面料成分']}"),
        ("克重", str(BASE_ROW["克重"])),
        ("厚薄", str(BASE_ROW["厚薄"])),
        ("版型", str(BASE_ROW["版型"])),
        ("领型", str(BASE_ROW["领型"])),
        ("颜色", str(BASE_ROW["颜色"])),
        ("尺码", str(BASE_ROW["尺码"])),
        ("建议卖点", "300g 重磅、纯棉、圆领、宽松、加厚、降低透感"),
        ("4 月定位", "单穿主力、内搭补充、适合基础都市休闲男装"),
    ]

    y = 250
    for label, value in specs:
        draw.text((100, y), label, fill=MUTED, font=FONT_BODY)
        lines = wrap_text(draw, value, FONT_BODY, 660)
        inner_y = y
        for line in lines:
            draw.text((280, inner_y), line, fill=TEXT, font=FONT_BODY)
            inner_y += 40
        y = inner_y + 24

    out_path = save_image_card(file_label, canvas)
    return {
        "成品文件": str(out_path),
        "类型": "信息图",
        "参考模块": "规格说明",
        "来源": "第 6 步供应商评分",
        "生成方式": "本地信息图生成",
        "说明": "可拆成详情页规格、面料、尺码、颜色说明。",
    }


def make_three_panel_module(file_label: str) -> dict[str, str]:
    canvas = Image.new("RGB", (1080, 1280), BG)
    draw = ImageDraw.Draw(canvas)
    draw.text((70, 56), "三联卖点模块", fill=TEXT, font=FONT_TITLE)
    draw.text((70, 114), "补齐参考页中段的 3 列卖点展示结构。", fill=MUTED, font=FONT_SUBTITLE)

    panels = [
        {
            "image": "白色平铺",
            "title": "轮廓干净，基础好搭",
            "body": "用白色平铺图承接参考页左侧的人物展示逻辑，强调这件白T本身的版型与基础搭配价值。",
            "tags": ["基础主力", "4月单穿"],
        },
        {
            "image": "白色领口特写",
            "title": "圆领罗纹，细节更稳",
            "body": "用近景特写支撑领口、走线和罗纹细节，代替参考页中段的局部工艺说明图。",
            "tags": ["圆领", "细节清晰"],
        },
        {
            "image": "白色折叠特写",
            "title": "重磅棉感，降低透感",
            "body": "用折叠与面料近景表现厚实感，承接参考页右侧的面料卖点模块。",
            "tags": ["300g重磅", "降低透感"],
        },
    ]

    lefts = [70, 380, 690]
    top = 180
    panel_w = 320
    panel_h = 410

    for index, panel in enumerate(panels):
        x = lefts[index]
        draw.rounded_rectangle((x, top, x + panel_w, top + panel_h), radius=26, fill=CARD)
        src = Image.open(SELECTED_SOURCE_IMAGES[panel["image"]]).convert("RGB")
        art = contain_image(src, (panel_w - 24, 300))
        canvas.paste(art, (x + (panel_w - art.width) // 2, top + 12))

        tag_x = x + 20
        for tag in panel["tags"]:
            tag_w = int(draw.textlength(tag, font=FONT_SMALL) + 28)
            tag_box = (tag_x, top + 322, tag_x + tag_w, top + 356)
            draw.rounded_rectangle(tag_box, radius=18, fill=(255, 255, 255))
            draw.text((tag_x + 14, top + 329), tag, fill=TEXT, font=FONT_SMALL)
            tag_x += tag_w + 10

        title_lines = wrap_text(draw, panel["title"], FONT_BODY, panel_w - 8)
        body_lines = wrap_text(draw, panel["body"], FONT_SMALL, panel_w - 8)
        y = top + panel_h + 24
        for line in title_lines:
            draw.text((x, y), line, fill=TEXT, font=FONT_BODY)
            y += 38
        y += 8
        for line in body_lines[:5]:
            draw.text((x, y), line, fill=MUTED, font=FONT_SMALL)
            y += 30

    out_path = save_image_card(file_label, canvas)
    return {
        "成品文件": str(out_path),
        "类型": "营销图",
        "参考模块": "三联卖点模块",
        "来源": "白色平铺/白色领口特写/白色折叠特写",
        "生成方式": "本地结构化营销图生成",
        "说明": "补齐参考页中段三列卖点结构。",
    }


def build_html_preview(image_rows: list[dict[str, str]]) -> None:
    image_tags = "\n".join(
        f'<section class="card"><img src="02_案例版成品图/{Path(row["成品文件"]).name}" alt="{row["参考模块"]}"><p>{row["参考模块"]}</p></section>'
        for row in image_rows
    )
    html = f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <title>T01 白 T 案例版详情页预览</title>
  <style>
    body {{ margin: 0; background: #f6f4ef; font-family: "Microsoft YaHei", sans-serif; color: #222; }}
    .wrap {{ max-width: 980px; margin: 0 auto; padding: 40px 24px 80px; }}
    h1 {{ font-size: 32px; margin: 0 0 8px; }}
    p.lead {{ color: #666; margin: 0 0 28px; }}
    .card {{ background: #fff; margin: 0 0 28px; border-radius: 24px; overflow: hidden; box-shadow: 0 12px 30px rgba(0,0,0,0.06); }}
    .card img {{ width: 100%; display: block; }}
    .card p {{ margin: 0; padding: 14px 18px 18px; font-size: 14px; color: #666; }}
  </style>
</head>
<body>
  <main class="wrap">
    <h1>T01 厚实不透白 T</h1>
    <p class="lead">按 Oakln 参考页结构整理的案例版详情素材预览。主视觉统一使用象山阿姆达同款图，避免混款。</p>
    {image_tags}
  </main>
</body>
</html>
"""
    HTML_PREVIEW.write_text(html, encoding="utf-8")


def build_workbook(
    original_rows: list[dict[str, str]],
    image_rows: list[dict[str, str]],
    reference_rows: list[dict[str, str]],
    copy_rows: list[dict[str, str]],
    prompt_rows: list[dict[str, str]],
    gap_rows: list[dict[str, str]],
) -> None:
    with pd.ExcelWriter(XLSX_PATH, engine="openpyxl") as writer:
        pd.DataFrame(reference_rows).to_excel(writer, sheet_name="案例参考结构", index=False)
        pd.DataFrame(original_rows).to_excel(writer, sheet_name="同款精选原图", index=False)
        pd.DataFrame(image_rows).to_excel(writer, sheet_name="案例版成品图", index=False)
        pd.DataFrame(copy_rows).to_excel(writer, sheet_name="日文文案", index=False)
        pd.DataFrame(prompt_rows).to_excel(writer, sheet_name="AlphaShop补图提示词", index=False)
        pd.DataFrame(gap_rows).to_excel(writer, sheet_name="缺口与补法", index=False)


def build_markdown(image_rows: list[dict[str, str]], gap_rows: list[dict[str, str]]) -> None:
    lines = [
        "# T01 白 T 案例版详情素材",
        "",
        f"- 输出目录：`{CASE_DIR}`",
        f"- 成品图数量：`{len(image_rows)}`",
        "- 主视觉来源：`象山阿姆达服饰有限公司`",
        "- 处理原则：只用已确认同款做最终对外素材；其它供应商仅保留作内部备份。",
        "",
        "## 成品图",
    ]
    for row in image_rows:
        lines.append(f"- `{Path(row['成品文件']).name}`：{row['说明']}")
    lines.append("")
    lines.append("## 缺口与补法")
    for row in gap_rows:
        lines.append(f"- {row['缺口']}：{row['处理方式']}")
    MD_PATH.write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    if REFERENCE_DIR.exists():
        for file in REFERENCE_DIR.iterdir():
            if file.is_file():
                shutil.copy2(file, REFERENCE_OUT_DIR / file.name)

    original_rows = copy_selected_originals()

    image_rows = [
        make_single_image_card("01_白T正面单图", "首屏主商品图", "保留干净白底与同款版型，适合首屏和 SKU 图。", "白色平铺"),
        make_single_image_card("02_挂拍陈列图", "挂拍与陈列", "模拟参考页第二张商品陈列图，强化基础款稳定感。", "白T挂拍主图", crop_badge=True),
        make_single_image_card("03_领口细节图", "领口与罗纹", "突出圆领走线与领口厚度，适合细节段落。", "白色领口特写"),
        make_single_image_card("04_折叠面料图", "面料与触感", "补足参考页中的折叠堆叠和面料质感模块。", "折叠领口"),
        make_single_image_card("05_排挂近景图", "排挂近景", "用同款排挂图表现基础色系与店铺陈列感。", "排挂近景"),
        make_color_grid_card("06_色卡一览"),
        make_single_image_card("06B_多色挂拍图", "多色挂拍", "模拟参考页后段的多色陈列模块。", "多色挂拍", crop_badge=True),
        make_feature_card(
            "07_卖点_300g纯棉",
            "核心卖点",
            "300g 重磅纯棉",
            ["300g 重磅面料，视觉更挺。", "纯棉基底，适合四月单穿。", "主打基础都市休闲场景。"],
            "白色平铺",
        ),
        make_feature_card(
            "08_卖点_加厚降低透感",
            "穿着感知",
            "加厚，降低透感",
            ["相对普通薄款更安心。", "白色单穿更容易撑住轮廓。", "可作为白 T 主款卖点表达。"],
            "白T挂拍主图",
        ),
        make_feature_card(
            "09_卖点_圆领宽松版型",
            "版型卖点",
            "圆领，宽松版型",
            ["圆领结构基础，搭配门槛低。", "宽松版型更贴近当前日本男装方向。", "适合四月做单穿和叠穿。"],
            "白色领口特写",
        ),
        make_three_panel_module("09B_三联卖点模块"),
        make_spec_card("10_规格信息图"),
    ]

    build_html_preview(image_rows)

    reference_rows = [
        {"参考模块": "首屏主商品图", "Oakln参考特征": "单品正面大图，背景干净。", "我们对应素材": "01_白T正面单图.png", "状态": "已完成"},
        {"参考模块": "挂拍/陈列", "Oakln参考特征": "第二屏延续商品展示，强调质感。", "我们对应素材": "02_挂拍陈列图.png", "状态": "已完成"},
        {"参考模块": "领口与织法细节", "Oakln参考特征": "近距离细节图，用于支撑卖点。", "我们对应素材": "03_领口细节图.png", "状态": "已完成"},
        {"参考模块": "面料触感", "Oakln参考特征": "折叠或堆叠图，传递厚实感。", "我们对应素材": "04_折叠面料图.png", "状态": "已完成"},
        {"参考模块": "多色展示", "Oakln参考特征": "后段给出多色排列与平铺。", "我们对应素材": "05_排挂近景图.png / 06_色卡一览.png", "状态": "已完成"},
        {"参考模块": "三联卖点模块", "Oakln参考特征": "三列图片加说明文字，承接详情中段核心卖点。", "我们对应素材": "09B_三联卖点模块.png", "状态": "已完成"},
        {"参考模块": "卖点说明", "Oakln参考特征": "详情文字说明，不依赖图片堆字。", "我们对应素材": "07-10 + 日文文案", "状态": "已完成"},
        {"参考模块": "规格说明", "Oakln参考特征": "页面下方补充材质、尺寸、产地。", "我们对应素材": "10_规格信息图.png + 日文文案", "状态": "已完成"},
    ]

    jp_title = "300g重磅コットン クルーネックTシャツ メンズ 半袖 ゆったり 厚手 透けにくい ベーシック 4月向け"
    jp_description = (
        "肉感のある300gのコットン生地を使用した、クルーネックの半袖Tシャツです。"
        "ゆったりとしたベーシックなシルエットで、一枚着でもインナーでも使いやすく、"
        "春の立ち上がりから初夏まで長く活躍します。"
    )
    jp_feature_lines = [
        "厚手のコットン素材で、白でも透け感を抑えやすい。",
        "クルーネック仕様で首回りがすっきり見える。",
        "ゆったりシルエットで今のメンズカジュアルに合わせやすい。",
        "4月は単品主役、シャツやライトアウターのインナーにも対応。",
    ]
    copy_rows = [
        {"字段": "日文标题", "内容": jp_title},
        {"字段": "短卖点1", "内容": "300gの重厚感ある生地で、シルエットがきれいに見える。"},
        {"字段": "短卖点2", "内容": "白でも透け感を抑えやすく、春の一枚着に向く。"},
        {"字段": "短卖点3", "内容": "ベーシックなクルーネックで着回しやすい。"},
        {"字段": "商品说明", "内容": jp_description},
        {"字段": "详情卖点段落", "内容": " / ".join(jp_feature_lines)},
        {"字段": "材质", "内容": "素材：コットン"},
        {"字段": "克重", "内容": "生地感：300gの厚手タイプ"},
        {"字段": "尺码", "内容": "サイズ：XS / S / M / L / XL / XXL / XXXL"},
        {"字段": "颜色", "内容": "カラー：ホワイト / ブラック / グレー / オリーブ / アイボリー / ベージュ"},
        {"字段": "产地", "内容": "原産地：中国"},
        {"字段": "季节建议", "内容": "春から初夏向け。4月の主力カットソーとして提案しやすい。"},
    ]

    prompt_rows = [
        {
            "用途": "模特上身",
            "输入底图": str(SELECTED_SOURCE_IMAGES["白色平铺"]),
            "AlphaShop提示词": "请保持上传白色短袖T恤的同款领口、肩线和长度不变，生成25-32岁东亚男性模特上身图，简洁都市风，站姿自然，浅灰背景，突出300g重磅纯棉白T的厚实感，不要加入图案、印花、配饰遮挡。",
        },
        {
            "用途": "营销图生成",
            "输入底图": str(SELECTED_SOURCE_IMAGES["白T挂拍主图"]),
            "AlphaShop提示词": "基于上传的同款白T挂拍图，生成极简电商营销图，保留同款衣型与圆领，不改变款式，不新增口袋和印花，背景为米白色和浅灰色，画面强调重磅、加厚、基础款。",
        },
        {
            "用途": "背景图生成",
            "输入底图": str(SELECTED_SOURCE_IMAGES["排挂近景"]),
            "AlphaShop提示词": "请把同款白T排挂图扩展成更干净的电商背景图，保留白T和木质衣架，去掉杂乱元素，适合商品详情页中段使用。",
        },
    ]

    gap_rows = [
        {"缺口": "安全同款视频", "处理方式": "当前不放入最终对外素材。若要补视频，建议先用白色平铺图走 AlphaShop 的模特上身或营销图，再转视频。"},
        {"缺口": "真人男模上身图", "处理方式": "现有同款原图以平铺和挂拍为主，已附 AlphaShop 提示词用于后补生成。"},
        {"缺口": "尺码表原始图", "处理方式": "当前改用规格信息图和日文文案承接；若供应商后续补尺码图，再替换。"},
    ]

    build_workbook(original_rows, image_rows, reference_rows, copy_rows, prompt_rows, gap_rows)
    build_markdown(image_rows, gap_rows)

    payload = {
        "case_dir": str(CASE_DIR),
        "reference_dir": str(REFERENCE_OUT_DIR),
        "raw_dir": str(RAW_DIR),
        "case_image_dir": str(CASE_IMAGE_DIR),
        "html_preview": str(HTML_PREVIEW),
        "xlsx": str(XLSX_PATH),
        "markdown": str(MD_PATH),
        "base_supplier": "象山阿姆达服饰有限公司",
        "selected_source_images": {key: str(value) for key, value in SELECTED_SOURCE_IMAGES.items()},
        "case_images": image_rows,
        "copy_rows": copy_rows,
        "gap_rows": gap_rows,
    }
    JSON_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
