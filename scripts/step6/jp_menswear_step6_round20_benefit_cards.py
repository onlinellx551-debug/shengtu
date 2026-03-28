from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont


ROOT = Path(__file__).resolve().parent
ROUND20 = ROOT / "step6_output" / "material_bundle" / "T01_round20_2026-03-26"
ASSETS = ROUND20 / "03_assets"


def fit_cover(img: Image.Image, size: tuple[int, int]) -> Image.Image:
    tw, th = size
    ratio = max(tw / img.width, th / img.height)
    resized = img.resize((max(1, int(img.width * ratio)), max(1, int(img.height * ratio))), Image.LANCZOS)
    left = max(0, (resized.width - tw) // 2)
    top = max(0, (resized.height - th) // 2)
    return resized.crop((left, top, left + tw, top + th))


def get_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = []
    if bold:
        candidates.extend(
            [
                "C:/Windows/Fonts/msyhbd.ttc",
                "C:/Windows/Fonts/simhei.ttf",
                "C:/Windows/Fonts/arialbd.ttf",
            ]
        )
    candidates.extend(
        [
            "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/arial.ttf",
            "C:/Windows/Fonts/segoeui.ttf",
        ]
    )
    for candidate in candidates:
        path = Path(candidate)
        if path.exists():
            return ImageFont.truetype(str(path), size=size)
    return ImageFont.load_default()


def add_bottom_gradient(base: Image.Image, height: int = 210) -> Image.Image:
    overlay = Image.new("RGBA", base.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    w, h = base.size
    for i in range(height):
        alpha = int(150 * (i / max(1, height - 1)))
        y = h - height + i
        draw.rectangle((0, y, w, y + 1), fill=(14, 18, 24, alpha))
    return Image.alpha_composite(base.convert("RGBA"), overlay)


def chip(draw: ImageDraw.ImageDraw, xy: tuple[int, int], text: str) -> None:
    x, y = xy
    font = get_font(26, bold=True)
    bbox = draw.textbbox((0, 0), text, font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    pad_x = 22
    pad_y = 16
    box = (x, y, x + tw + pad_x * 2, y + th + pad_y * 2)
    draw.rounded_rectangle(box, radius=18, fill=(255, 255, 255, 220), outline=(255, 255, 255, 235), width=2)
    draw.text((x + pad_x, y + pad_y - 2), text, fill=(35, 43, 56, 230), font=font)


def build_card(src: Path, dest: Path, title: str, chips: list[str], crop_box: tuple[int, int, int, int] | None = None) -> None:
    base = Image.open(src).convert("RGB")
    if crop_box:
        base = base.crop(crop_box)
    base = fit_cover(base, (1200, 1200))
    base = add_bottom_gradient(base)
    base = base.filter(ImageFilter.UnsharpMask(radius=1.2, percent=120, threshold=2))
    draw = ImageDraw.Draw(base, "RGBA")

    title_font = get_font(52, bold=True)
    subtitle_font = get_font(26, bold=False)
    draw.text((72, 820), title, fill=(255, 255, 255, 244), font=title_font)
    draw.text((72, 888), "围绕真实需求整理的核心卖点", fill=(255, 255, 255, 210), font=subtitle_font)

    chip_y = 1002
    chip_x = 72
    for label in chips:
        chip(draw, (chip_x, chip_y), label)
        bbox = draw.textbbox((0, 0), label, font=get_font(26, bold=True))
        chip_x += (bbox[2] - bbox[0]) + 92

    base.convert("RGB").save(dest, quality=96)


def main() -> None:
    build_card(
        ASSETS / "top_08_collar.jpg",
        ASSETS / "benefit_02_clean.jpg",
        "领口稳定",
        ["不易松垮", "走线干净"],
        crop_box=(90, 40, 690, 730),
    )
    build_card(
        ASSETS / "top_06_fold.jpg",
        ASSETS / "benefit_03_clean.jpg",
        "300g厚实棉感",
        ["单穿不透", "可机洗"],
        crop_box=(40, 20, 730, 710),
    )


if __name__ == "__main__":
    main()
