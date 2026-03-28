from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
ROUND20 = ROOT / "step6_output" / "material_bundle" / "T01_round20_2026-03-26"
REF = ROUND20 / "01_reference_exact" / "ref_top07.webp"
BEST = ROUND20 / "03_assets" / "top_07_fold_round14.jpg"
FRONT = ROUND20 / "01_real_candidates" / "top01_real_front_whitespace.jpg"
COLLAR = ROUND20 / "03_assets" / "top_08_collar.jpg"
OUT = ROUND20 / "01_input_boards_exact_v2" / "top07_exact_board_v2.jpg"


def fit_cover(img: Image.Image, size: tuple[int, int]) -> Image.Image:
    tw, th = size
    ratio = max(tw / img.width, th / img.height)
    resized = img.resize((max(1, int(img.width * ratio)), max(1, int(img.height * ratio))), Image.LANCZOS)
    left = max(0, (resized.width - tw) // 2)
    top = max(0, (resized.height - th) // 2)
    return resized.crop((left, top, left + tw, top + th))


def fit_contain(img: Image.Image, size: tuple[int, int], bg: str = "#f7f5ef") -> Image.Image:
    tw, th = size
    ratio = min(tw / img.width, th / img.height)
    resized = img.resize((max(1, int(img.width * ratio)), max(1, int(img.height * ratio))), Image.LANCZOS)
    canvas = Image.new("RGB", size, bg)
    canvas.paste(resized, ((tw - resized.width) // 2, (th - resized.height) // 2))
    return canvas


def font(size: int, bold: bool = False):
    path = "C:/Windows/Fonts/msyhbd.ttc" if bold else "C:/Windows/Fonts/msyh.ttc"
    return ImageFont.truetype(path, size=size)


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    canvas = Image.new("RGB", (1600, 920), "#f4f1ea")
    draw = ImageDraw.Draw(canvas)
    draw.rounded_rectangle((26, 26, 520, 894), radius=24, fill="#fbfaf7", outline="#dbd4c8", width=2)
    draw.rounded_rectangle((548, 26, 1574, 894), radius=24, fill="#fbfaf7", outline="#dbd4c8", width=2)
    draw.text((52, 48), "OUR TEE / KEEP THIS PRODUCT", fill="#44403a", font=font(32, True))
    draw.text((576, 48), "EXACT TARGET / MATCH THIS SHOT", fill="#44403a", font=font(32, True))

    cards = [
        ("CURRENT BEST", Image.open(BEST).convert("RGB")),
        ("FRONT FLAT", Image.open(FRONT).convert("RGB")),
        ("COLLAR", Image.open(COLLAR).convert("RGB")),
    ]
    positions = [
        (52, 110, 494, 330),
        (52, 354, 270, 544),
        (276, 354, 494, 544),
    ]
    for (title, img), (x1, y1, x2, y2) in zip(cards, positions):
        draw.rounded_rectangle((x1, y1, x2, y2), radius=16, fill="white", outline="#ddd6ca", width=2)
        fitted = fit_contain(img, (x2 - x1 - 24, y2 - y1 - 54))
        canvas.paste(fitted, (x1 + 12, y1 + 12))
        draw.text((x1 + 14, y2 - 28), title, fill="#66605a", font=font(18))

    draw.rounded_rectangle((52, 584, 494, 860), radius=16, fill="white", outline="#ddd6ca", width=2)
    draw.text((72, 612), "TOP07 REQUIREMENTS", fill="#44403a", font=font(24, True))
    notes = [
        "same folded multicolor stack as target",
        "same warm tabletop and quiet premium mood",
        "must read as folded short sleeve cotton T-shirts",
        "visible labels should look natural and minimal",
        "colors only from our real range: white / ivory / gray / taupe",
        "keep the stack tidy, not messy",
        "no cat, no extra props, no text",
    ]
    y = 650
    for note in notes:
        draw.text((72, y), note, fill="#6a635b", font=font(18))
        y += 28

    ref = fit_cover(Image.open(REF).convert("RGB"), (960, 790))
    canvas.paste(ref, (578, 90))
    canvas.save(OUT, quality=96)
    print(OUT)


if __name__ == "__main__":
    main()
