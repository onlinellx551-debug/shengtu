from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
STEP6 = ROOT / "step6_output" / "material_bundle"
ROUND10 = STEP6 / "T01_round10_2026-03-24"
ROUND12 = STEP6 / "T01_round12_2026-03-25"
ROUND13 = STEP6 / "T01_round13_2026-03-25"
OUT_DIR = ROUND13 / "01_input_boards_v2"


def fit_cover(img: Image.Image, size: tuple[int, int]) -> Image.Image:
    tw, th = size
    ratio = max(tw / img.width, th / img.height)
    resized = img.resize((int(img.width * ratio), int(img.height * ratio)), Image.LANCZOS)
    left = max(0, (resized.width - tw) // 2)
    top = max(0, (resized.height - th) // 2)
    return resized.crop((left, top, left + tw, top + th))


def fit_contain(img: Image.Image, size: tuple[int, int], bg: str = "#f7f5ef") -> Image.Image:
    tw, th = size
    ratio = min(tw / img.width, th / img.height)
    resized = img.resize((max(1, int(img.width * ratio)), max(1, int(img.height * ratio))), Image.LANCZOS)
    canvas = Image.new("RGB", size, bg)
    left = (tw - resized.width) // 2
    top = (th - resized.height) // 2
    canvas.paste(resized, (left, top))
    return canvas


def load(path: Path) -> Image.Image:
    return Image.open(path).convert("RGB")


def font(size: int):
    candidates = [
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/segoeui.ttf",
        "C:/Windows/Fonts/msyh.ttc",
    ]
    for candidate in candidates:
        p = Path(candidate)
        if p.exists():
            return ImageFont.truetype(str(p), size=size)
    return ImageFont.load_default()


def draw_card(canvas: Image.Image, box: tuple[int, int, int, int], title: str, img: Image.Image) -> None:
    draw = ImageDraw.Draw(canvas)
    x1, y1, x2, y2 = box
    draw.rounded_rectangle(box, radius=18, fill="white", outline="#ddd6ca", width=2)
    inner = (x1 + 14, y1 + 16, x2 - 14, y2 - 42)
    fitted = fit_contain(img, (inner[2] - inner[0], inner[3] - inner[1]))
    canvas.paste(fitted, (inner[0], inner[1]))
    draw.text((x1 + 18, y2 - 30), title, fill="#66605a", font=font(18))


def build(board_name: str, ref_name: str, footer: str) -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    base = Image.new("RGB", (1600, 900), "#f4f1ea")
    draw = ImageDraw.Draw(base)
    draw.rounded_rectangle((26, 26, 524, 874), radius=24, fill="#fbfaf7", outline="#dbd4c8", width=2)
    draw.rounded_rectangle((554, 26, 1574, 874), radius=24, fill="#fbfaf7", outline="#dbd4c8", width=2)
    draw.text((52, 48), "OUR ACTUAL T-SHIRT", fill="#44403a", font=font(36))
    draw.text((582, 48), "EXACT REFERENCE COMPOSITION", fill="#44403a", font=font(36))

    assets = ROUND12 / "04_web_preview" / "assets"
    cards = [
        ("WHITE", load(assets / "tail_03_white_flat.jpg")),
        ("BLACK", load(assets / "tail_02_black_flat.jpg")),
        ("GRAY", load(assets / "tail_04_gray_flat.jpg")),
        ("IVORY", load(assets / "tail_05_ivory_flat.jpg")),
        ("OLIVE", load(assets / "tail_06_olive_flat.jpg")),
        ("LABEL", load(assets / "top_02_collar.jpg")),
    ]

    positions = [
        (52, 108, 186, 258),
        (208, 108, 342, 258),
        (52, 280, 186, 430),
        (208, 280, 342, 430),
        (52, 452, 186, 602),
        (330, 172, 492, 392),
    ]
    for (title, img), pos in zip(cards, positions):
        draw_card(base, pos, title, img)

    hanging = fit_cover(load(assets / "top_04_hanging.jpg"), (440, 220))
    base.paste(hanging, (52, 620))
    draw.rounded_rectangle((52, 620, 492, 840), radius=18, outline="#ddd6ca", width=2)

    draw.text((52, 846), footer, fill="#7a746d", font=font(18))

    ref = fit_cover(load(ROUND10 / "01_input_boards" / ref_name), (970, 760))
    # Crop only the right-side exact reference panel from the original board.
    ref = ref.crop((360, 0, ref.width, ref.height))
    ref = fit_cover(ref, (920, 760))
    base.paste(ref, (580, 96))

    out = OUT_DIR / board_name
    base.save(out, quality=96)


def main() -> None:
    build(
        "G05_stack_multicolor_board_v2.jpg",
        "G05_stack_multicolor_board_v1.jpg",
        "Use these exact short-sleeve tee colors and our label. Match the right-side stack composition.",
    )
    build(
        "G07_fold_multicolor_board_v2.jpg",
        "G07_fold_multicolor_board_v1.jpg",
        "Use these exact short-sleeve tee colors and our label. Match the right-side close fold composition.",
    )


if __name__ == "__main__":
    main()
