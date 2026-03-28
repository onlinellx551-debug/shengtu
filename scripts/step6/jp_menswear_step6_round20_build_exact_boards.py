from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
ROUND20 = ROOT / "step6_output" / "material_bundle" / "T01_round20_2026-03-26"
REF_DIR = ROUND20 / "01_reference_exact"
OUT_DIR = ROUND20 / "01_input_boards_exact"


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


def load(path: Path) -> Image.Image:
    return Image.open(path).convert("RGB")


def font(size: int, bold: bool = False):
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
        p = Path(candidate)
        if p.exists():
            return ImageFont.truetype(str(p), size=size)
    return ImageFont.load_default()


def draw_card(canvas: Image.Image, box: tuple[int, int, int, int], title: str, img: Image.Image) -> None:
    draw = ImageDraw.Draw(canvas)
    x1, y1, x2, y2 = box
    draw.rounded_rectangle(box, radius=16, fill="white", outline="#ddd6ca", width=2)
    inner = (x1 + 12, y1 + 14, x2 - 12, y2 - 40)
    fitted = fit_contain(img, (inner[2] - inner[0], inner[3] - inner[1]))
    canvas.paste(fitted, (inner[0], inner[1]))
    draw.text((x1 + 14, y2 - 28), title, fill="#66605a", font=font(18))


def build(ref_name: str, board_name: str, footer: str) -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    base = Image.new("RGB", (1600, 900), "#f4f1ea")
    draw = ImageDraw.Draw(base)

    draw.rounded_rectangle((26, 26, 524, 874), radius=24, fill="#fbfaf7", outline="#dbd4c8", width=2)
    draw.rounded_rectangle((554, 26, 1574, 874), radius=24, fill="#fbfaf7", outline="#dbd4c8", width=2)
    draw.text((52, 48), "OUR REAL PRODUCT", fill="#44403a", font=font(34, bold=True))
    draw.text((582, 48), "EXACT REFERENCE SHOT", fill="#44403a", font=font(34, bold=True))

    assets_dir = ROUND20 / "03_assets"
    generated_dir = ROUND20 / "02_generated"
    real_dir = ROUND20 / "01_real_candidates"
    cards = [
        ("FRONT", load(real_dir / "top01_real_front_whitespace.jpg")),
        ("ON BODY A", load(generated_dir / "top02_onbody_round20_clean.jpg")),
        ("ON BODY B", load(generated_dir / "top03_onbody_round20_clean.jpg")),
        ("COLLAR", load(assets_dir / "top_08_collar.jpg")),
        ("STACK", load(assets_dir / "top_05_stack_round14.jpg")),
        ("FOLD", load(assets_dir / "top_06_fold.jpg")),
    ]
    positions = [
        (52, 110, 186, 260),
        (208, 110, 342, 260),
        (364, 110, 498, 260),
        (52, 282, 186, 432),
        (208, 282, 342, 432),
        (364, 282, 498, 432),
    ]
    for (title, img), pos in zip(cards, positions):
        draw_card(base, pos, title, img)

    info_box = (58, 472, 492, 786)
    draw.rounded_rectangle(info_box, radius=16, fill="white", outline="#ddd6ca", width=2)
    draw.text((82, 500), "MANDATORY", fill="#44403a", font=font(24, bold=True))
    lines = [
        "adult East Asian male only",
        "our real heavyweight crew neck tee",
        "short sleeve, thick cotton jersey",
        "no knitwear, no sweatshirt, no long sleeve",
        "same crop / pose / light / mood as reference",
        "no text, no collage, no extra objects",
    ]
    y = 548
    for line in lines:
        draw.text((82, y), line, fill="#6a635b", font=font(19))
        y += 36
    draw.text((82, 760), footer, fill="#7a746d", font=font(18))

    ref = fit_cover(load(REF_DIR / ref_name), (920, 760))
    base.paste(ref, (580, 96))
    base.save(OUT_DIR / board_name, quality=96)


def main() -> None:
    build("ref_top02.webp", "top02_exact_board_round20.jpg", "Match this exact top-right on-body shot.")
    build("ref_top03.webp", "top03_exact_board_round20.jpg", "Match this exact lower-left on-body shot.")
    build("ref_top04.webp", "top04_exact_board_round20.jpg", "Match this exact lower-right on-body shot.")
    build("ref_top06.webp", "top06_exact_board_round20.jpg", "Match this exact folded structure close-up.")
    build("ref_top07.webp", "top07_exact_board_round20.jpg", "Match this exact multicolor folded still-life.")
    build("ref_top08.webp", "top08_exact_board_round20.jpg", "Match this exact collar structure close-up.")


if __name__ == "__main__":
    main()
