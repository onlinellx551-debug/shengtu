from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
ROUND20 = ROOT / "step6_output" / "material_bundle" / "T01_round20_2026-03-26"
REF_DIR = ROUND20 / "01_reference_exact"
OUT_DIR = ROUND20 / "01_input_boards_exact_v2"


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


def build(task: str, ref_name: str, winner_name: str, footer: str) -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    generated = ROUND20 / "02_generated"
    real = ROUND20 / "01_real_candidates"
    assets = ROUND20 / "03_assets"

    canvas = Image.new("RGB", (1600, 920), "#f4f1ea")
    draw = ImageDraw.Draw(canvas)
    draw.rounded_rectangle((26, 26, 520, 894), radius=24, fill="#fbfaf7", outline="#dbd4c8", width=2)
    draw.rounded_rectangle((548, 26, 1574, 894), radius=24, fill="#fbfaf7", outline="#dbd4c8", width=2)
    draw.text((52, 48), "OUR TEE / KEEP THIS PRODUCT", fill="#44403a", font=font(32, bold=True))
    draw.text((576, 48), "EXACT TARGET / MATCH THIS SHOT", fill="#44403a", font=font(32, bold=True))

    cards = [
        ("FRONT FLAT", load(real / "top01_real_front_whitespace.jpg")),
        ("CURRENT BEST", load(generated / winner_name)),
        ("COLLAR", load(assets / "top_08_collar.jpg")),
    ]
    positions = [
        (52, 110, 494, 300),
        (52, 324, 494, 514),
        (52, 538, 494, 728),
    ]
    for (title, img), (x1, y1, x2, y2) in zip(cards, positions):
        draw.rounded_rectangle((x1, y1, x2, y2), radius=16, fill="white", outline="#ddd6ca", width=2)
        fitted = fit_contain(img, (x2 - x1 - 24, y2 - y1 - 54))
        canvas.paste(fitted, (x1 + 12, y1 + 12))
        draw.text((x1 + 14, y2 - 28), title, fill="#66605a", font=font(18))

    notes_box = (52, 752, 494, 868)
    draw.rounded_rectangle(notes_box, radius=16, fill="white", outline="#ddd6ca", width=2)
    draw.text((72, 776), footer, fill="#5f5850", font=font(19))
    draw.text((72, 814), "adult East Asian male / white crew neck tee / beige trousers / realistic cotton", fill="#7a746d", font=font(16))
    draw.text((72, 842), "no knitwear / no hoodie / no cat / no extra props / one clean square image", fill="#7a746d", font=font(16))

    ref = fit_cover(load(REF_DIR / ref_name), (960, 790))
    canvas.paste(ref, (578, 90))
    canvas.save(OUT_DIR / f"{task}_exact_board_v2.jpg", quality=96)


def main() -> None:
    build(
        "top02",
        "ref_top02.webp",
        "top02_onbody_round20_clean.jpg",
        "Top02: cooler wall, centered torso, left hand in pocket, right arm relaxed, calm crop.",
    )
    build(
        "top03",
        "ref_top03.webp",
        "top03_onbody_round20_clean.jpg",
        "Top03: warmer wall, slight body turn, left hand in pocket, right hand visible, softer mood.",
    )
    build(
        "top04",
        "ref_top04.webp",
        "top04_onbody_round20_clean.jpg",
        "Top04: darker wall, straighter front view, firmer torso shape, cleaner silhouette.",
    )


if __name__ == "__main__":
    main()
