from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageFilter, ImageOps


ROOT = Path(__file__).resolve().parent
STEP6 = ROOT / "step6_output" / "material_bundle"
ROUND10 = STEP6 / "T01_round10_2026-03-24"
ROUND13 = STEP6 / "T01_round13_2026-03-25"
OUT = ROUND13 / "04_harvested"


def feathered_polygon_mask(size: tuple[int, int], points: list[tuple[int, int]], blur: float = 7) -> Image.Image:
    mask = Image.new("L", size, 0)
    draw = ImageDraw.Draw(mask)
    draw.polygon(points, fill=255)
    return mask.filter(ImageFilter.GaussianBlur(blur))


def bright_garment_mask(base: Image.Image, threshold: int = 135) -> Image.Image:
    rgb = base.convert("RGB")
    out = Image.new("L", rgb.size, 0)
    src = rgb.load()
    dst = out.load()
    for y in range(rgb.height):
        for x in range(rgb.width):
            r, g, b = src[x, y]
            brightness = (r + g + b) // 3
            spread = max(r, g, b) - min(r, g, b)
            if brightness > threshold and spread < 34:
                dst[x, y] = 255
    return out.filter(ImageFilter.GaussianBlur(2))


def colorize_region(base: Image.Image, mask: Image.Image, rgb: tuple[int, int, int]) -> Image.Image:
    gray = ImageOps.grayscale(base)
    shadow = tuple(max(0, int(v * 0.45)) for v in rgb)
    highlight = tuple(min(255, int(v * 1.06 + 6)) for v in rgb)
    colored = ImageOps.colorize(gray, shadow, highlight).convert("RGB")
    out = base.copy()
    out.paste(colored, (0, 0), mask)
    return out


def colorize_region_strong(base: Image.Image, mask: Image.Image, rgb: tuple[int, int, int]) -> Image.Image:
    gray = ImageOps.grayscale(base)
    shadow = tuple(max(0, int(v * 0.28)) for v in rgb)
    highlight = tuple(min(255, int(v * 0.92 + 4)) for v in rgb)
    colored = ImageOps.colorize(gray, shadow, highlight).convert("RGB")
    out = base.copy()
    out.paste(colored, (0, 0), mask)
    return out


def build_stack() -> Path:
    base = Image.open(ROUND10 / "03_lovart_raw" / "stack_exact_final.png").convert("RGB")
    size = base.size
    garment = bright_garment_mask(base, 110)

    # Top drape, same taupe-gray feel as reference.
    top_drape = feathered_polygon_mask(
        size,
        [(210, 85), (748, 85), (744, 268), (541, 270), (388, 256), (193, 655), (30, 713), (22, 642), (118, 518), (248, 271)],
        blur=8,
    )
    # Inner stacked shirts from top to bottom.
    white_layer = feathered_polygon_mask(size, [(385, 202), (713, 208), (714, 320), (376, 322)], blur=5)
    ivory_layer = feathered_polygon_mask(size, [(368, 296), (716, 302), (714, 394), (357, 398)], blur=5)
    taupe_layer = feathered_polygon_mask(size, [(353, 382), (717, 388), (714, 478), (340, 482)], blur=5)
    charcoal_layer = feathered_polygon_mask(size, [(340, 467), (720, 474), (719, 558), (326, 563)], blur=5)
    black_layer = feathered_polygon_mask(size, [(327, 546), (722, 552), (721, 645), (314, 652)], blur=5)

    top_drape = ImageChops.multiply(top_drape, garment)
    white_layer = ImageChops.multiply(white_layer, garment)
    ivory_layer = ImageChops.multiply(ivory_layer, garment)
    taupe_layer = ImageChops.multiply(taupe_layer, garment)
    charcoal_layer = ImageChops.multiply(charcoal_layer, garment)
    black_layer = ImageChops.multiply(black_layer, garment)

    out = base
    out = colorize_region_strong(out, top_drape, (176, 159, 148))
    out = colorize_region_strong(out, white_layer, (243, 241, 237))
    out = colorize_region_strong(out, ivory_layer, (226, 216, 198))
    out = colorize_region_strong(out, taupe_layer, (169, 158, 147))
    out = colorize_region_strong(out, charcoal_layer, (118, 122, 129))
    out = colorize_region_strong(out, black_layer, (28, 28, 30))

    dest = OUT / "stack_exact_multicolor_round13_v1.jpg"
    out.save(dest, quality=96)
    return dest


def build_fold() -> Path:
    base = Image.open(OUT / "lovart_fold_retry_simple.jpg").convert("RGB")
    size = base.size
    garment = bright_garment_mask(base, 150)

    top = feathered_polygon_mask(size, [(0, 0), (398, 0), (351, 162), (194, 119), (0, 212)], blur=6)
    second = feathered_polygon_mask(size, [(172, 85), (503, 80), (485, 254), (130, 274)], blur=6)
    third = feathered_polygon_mask(size, [(55, 195), (408, 188), (420, 365), (0, 396)], blur=6)
    bottom = feathered_polygon_mask(size, [(0, 332), (356, 278), (512, 331), (423, 512), (25, 512)], blur=6)

    top = ImageChops.multiply(top, garment)
    second = ImageChops.multiply(second, garment)
    third = ImageChops.multiply(third, garment)
    bottom = ImageChops.multiply(bottom, garment)

    out = base
    out = colorize_region(out, top, (170, 154, 148))
    out = colorize_region(out, second, (245, 243, 239))
    out = colorize_region(out, third, (231, 223, 206))
    out = colorize_region(out, bottom, (183, 186, 192))

    dest = OUT / "fold_exact_multicolor_round13_v1.jpg"
    out.save(dest, quality=96)
    return dest


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    print(build_stack())
    print(build_fold())


if __name__ == "__main__":
    main()
