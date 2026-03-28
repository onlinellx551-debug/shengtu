from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageChops, ImageFilter, ImageOps


ROOT = Path(__file__).resolve().parent
STEP6 = ROOT / "step6_output" / "素材包"
ROUND12 = STEP6 / "T01_round12_2026-03-25"
ROUND13 = STEP6 / "T01_round13_2026-03-25"
ASSETS = ROUND12 / "04_web_preview" / "assets"
OUT = ROUND13 / "04_harvested"


def load(path: Path) -> Image.Image:
    return Image.open(path).convert("RGB")


def mask_from_white(img: Image.Image, threshold: int = 238) -> Image.Image:
    gray = ImageOps.grayscale(img)
    inv = ImageOps.invert(gray)
    mask = inv.point(lambda v: 255 if v > (255 - threshold) else int(max(0, (v - 8) * 1.5)))
    return mask.filter(ImageFilter.GaussianBlur(1.2))


def recolor_with_luminance(img: Image.Image, rgb: tuple[int, int, int]) -> Image.Image:
    gray = ImageOps.grayscale(img)
    shadow = tuple(max(0, int(v * 0.52)) for v in rgb)
    highlight = tuple(min(255, int(v * 1.06 + 8)) for v in rgb)
    return ImageOps.colorize(gray, shadow, highlight).convert("RGB")


def paste_with_shadow(
    base: Image.Image,
    fg: Image.Image,
    mask: Image.Image,
    xy: tuple[int, int],
    shadow_offset: tuple[int, int] = (10, 12),
    shadow_blur: int = 16,
    shadow_alpha: int = 110,
) -> None:
    shadow = Image.new("RGBA", base.size, (0, 0, 0, 0))
    sh = Image.new("RGBA", fg.size, (0, 0, 0, 255))
    sh.putalpha(mask)
    shadow.paste(sh, (xy[0] + shadow_offset[0], xy[1] + shadow_offset[1]), sh)
    shadow = shadow.filter(ImageFilter.GaussianBlur(shadow_blur))
    alpha = shadow.getchannel("A").point(lambda v: int(v * shadow_alpha / 255))
    shadow.putalpha(alpha)
    base.alpha_composite(shadow)
    rgba = fg.convert("RGBA")
    rgba.putalpha(mask)
    base.alpha_composite(rgba, xy)


def build_fold() -> Path:
    bg = Image.new("RGBA", (768, 768), "#cba27a")
    cloth = load(ASSETS / "detail_fold.jpg")
    mask = mask_from_white(cloth, 245)

    colors = [
        (168, 154, 148),  # taupe gray
        (245, 243, 238),  # white
        (234, 226, 208),  # ivory
        (183, 186, 191),  # light gray
    ]
    positions = [(42, 18), (92, 148), (120, 290), (70, 430)]
    sizes = [(560, 340), (528, 322), (510, 310), (520, 314)]

    for i, (color, pos, size) in enumerate(zip(colors, positions, sizes)):
        recolored = recolor_with_luminance(cloth, color).resize(size, Image.LANCZOS)
        m = mask.resize(size, Image.LANCZOS)
        if i > 0:
            # Hide repetitive lower labels slightly.
            wipe = Image.new("L", size, 255)
            wipe.paste(0, (120, 70, 280, 180))
            m = ImageChops.multiply(m, wipe)
        paste_with_shadow(bg, recolored, m, pos, shadow_offset=(8, 10), shadow_blur=14, shadow_alpha=85)

    out = OUT / "local_fold_real_round13_v1.jpg"
    bg.convert("RGB").save(out, quality=96)
    return out


def crop_spool() -> Image.Image:
    source = load(OUT / "lovart_stack_retry_simple.jpg")
    crop = source.crop((0, 30, 180, 300))
    return crop


def build_stack() -> Path:
    bg = Image.new("RGBA", (768, 768), "#cba27a")

    spool = crop_spool()
    bg.alpha_composite(spool.convert("RGBA"), (14, 210))

    flat_files = [
        (ASSETS / "tail_07_black_flat.jpg", (22, 22, 24)),
        (ASSETS / "tail_04_gray_flat.jpg", (176, 180, 185)),
        (ASSETS / "tail_05_ivory_flat.jpg", (223, 215, 199)),
        (ASSETS / "tail_03_white_flat.jpg", (244, 242, 238)),
    ]
    stack_y = [518, 438, 360, 286]
    stack_x = 360
    stack_sizes = [(322, 84), (332, 88), (344, 92), (360, 100)]

    for (path, color), y, size in zip(flat_files, stack_y, stack_sizes):
        tee = load(path)
        body = tee.crop((150, 200, 620, 640))
        body_mask = mask_from_white(body, 246).resize(size, Image.LANCZOS)
        body_rgb = recolor_with_luminance(body, color).resize(size, Image.LANCZOS)
        paste_with_shadow(bg, body_rgb, body_mask, (stack_x, y), shadow_offset=(6, 8), shadow_blur=10, shadow_alpha=70)

    top = load(ASSETS / "tail_05_ivory_flat.jpg")
    top_mask = mask_from_white(top, 246)
    top_rgb = recolor_with_luminance(top, (176, 160, 151))
    top_rgb = top_rgb.resize((510, 510), Image.LANCZOS).rotate(-8, expand=True, resample=Image.BICUBIC)
    top_mask = top_mask.resize((510, 510), Image.LANCZOS).rotate(-8, expand=True, resample=Image.BICUBIC)
    # Crop the top shirt so it behaves like a draped short-sleeve tee.
    top_rgb = top_rgb.crop((28, 54, 520, 412))
    top_mask = top_mask.crop((28, 54, 520, 412))
    paste_with_shadow(bg, top_rgb, top_mask, (244, 126), shadow_offset=(10, 12), shadow_blur=16, shadow_alpha=95)

    out = OUT / "local_stack_real_round13_v1.jpg"
    bg.convert("RGB").save(out, quality=96)
    return out


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    print(build_stack())
    print(build_fold())


if __name__ == "__main__":
    main()
