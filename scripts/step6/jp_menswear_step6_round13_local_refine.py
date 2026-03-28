from __future__ import annotations

from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFilter


ROOT = Path(__file__).resolve().parent
STEP6 = ROOT / "step6_output" / "material_bundle"
ROUND10 = STEP6 / "T01_round10_2026-03-24"
ROUND12 = STEP6 / "T01_round12_2026-03-25"
ROUND13 = STEP6 / "T01_round13_2026-03-25"
HARVEST = ROUND13 / "04_harvested"


def load_rgb(path: Path) -> Image.Image:
    return Image.open(path).convert("RGB")


def polygon_mask(size: tuple[int, int], points: list[tuple[int, int]], blur: int = 10) -> Image.Image:
    mask = Image.new("L", size, 0)
    ImageDraw.Draw(mask).polygon(points, fill=255)
    return mask.filter(ImageFilter.GaussianBlur(blur))


def tint_region(base: Image.Image, current: Image.Image, mask: Image.Image, color: tuple[int, int, int], strength: float = 1.0) -> Image.Image:
    base_arr = np.asarray(base).astype(np.float32)
    cur_arr = np.asarray(current).astype(np.float32)
    mask_arr = np.asarray(mask).astype(np.float32) / 255.0
    lum = np.asarray(base.convert("L")).astype(np.float32) / 255.0
    color_arr = np.array(color, dtype=np.float32).reshape(1, 1, 3)
    shaded = lum[..., None] * color_arr
    mixed = cur_arr * (1.0 - mask_arr[..., None] * strength) + shaded * (mask_arr[..., None] * strength)
    return Image.fromarray(np.clip(mixed, 0, 255).astype(np.uint8))


def preserve_region(src: Image.Image, dst: Image.Image, box: tuple[int, int, int, int]) -> Image.Image:
    out = dst.copy()
    out.paste(src.crop(box), box)
    return out


def build_stack() -> Path:
    base = load_rgb(ROUND10 / "03_lovart_raw" / "stack_exact_final.png")
    cur = base.copy()
    w, h = base.size

    masks = [
        (
            [(240, 88), (598, 84), (754, 118), (760, 220), (646, 252), (528, 236), (436, 232), (300, 230), (188, 520), (120, 636), (38, 636), (160, 354), (210, 246)],
            (174, 164, 162),
            0.95,
        ),
        (
            [(346, 226), (584, 220), (628, 282), (584, 360), (340, 350), (292, 292)],
            (245, 245, 242),
            1.0,
        ),
        (
            [(318, 346), (608, 342), (652, 406), (612, 474), (314, 466), (266, 404)],
            (227, 220, 208),
            0.98,
        ),
        (
            [(300, 456), (626, 452), (668, 520), (624, 584), (294, 578), (246, 512)],
            (164, 159, 156),
            0.98,
        ),
        (
            [(278, 566), (644, 562), (682, 626), (640, 694), (254, 688), (208, 622)],
            (70, 72, 78),
            1.0,
        ),
    ]

    for points, color, strength in masks:
        cur = tint_region(base, cur, polygon_mask((w, h), points, blur=10), color, strength)

    # Darken the lowest visible strip to read closer to the reference bottom layer.
    bottom_strip = polygon_mask((w, h), [(264, 642), (650, 642), (692, 718), (260, 718)], blur=8)
    cur = tint_region(base, cur, bottom_strip, (24, 25, 29), 0.86)

    cur_arr = np.asarray(cur).copy()
    base_arr = np.asarray(base)
    # Keep the label crisp and realistic.
    cur_arr[205:334, 430:560] = base_arr[205:334, 430:560]
    # Restore the clean background around the yarn cone to avoid color spill.
    cur_arr[0:340, 0:360] = base_arr[0:340, 0:360]
    cur = Image.fromarray(cur_arr.astype(np.uint8))

    out = HARVEST / "stack_local_recolor_round13_v2.jpg"
    cur.save(out, quality=96)
    return out


def main() -> None:
    HARVEST.mkdir(parents=True, exist_ok=True)
    stack = build_stack()
    fold_src = ROUND12 / "04_harvested" / "lovart_fold_multicolor_exact.jpg"
    fold_dst = HARVEST / "fold_selected_round13_v2.jpg"
    load_rgb(fold_src).save(fold_dst, quality=96)
    print(stack)
    print(fold_dst)


if __name__ == "__main__":
    main()

