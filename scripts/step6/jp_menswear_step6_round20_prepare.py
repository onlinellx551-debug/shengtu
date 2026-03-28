from __future__ import annotations

import shutil
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
STEP6 = ROOT / "step6_output" / "material_bundle"
ROUND19 = STEP6 / "T01_round19_2026-03-26"
ROUND20 = STEP6 / "T01_round20_2026-03-26"
BOARD_DIR = ROUND20 / "01_input_boards"
REAL_DIR = ROUND20 / "01_real_candidates"
GEN_DIR = ROUND20 / "02_generated"
ASSET_DIR = ROUND20 / "03_assets"
PREVIEW_DIR = ROUND20 / "04_web_preview"

FONT_REG = Path(r"C:\Windows\Fonts\msyh.ttc")
FONT_BOLD = Path(r"C:\Windows\Fonts\msyhbd.ttc")


def font(size: int, bold: bool = False):
    path = FONT_BOLD if bold and FONT_BOLD.exists() else FONT_REG
    if path.exists():
        return ImageFont.truetype(str(path), size=size)
    return ImageFont.load_default()


def reset_dir(path: Path) -> None:
    if path.exists():
        shutil.rmtree(path)
    path.mkdir(parents=True, exist_ok=True)


def copy(src: Path, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dest)


def cover(src: Path, size: tuple[int, int]) -> Image.Image:
    img = Image.open(src).convert("RGB")
    sw, sh = img.size
    tw, th = size
    scale = max(tw / sw, th / sh)
    nw, nh = int(sw * scale), int(sh * scale)
    img = img.resize((nw, nh), Image.LANCZOS)
    left = max(0, (nw - tw) // 2)
    top = max(0, (nh - th) // 2)
    return img.crop((left, top, left + tw, top + th))


def build_board(title: str, ref_path: Path, product_path: Path, out_path: Path, note: str) -> None:
    canvas = Image.new("RGB", (2048, 2048), "#f4efe8")
    draw = ImageDraw.Draw(canvas)

    ref_img = cover(ref_path, (1260, 1560))
    product_img = cover(product_path, (520, 520))

    canvas.paste(ref_img, (80, 280))
    draw.rounded_rectangle((1390, 315, 1980, 905), radius=28, outline="#ddd3c5", width=4, fill="#fbf8f3")
    canvas.paste(product_img, (1420, 345))

    draw.text((86, 74), title, font=font(72, True), fill="#1f1a17")
    draw.text((88, 166), "左：参考镜头  右：我们的产品（严格换成白T，不换镜头结构）", font=font(36), fill="#6c655d")
    draw.text((1420, 914), note, font=font(30), fill="#514a42")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(out_path, quality=95)


def crop_tail_detail(src: Path, dest: Path) -> None:
    img = Image.open(src).convert("RGB")
    w, h = img.size
    cropped = img.crop((80, 250, w - 80, h - 120))
    cropped = cropped.resize((768, 768), Image.LANCZOS)
    dest.parent.mkdir(parents=True, exist_ok=True)
    cropped.save(dest, quality=96)


def main() -> None:
    reset_dir(ROUND20)
    BOARD_DIR.mkdir(parents=True, exist_ok=True)
    REAL_DIR.mkdir(parents=True, exist_ok=True)
    GEN_DIR.mkdir(parents=True, exist_ok=True)
    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    PREVIEW_DIR.mkdir(parents=True, exist_ok=True)

    round18_preview = STEP6 / "T01_round18_2026-03-26" / "04_web_preview"
    round18_generated = STEP6 / "T01_round18_2026-03-26" / "01_generated"

    product_flat = ROOT / "step6_output" / "gen_stage" / "product_white_tee.jpg"
    top01_real = ROUND19 / "01_real_candidates" / "top01_real_front_whitespace.jpg"

    copy(top01_real, REAL_DIR / "top01_real_front_whitespace.jpg")
    copy(round18_generated / "G02.png", GEN_DIR / "top02_onbody_round18.png")
    copy(round18_generated / "G12.png", GEN_DIR / "benefit01_lifestyle_round18.png")
    copy(round18_generated / "G15.png", GEN_DIR / "staff_01_round18.png")
    copy(round18_generated / "G16.png", GEN_DIR / "staff_02_round18.png")
    copy(round18_generated / "G17.png", GEN_DIR / "staff_03_round18.png")

    for name in [
        "top_05_stack_round14.jpg",
        "top_07_fold_round14.jpg",
        "top_08_collar.jpg",
        "tail_02_black_flat.jpg",
        "tail_03_white_flat.jpg",
        "tail_04_gray_flat.jpg",
        "tail_05_ivory_flat.jpg",
        "tail_06_olive_flat.jpg",
        "tail_07_black_flat.jpg",
        "tail_08_gray_flat.jpg",
    ]:
        copy(round18_preview / "assets" / name, ASSET_DIR / name)

    copy(round18_preview / "assets" / "top_06_fold.jpg", ASSET_DIR / "top_06_fold.jpg")
    copy(round18_preview / "assets" / "top_08_collar.jpg", ASSET_DIR / "benefit_02_clean.jpg")
    copy(round18_preview / "assets" / "top_07_fold_round14.jpg", ASSET_DIR / "benefit_03_clean.jpg")
    crop_tail_detail(round18_generated / "G15.png", ASSET_DIR / "tail_01_hand_detail.jpg")

    top03_ref = ROOT / "step6_output" / "gen_stage" / "G03_ref.webp"
    top04_ref = ROUND19 / "04_ref_only" / "top04_ref_exact2.png"
    top06_ref = ROUND19 / "04_ref_only" / "top06_ref_exact2.png"

    build_board(
        "T01 顶部上身位 03 参考板",
        top03_ref,
        product_flat,
        BOARD_DIR / "top03_onbody_board_round20.jpg",
        "必须是真实短袖白T，成熟男装电商质感，不要卫衣，不要针织，不要AI味。",
    )
    build_board(
        "T01 顶部上身位 04 参考板",
        top04_ref,
        product_flat,
        BOARD_DIR / "top04_onbody_board_round20.jpg",
        "保持右下格上身镜头的裁切和姿态，严格替换成我们的厚实白T。",
    )
    build_board(
        "T01 顶部结构位 06 参考板",
        top06_ref,
        product_flat,
        BOARD_DIR / "top06_structure_board_round20.jpg",
        "保留折叠边、厚度和针织般卷边的摄影感觉，但必须是白T真实结构近景。",
    )

    (ROUND20 / "00_notes.txt").write_text(
        "round20 开始按正确槽位顺序重建：top01 用真实白底单品，top03/top04 补生成，top06 结构位微重生，其余位先回填较强版本。",
        encoding="utf-8",
    )
    print(ROUND20)


if __name__ == "__main__":
    main()
