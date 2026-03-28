from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
STEP6 = ROOT / "step6_output" / "material_bundle"
ROUND17 = STEP6 / "T01_round17_2026-03-25"
BOARD_DIR = ROUND17 / "01_input_boards"
GEN_STAGE = ROOT / "step6_output" / "gen_stage"

FONT_REG = Path(r"C:\Windows\Fonts\msyh.ttc")
FONT_BOLD = Path(r"C:\Windows\Fonts\msyhbd.ttc")


def font(size: int, bold: bool = False):
    path = FONT_BOLD if bold and FONT_BOLD.exists() else FONT_REG
    if path.exists():
        return ImageFont.truetype(str(path), size=size)
    return ImageFont.load_default()


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


def board(title: str, ref_path: Path, product_path: Path, out_path: Path) -> None:
    canvas = Image.new("RGB", (2048, 2048), "#f4efe8")
    draw = ImageDraw.Draw(canvas)

    ref_img = cover(ref_path, (1220, 1600))
    product_img = cover(product_path, (540, 540))

    canvas.paste(ref_img, (90, 280))
    canvas.paste(product_img, (1400, 340))

    draw.rounded_rectangle((1370, 310, 1970, 910), radius=26, outline="#d8cec0", width=4, fill="#fbf8f3")
    canvas.paste(product_img, (1400, 340))

    title_font = font(72, bold=True)
    sub_font = font(36, bold=False)
    note_font = font(32, bold=False)

    draw.text((92, 72), title, font=title_font, fill="#1f1a17")
    draw.text((94, 165), "左：参照镜头  右：我们的产品（严格换成白T，不换镜头结构）", font=sub_font, fill="#6c655d")
    draw.text((1402, 900), "必须保留：真实棉质T恤、成熟男装摄影感、不要AI味", font=note_font, fill="#524a42")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(out_path, quality=95)


def main() -> None:
    BOARD_DIR.mkdir(parents=True, exist_ok=True)

    product = GEN_STAGE / "product_white_tee.jpg"
    board("T01 顶部上身位 01 参考板", GEN_STAGE / "G02_ref.webp", product, BOARD_DIR / "top01_onbody_board.jpg")
    board("T01 顶部上身位 03 参考板", GEN_STAGE / "G03_ref.webp", product, BOARD_DIR / "top03_onbody_board.jpg")
    board("T01 中段生活方式位 参考板", GEN_STAGE / "G12_ref.webp", product, BOARD_DIR / "benefit01_lifestyle_board.jpg")

    print(BOARD_DIR)


if __name__ == "__main__":
    main()
