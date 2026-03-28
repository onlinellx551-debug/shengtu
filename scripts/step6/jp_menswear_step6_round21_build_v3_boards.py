from __future__ import annotations

import json
from pathlib import Path

import requests
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
STEP6 = ROOT / "step6_output"
ROUND20 = STEP6 / "material_bundle" / "T01_round20_2026-03-26"
REF_DIR = ROUND20 / "01_reference_exact"
OUT_DIR = ROUND20 / "01_input_boards_exact_v3"
URLS_JSON = ROUND20 / "05_probe_lovart" / "round20_exact_tmpfiles_urls.json"
ORIG = STEP6 / "T01_案例版详情素材_2026-03-20" / "01_同款精选原图"
GENERATED = ROUND20 / "02_generated"


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


def load(path: Path) -> Image.Image:
    return Image.open(path).convert("RGB")


def panel_card(canvas: Image.Image, rect: tuple[int, int, int, int], title: str, image: Image.Image) -> None:
    draw = ImageDraw.Draw(canvas)
    x1, y1, x2, y2 = rect
    draw.rounded_rectangle((x1, y1, x2, y2), radius=16, fill="white", outline="#ddd6ca", width=2)
    fitted = fit_contain(image, (x2 - x1 - 24, y2 - y1 - 56))
    canvas.paste(fitted, (x1 + 12, y1 + 12))
    draw.text((x1 + 14, y2 - 28), title, fill="#66605a", font=font(18))


def build_onbody(task: str, ref_name: str, color_title: str, flat_name: str, winner_name: str, footer: str) -> Path:
    canvas = Image.new("RGB", (1600, 920), "#f4f1ea")
    draw = ImageDraw.Draw(canvas)
    draw.rounded_rectangle((26, 26, 520, 894), radius=24, fill="#fbfaf7", outline="#dbd4c8", width=2)
    draw.rounded_rectangle((548, 26, 1574, 894), radius=24, fill="#fbfaf7", outline="#dbd4c8", width=2)
    draw.text((52, 48), "OUR TEE / KEEP THIS PRODUCT", fill="#44403a", font=font(32, True))
    draw.text((576, 48), "EXACT TARGET / MATCH THIS SHOT", fill="#44403a", font=font(32, True))

    panel_card(canvas, (52, 110, 494, 300), color_title, load(ORIG / flat_name))
    panel_card(canvas, (52, 324, 494, 514), "CURRENT BEST", load(GENERATED / winner_name))
    panel_card(canvas, (52, 538, 494, 728), "COLLAR DETAIL", load(ORIG / "白色领口特写.jpg"))

    draw.rounded_rectangle((52, 752, 494, 868), radius=16, fill="white", outline="#ddd6ca", width=2)
    draw.text((72, 776), footer, fill="#5f5850", font=font(19))
    draw.text((72, 814), "adult East Asian male / real cotton tee / beige trousers / no knitwear", fill="#7a746d", font=font(16))
    draw.text((72, 842), "keep the requested color exactly / one clean square image only", fill="#7a746d", font=font(16))

    ref = fit_cover(load(REF_DIR / ref_name), (960, 790))
    canvas.paste(ref, (578, 90))

    out = OUT_DIR / f"{task}_exact_board_v3.jpg"
    canvas.save(out, quality=96)
    return out


def build_detail(task: str, ref_name: str, cards: list[tuple[str, Path]], notes: list[str]) -> Path:
    canvas = Image.new("RGB", (1600, 920), "#f4f1ea")
    draw = ImageDraw.Draw(canvas)
    draw.rounded_rectangle((26, 26, 520, 894), radius=24, fill="#fbfaf7", outline="#dbd4c8", width=2)
    draw.rounded_rectangle((548, 26, 1574, 894), radius=24, fill="#fbfaf7", outline="#dbd4c8", width=2)
    draw.text((52, 48), "OUR TEE / REAL MATERIAL", fill="#44403a", font=font(32, True))
    draw.text((576, 48), "EXACT TARGET / MATCH THIS SHOT", fill="#44403a", font=font(32, True))

    boxes = [
        (52, 110, 494, 290),
        (52, 314, 270, 504),
        (276, 314, 494, 504),
        (52, 528, 494, 708),
    ]
    for rect, (title, path) in zip(boxes, cards):
        panel_card(canvas, rect, title, load(path))

    draw.rounded_rectangle((52, 732, 494, 868), radius=16, fill="white", outline="#ddd6ca", width=2)
    y = 756
    for note in notes:
        draw.text((70, y), note, fill="#645d55", font=font(18))
        y += 24

    ref = fit_cover(load(REF_DIR / ref_name), (960, 790))
    canvas.paste(ref, (578, 90))
    out = OUT_DIR / f"{task}_exact_board_v3.jpg"
    canvas.save(out, quality=96)
    return out


def upload_tmpfiles(path: Path) -> dict[str, str]:
    with path.open("rb") as f:
        resp = requests.post(
            "https://tmpfiles.org/api/v1/upload",
            files={"file": (path.name, f, "image/jpeg")},
            timeout=120,
        )
    resp.raise_for_status()
    data = resp.json()["data"]
    url = data["url"].replace("http://", "https://")
    direct_url = url.replace("https://tmpfiles.org/", "https://tmpfiles.org/dl/")
    return {"url": url, "direct_url": direct_url}


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    built = {
        "top02_exact_board_round20.jpg": build_onbody(
            "top02",
            "ref_top02.webp",
            "WHITE FLAT",
            "白色平铺.jpg",
            "top02_onbody_round20_clean.jpg",
            "Top02: same white color as top01, same cool wall, same centered torso.",
        ),
        "top03_exact_board_round20.jpg": build_onbody(
            "top03",
            "ref_top03.webp",
            "IVORY FLAT",
            "米白平铺.jpg",
            "top03_onbody_round20_clean.jpg",
            "Top03: ivory / beige tee, warmer wall, slight body turn, premium quiet mood.",
        ),
        "top04_exact_board_round20.jpg": build_onbody(
            "top04",
            "ref_top04.webp",
            "GRAY FLAT",
            "灰色平铺.jpg",
            "top04_onbody_round20_clean.jpg",
            "Top04: heather gray tee, darker wall, firmer front view, premium silhouette.",
        ),
        "top05_exact_board_round20.jpg": build_detail(
            "top05",
            "ref_top05.webp",
            [
                ("WHITE FLAT", ORIG / "白色平铺.jpg"),
                ("IVORY FLAT", ORIG / "米白平铺.jpg"),
                ("GRAY FLAT", ORIG / "灰色平铺.jpg"),
                ("HANGER COLOR RANGE", ORIG / "排挂近景.jpg"),
            ],
            [
                "must be folded short sleeve T-shirts, not sweaters or sweatshirts",
                "keep one light draped tee edge only, not a long sleeve garment",
                "warm beige tabletop, premium real photo, orderly stack",
                "colors only from our real range: white / ivory / gray / black",
            ],
        ),
        "top06_exact_board_round20.jpg": build_detail(
            "top06",
            "ref_top06.webp",
            [
                ("WHITE FOLD", ORIG / "白色折叠特写.jpg"),
                ("WHITE COLLAR", ORIG / "白色领口特写.jpg"),
                ("WHITE FLAT", ORIG / "白色平铺.jpg"),
                ("IVORY COLLAR", ORIG / "领口近景.jpg"),
            ],
            [
                "must read as thick cotton jersey fold, not knit rib fabric",
                "show edge thickness and cotton texture clearly",
                "no sweater rib body, no hoodie, no fantasy fabric",
                "keep the exact fold composition from the target",
            ],
        ),
        "top07_exact_board_round20.jpg": build_detail(
            "top07",
            "ref_top07.webp",
            [
                ("WHITE FLAT", ORIG / "白色平铺.jpg"),
                ("IVORY FLAT", ORIG / "米白平铺.jpg"),
                ("GRAY FLAT", ORIG / "灰色平铺.jpg"),
                ("WHITE COLLAR", ORIG / "白色领口特写.jpg"),
            ],
            [
                "must be folded short sleeve tees with visible collar openings",
                "use white / ivory / gray tones only",
                "natural labels are allowed but must stay minimal",
                "real tabletop product photo, premium but not over-styled",
            ],
        ),
        "top08_exact_board_round20.jpg": build_detail(
            "top08",
            "ref_top08.webp",
            [
                ("WHITE COLLAR", ORIG / "白色领口特写.jpg"),
                ("IVORY COLLAR", ORIG / "领口近景.jpg"),
                ("WHITE FOLD", ORIG / "白色折叠特写.jpg"),
                ("WHITE ONBODY", GENERATED / "top02_onbody_round20_clean.jpg"),
            ],
            [
                "focus on collar structure on neck / neck-to-collar relationship",
                "must still read as our thick cotton tee, not sweatshirt or knitwear",
                "premium skin tone, realistic neck texture, clean square frame",
                "no text, no props, no AI fantasy smoothing",
            ],
        ),
    }

    payload = json.loads(URLS_JSON.read_text(encoding="utf-8"))
    for key, path in built.items():
        payload[key] = upload_tmpfiles(path)
    URLS_JSON.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(URLS_JSON)


if __name__ == "__main__":
    main()
