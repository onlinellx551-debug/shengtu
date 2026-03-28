from __future__ import annotations

import time
import urllib.request
from pathlib import Path

from PIL import Image
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
STEP6 = ROOT / "step6_output" / "material_bundle"
ROUND17 = STEP6 / "T01_round17_2026-03-25"
RAW_DIR = ROUND17 / "03_jimeng_raw"
HARVEST_DIR = ROUND17 / "04_harvested"
PROBE_DIR = ROUND17 / "06_probe_jimeng"

JIMENG_WS = "ws://127.0.0.1:54902/devtools/browser/03b0c484-60cf-482c-b89b-a60766f67580"

TASKS = [
    {
        "page_index": 14,
        "name": "top01_onbody_round17",
        "board": ROUND17 / "01_input_boards" / "top01_onbody_board.jpg",
        "png": RAW_DIR / "top01_onbody_round17.png",
        "jpg": HARVEST_DIR / "jimeng_top01_onbody_round17.jpg",
        "prompt": (
            "请严格参考我上传的参考板。保留参考图的人像构图、镜头距离、姿势、裤装和高级日系男装电商摄影感，"
            "但把上衣完整替换成我们自己的重磅白色圆领短袖T恤。必须是成熟成年东亚男模，必须是真实棉质白T，"
            "不是卫衣，不是针织，不是长袖。白T要有厚度感、自然褶皱、真实袖口和领口缝线，不要AI感。"
            "不要文字，不要logo，不要拼贴，只要一张正方形成品图。"
        ),
    },
    {
        "page_index": 10,
        "name": "top03_onbody_round17",
        "board": ROUND17 / "01_input_boards" / "top03_onbody_board.jpg",
        "png": RAW_DIR / "top03_onbody_round17.png",
        "jpg": HARVEST_DIR / "jimeng_top03_onbody_round17.jpg",
        "prompt": (
            "请严格参考我上传的参考板，生成另一张与参考镜头一一对应的成熟日系男装上身图。"
            "保留原参考图的构图和姿态，但上衣必须替换成我们自己的重磅白色圆领短袖T恤。"
            "不能有针织或长袖特征，必须是厚实真实的短袖白T，质感自然，难以看出AI。"
            "不要文字，不要logo，不要拼贴，只要一张正方形成品图。"
        ),
    },
    {
        "page_index": 5,
        "name": "benefit01_lifestyle_round17",
        "board": ROUND17 / "01_input_boards" / "benefit01_lifestyle_board.jpg",
        "png": RAW_DIR / "benefit01_lifestyle_round17.png",
        "jpg": HARVEST_DIR / "jimeng_benefit01_lifestyle_round17.jpg",
        "prompt": (
            "请严格参考我上传的参考板，生成中段生活方式卖点图。保留参考图的坐姿、暖色室内光线和成熟男装氛围，"
            "但把上衣替换成我们自己的重磅白色圆领短袖T恤。要像真实电商拍摄，不要AI味，不要文字，不要logo，"
            "不要拼贴，只要一张正方形成品图。"
        ),
    },
]


def ensure_dirs() -> None:
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    HARVEST_DIR.mkdir(parents=True, exist_ok=True)
    PROBE_DIR.mkdir(parents=True, exist_ok=True)


def visible_submit(page):
    best = None
    loc = page.locator("button")
    for i in range(loc.count()):
        btn = loc.nth(i)
        try:
            box = btn.bounding_box()
            cls = btn.get_attribute("class") or ""
        except Exception:
            box = None
            cls = ""
        if not box:
            continue
        if "lv-btn-primary" not in cls:
            continue
        if box["width"] < 28 or box["height"] < 28 or box["y"] < 700 or box["x"] > 900:
            continue
        if best is None or (box["y"], box["x"]) > best[0]:
            best = ((box["y"], box["x"]), btn)
    if best:
        return best[1]
    raise RuntimeError("Jimeng visible submit button not found")


def upload_board(page, board: Path) -> None:
    inputs = page.locator('input[type="file"]')
    if inputs.count() == 0:
        raise RuntimeError("Jimeng file input not found")
    target = 1 if inputs.count() > 1 else 0
    inputs.nth(target).set_input_files(str(board))
    page.wait_for_timeout(4000)


def wait_for_result(page) -> str:
    end = time.time() + 420
    while time.time() < end:
        imgs = page.locator("div.agentic-image-item-_quyBP img.image-TLmgkP")
        if imgs.count():
            src = imgs.last.get_attribute("src") or ""
            if "http" in src:
                return src
        page.wait_for_timeout(3000)
    raise TimeoutError("Jimeng wait timeout")


def visible_editor(page):
    loc = page.locator('[contenteditable="true"]')
    for i in range(loc.count()):
        ed = loc.nth(i)
        try:
            box = ed.bounding_box()
        except Exception:
            box = None
        if not box:
            continue
        if box["x"] < 800 and box["width"] > 200 and box["height"] > 50:
            return ed
    raise RuntimeError("Jimeng visible editor not found")


def download(url: str, dest: Path) -> None:
    request = urllib.request.Request(
        url,
        headers={"User-Agent": "Mozilla/5.0", "Referer": "https://jimeng.jianying.com/"},
    )
    with urllib.request.urlopen(request) as resp:
        dest.write_bytes(resp.read())


def convert_to_jpg(src: Path, dest: Path) -> None:
    img = Image.open(src).convert("RGB")
    img.save(dest, quality=96)


def run_task(page, task: dict[str, Path | str]) -> None:
    page.goto("https://jimeng.jianying.com/ai-tool/canvas", wait_until="load")
    page.wait_for_timeout(2500)
    board = Path(task["board"])
    upload_board(page, board)

    editor = visible_editor(page)
    editor.click(force=True)
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")
    page.keyboard.insert_text(str(task["prompt"]))
    page.wait_for_timeout(500)

    probe_base = PROBE_DIR / str(task["name"])
    page.screenshot(path=str(probe_base.with_suffix(".before.png")), full_page=False)
    visible_submit(page).click(force=True)
    page.wait_for_timeout(1200)
    page.screenshot(path=str(probe_base.with_suffix(".after.png")), full_page=False)

    result_url = wait_for_result(page)
    png_path = Path(task["png"])
    download(result_url, png_path)
    convert_to_jpg(png_path, Path(task["jpg"]))
    page.screenshot(path=str(probe_base.with_suffix(".done.png")), full_page=False)
    print(task["name"], result_url)


def main() -> None:
    ensure_dirs()
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(JIMENG_WS)
        ctx = browser.contexts[0]
        for task in TASKS:
            page = ctx.pages[int(task["page_index"])]
            run_task(page, task)
        browser.close()


if __name__ == "__main__":
    main()
