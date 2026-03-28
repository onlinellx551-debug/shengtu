from __future__ import annotations

import time
import urllib.request
from pathlib import Path

from PIL import Image
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
STEP6 = ROOT / "step6_output" / "material_bundle"
ROUND10 = STEP6 / "T01_round10_2026-03-24"
ROUND13 = STEP6 / "T01_round13_2026-03-25"
RAW_DIR = ROUND13 / "03_jimeng_raw"
HARVEST_DIR = ROUND13 / "04_harvested"
PROBE_DIR = ROUND13 / "06_probe_jimeng"

JIMENG_WS = "ws://127.0.0.1:54902/devtools/browser/03b0c484-60cf-482c-b89b-a60766f67580"

TASKS = [
    {
        "page_index": 4,
        "name": "stack_multicolor_round13_v3",
        "board": ROUND13 / "01_input_boards_v2" / "G05_stack_multicolor_board_v2.jpg",
        "png": RAW_DIR / "stack_multicolor_round13_v3.png",
        "jpg": HARVEST_DIR / "jimeng_stack_multicolor_round13_v3.jpg",
        "prompt": (
            "请严格参考我上传的参考板右侧构图，生成一张真实摄影质感的日系男装电商静物图。"
            "主体必须是我们自己的重磅圆领短袖T恤，不是毛衣，不是针织衫，不是长袖。"
            "画面必须整齐克制：几件T恤整齐堆叠在暖米色台面上，只允许最上层轻微自然垂落，但整体仍以整齐堆叠为主。"
            "左侧保留线轴，机位、裁切、道具比例尽量贴近参考图。"
            "颜色只能使用我们的真实颜色：白色、米白、浅灰、灰褐、炭灰、黑色。"
            "黑色只能作为底层或边缘层次，不能整件压成前景主角。"
            "必须自然露出我们自己的领标。"
            "面料必须像真实厚实棉质T恤汗布，短袖轮廓和领口细节要真实，不能看出AI痕迹。"
            "不要文字，不要logo，不要拼贴，不要参考板，只出一张正方形成品图。"
        ),
    },
    {
        "page_index": 5,
        "name": "fold_multicolor_round13_v3",
        "board": ROUND13 / "01_input_boards_v2" / "G07_fold_multicolor_board_v2.jpg",
        "png": RAW_DIR / "fold_multicolor_round13_v3.png",
        "jpg": HARVEST_DIR / "jimeng_fold_multicolor_round13_v3.jpg",
        "prompt": (
            "请严格参考我上传的参考板右侧构图，生成一张真实摄影质感的日系男装电商近景静物图。"
            "主体是我们自己的重磅圆领短袖T恤，多件T恤按参考图那样做紧凑的层叠折叠近景，构图节奏和裁切尽量一致。"
            "颜色只能使用我们的真实颜色：灰褐、白色、米白、浅灰。"
            "必须自然露出我们自己的领标。"
            "面料必须像真实厚实棉质T恤汗布，不是毛衣针织。"
            "领口、折痕、阴影、台面、暖色光都要真实，不能看出AI痕迹。"
            "不要文字，不要logo，不要拼贴，不要参考板，只出一张正方形成品图。"
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
        headers={
            "User-Agent": "Mozilla/5.0",
            "Referer": "https://jimeng.jianying.com/",
        },
    )
    with urllib.request.urlopen(request) as resp:
        dest.write_bytes(resp.read())


def convert_to_jpg(src: Path, dest: Path) -> None:
    img = Image.open(src).convert("RGB")
    img.save(dest, quality=96)


def run_task(page, task: dict[str, Path | str]) -> None:
    page.goto("https://jimeng.jianying.com/ai-tool/canvas", wait_until="load")
    page.wait_for_timeout(2500)
    design = page.locator("button", has_text="创意设计")
    if design.count():
        design.first.click(force=True)
        page.wait_for_timeout(600)
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
            try:
                run_task(page, task)
            except Exception as exc:
                try:
                    page.screenshot(path=str(PROBE_DIR / f"{task['name']}.error.png"), full_page=False, timeout=5000)
                except Exception:
                    pass
                raise RuntimeError(f"{task['name']} failed: {exc}") from exc
        browser.close()


if __name__ == "__main__":
    main()
