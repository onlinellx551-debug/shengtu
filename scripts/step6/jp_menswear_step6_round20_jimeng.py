from __future__ import annotations

import time
import urllib.request
from pathlib import Path

from PIL import Image
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
STEP6 = ROOT / "step6_output" / "material_bundle"
ROUND20 = STEP6 / "T01_round20_2026-03-26"
BOARD_DIR = ROUND20 / "01_input_boards"
RAW_DIR = ROUND20 / "02_generated"
PROBE_DIR = ROUND20 / "06_probe_jimeng"

JIMENG_WS = "ws://127.0.0.1:54902/devtools/browser/03b0c484-60cf-482c-b89b-a60766f67580"

TASKS = [
    {
        "name": "top03_onbody_round20_jimeng",
        "board": BOARD_DIR / "top03_onbody_board_round20.jpg",
        "png": RAW_DIR / "top03_onbody_round20_jimeng.png",
        "jpg": RAW_DIR / "top03_onbody_round20_jimeng.jpg",
        "prompt": (
            "请严格参考我上传的参考板，生成与参考镜头一一对应的成熟日系男装上身图。"
            "保留参考图的人体比例、镜头距离、手插袋姿态、米色裤装和安静的高级电商氛围，"
            "但把上衣完整替换成我们自己的重磅白色圆领短袖T恤。"
            "必须是真实厚实的300g棉质白T，不是针织，不是卫衣，不是长袖。"
            "要有自然褶皱、稳定领口、真实袖口和自然布重感。"
            "不要文字，不要logo，不要拼贴，不要参考板，只要一张正方形成品图。"
        ),
    },
    {
        "name": "top04_onbody_round20_jimeng",
        "board": BOARD_DIR / "top04_onbody_board_round20.jpg",
        "png": RAW_DIR / "top04_onbody_round20_jimeng.png",
        "jpg": RAW_DIR / "top04_onbody_round20_jimeng.jpg",
        "prompt": (
            "请严格参考我上传的参考板，生成与参考镜头一一对应的成熟日系男装上身图。"
            "保留右下格上身图的镜头裁切、背景、手插袋姿态和米色裤装，"
            "但上衣必须完整替换成我们自己的重磅白色圆领短袖T恤。"
            "必须是成人东亚男模，真实短袖白T，厚实棉质，不可有针织和长袖特征。"
            "要求安静高级、像真实日系男装电商拍摄，难以看出AI。"
            "不要文字，不要logo，不要拼贴，不要参考板，只要一张正方形成品图。"
        ),
    },
    {
        "name": "top06_structure_round20_jimeng",
        "board": BOARD_DIR / "top06_structure_board_round20.jpg",
        "png": RAW_DIR / "top06_structure_round20_jimeng.png",
        "jpg": RAW_DIR / "top06_structure_round20_jimeng.jpg",
        "prompt": (
            "请严格参考我上传的参考板，生成一张高质量电商静物结构近景。"
            "镜头要和参考图尽量一致：暖中性色背景、浅景深、折叠边在右侧前景。"
            "但产品必须是我们自己的重磅白色圆领短袖T恤，是真实厚实棉质面料，"
            "要看到领口/卷边/厚度/走线的真实结构感，不要针织，不要卫衣。"
            "不要文字，不要logo，不要拼贴，不要参考板，只要一张正方形成品图。"
        ),
    },
]


def ensure_dirs() -> None:
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    PROBE_DIR.mkdir(parents=True, exist_ok=True)


def visible_submit(page):
    loc = page.locator("button")
    best = None
    for i in range(loc.count()):
        btn = loc.nth(i)
        try:
            box = btn.bounding_box()
            cls = btn.get_attribute("class") or ""
        except Exception:
            continue
        if not box:
            continue
        if "lv-btn-primary" not in cls and "button" not in cls:
            continue
        if box["x"] < 800 or box["y"] < 700:
            continue
        if best is None or (box["y"], box["x"]) > best[0]:
            best = ((box["y"], box["x"]), btn)
    if best:
        return best[1]
    raise RuntimeError("Jimeng submit button not found")


def upload_board(page, board: Path) -> None:
    inputs = page.locator('input[type="file"]')
    if inputs.count() == 0:
        raise RuntimeError("Jimeng file input not found")
    target = 1 if inputs.count() > 1 else 0
    inputs.nth(target).set_input_files(str(board))
    page.wait_for_timeout(4500)


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
        if box["width"] > 200 and box["height"] > 50:
            return ed
    raise RuntimeError("Jimeng editor not found")


def wait_for_result(page) -> str:
    end = time.time() + 420
    while time.time() < end:
        imgs = page.locator("img")
        urls = []
        for i in range(imgs.count()):
            try:
                src = imgs.nth(i).get_attribute("src") or ""
                if "dreamina-sign.byteimg.com" in src or "tos-cn-i" in src:
                    urls.append(src)
            except Exception:
                continue
        if urls:
            return urls[-1]
        page.wait_for_timeout(3000)
    raise TimeoutError("Jimeng wait timeout")


def download(url: str, dest: Path) -> None:
    request = urllib.request.Request(
        url,
        headers={"User-Agent": "Mozilla/5.0", "Referer": "https://jimeng.jianying.com/"},
    )
    with urllib.request.urlopen(request) as resp:
        dest.write_bytes(resp.read())


def convert(src: Path, dest: Path) -> None:
    Image.open(src).convert("RGB").save(dest, quality=96)


def run_task(page, task: dict[str, Path | str]) -> None:
    page.goto("https://jimeng.jianying.com/ai-tool/canvas", wait_until="load")
    page.wait_for_timeout(2500)
    upload_board(page, Path(task["board"]))

    editor = visible_editor(page)
    editor.click(force=True)
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")
    page.keyboard.insert_text(str(task["prompt"]))
    page.wait_for_timeout(800)

    probe = PROBE_DIR / str(task["name"])
    page.screenshot(path=str(probe.with_suffix(".before.png")), full_page=False)
    btn = visible_submit(page)
    handle = btn.element_handle()
    if handle is None:
        raise RuntimeError("Jimeng submit handle not found")
    page.evaluate("(el) => el.click()", handle)
    page.wait_for_timeout(1200)
    page.screenshot(path=str(probe.with_suffix(".after.png")), full_page=False)

    result_url = wait_for_result(page)
    png = Path(task["png"])
    jpg = Path(task["jpg"])
    download(result_url, png)
    convert(png, jpg)
    page.screenshot(path=str(probe.with_suffix(".done.png")), full_page=False)
    print(task["name"], result_url)


def main() -> None:
    ensure_dirs()
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(JIMENG_WS)
        ctx = browser.contexts[0]
        page = ctx.pages[14]
        for task in TASKS:
            run_task(page, task)
        browser.close()


if __name__ == "__main__":
    main()
