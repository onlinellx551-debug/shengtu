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
        "name": "top03_onbody_round17",
        "board": ROUND17 / "01_input_boards" / "top03_onbody_board.jpg",
        "webp": RAW_DIR / "top03_onbody_round17.webp",
        "jpg": HARVEST_DIR / "jimeng_top03_onbody_round17.jpg",
        "prompt": (
            "请严格参考我上传的参考板，生成与参考镜头一一对应的成熟日系男装上身图。"
            "保留原参考图的人像构图、镜头距离、姿势和裤装，但把上衣完整替换成我们自己的重磅白色圆领短袖T恤。"
            "必须是真实棉质短袖白T，不能有针织、卫衣或长袖特征，要有厚度感、自然褶皱、真实袖口和领口缝线。"
            "整体要像成熟男装电商摄影，难以看出AI。不要文字，不要logo，不要拼贴，只要一张正方形成品图。"
        ),
    },
    {
        "name": "benefit01_lifestyle_round17",
        "board": ROUND17 / "01_input_boards" / "benefit01_lifestyle_board.jpg",
        "webp": RAW_DIR / "benefit01_lifestyle_round17.webp",
        "jpg": HARVEST_DIR / "jimeng_benefit01_lifestyle_round17.jpg",
        "prompt": (
            "请严格参考我上传的参考板，生成中段生活方式卖点图。保留参考图的坐姿、暖色室内光线、成熟男装氛围和镜头关系，"
            "但把上衣替换成我们自己的重磅白色圆领短袖T恤。白T必须厚实、真实、自然，不要AI感。"
            "不要文字，不要logo，不要拼贴，只要一张正方形成品图。"
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
        if "lv-btn-primary" not in cls or "lv-btn-disabled" in cls:
            continue
        if box["width"] < 28 or box["height"] < 28:
            continue
        if best is None or (box["y"], box["x"]) > best[0]:
            best = ((box["y"], box["x"]), btn)
    if not best:
        raise RuntimeError("Jimeng submit button not found")
    return best[1]


def visible_editor(page):
    loc = page.locator('[contenteditable="true"]')
    for i in range(loc.count()):
        ed = loc.nth(i)
        try:
            box = ed.bounding_box()
        except Exception:
            box = None
        if box and box["width"] > 180 and box["height"] > 40:
            return ed
    raise RuntimeError("Jimeng editor not found")


def upload_board(page, board: Path) -> None:
    inputs = page.locator('input[type="file"]')
    if inputs.count() == 0:
        raise RuntimeError("Jimeng file input not found")
    target = 1 if inputs.count() > 1 else 0
    inputs.nth(target).set_input_files(str(board))
    page.wait_for_timeout(3500)


def latest_result_url(page) -> str | None:
    html = page.content()
    marker = "agentic-image-item-AfDUZW"
    idx = html.rfind(marker)
    if idx == -1:
        return None
    sub = html[max(0, idx - 1200):idx + 4000]
    key = 'src="'
    pos = sub.find(key)
    found = None
    while pos != -1:
        end = sub.find('"', pos + len(key))
        src = sub[pos + len(key):end]
        if "dreamina-sign.byteimg.com" in src:
            found = src.replace("&amp;", "&")
        pos = sub.find(key, end + 1)
    return found


def wait_for_result(page, previous_url: str | None) -> str:
    end = time.time() + 420
    while time.time() < end:
        url = latest_result_url(page)
        body = page.locator("body").inner_text()
        if url and url != previous_url and "已为您生成符合要求的图片" in body:
            return url
        page.wait_for_timeout(4000)
    raise TimeoutError("Jimeng result wait timeout")


def download(url: str, dest: Path) -> None:
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0", "Referer": "https://jimeng.jianying.com/"})
    with urllib.request.urlopen(req) as resp:
        dest.write_bytes(resp.read())


def convert(src: Path, dest: Path) -> None:
    Image.open(src).convert("RGB").save(dest, quality=96)


def run_task(page, task: dict[str, Path | str]) -> None:
    page.goto("https://jimeng.jianying.com/ai-tool/canvas", wait_until="load")
    page.wait_for_timeout(3000)
    board = Path(task["board"])
    upload_board(page, board)
    editor = visible_editor(page)
    editor.click(force=True)
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")
    page.keyboard.insert_text(str(task["prompt"]))
    prev = latest_result_url(page)
    probe = PROBE_DIR / str(task["name"])
    page.screenshot(path=str(probe.with_suffix(".before.png")), full_page=False)
    visible_submit(page).click(force=True)
    page.wait_for_timeout(1200)
    page.screenshot(path=str(probe.with_suffix(".after.png")), full_page=False)
    result = wait_for_result(page, prev)
    webp = Path(task["webp"])
    jpg = Path(task["jpg"])
    download(result, webp)
    convert(webp, jpg)
    page.screenshot(path=str(probe.with_suffix(".done.png")), full_page=False)
    print(task["name"], result)


def main() -> None:
    ensure_dirs()
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(JIMENG_WS)
        page = browser.contexts[0].pages[14]
        for task in TASKS:
            run_task(page, task)
        browser.close()


if __name__ == "__main__":
    main()
