from __future__ import annotations

import base64
import json
import time
import urllib.request
from pathlib import Path

from PIL import Image
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
STEP6 = ROOT / "step6_output" / "material_bundle"
ROUND20 = STEP6 / "T01_round20_2026-03-26"
BOARD_DIR = ROUND20 / "01_input_boards"
RAW_DIR = ROUND20 / "02_generated"
PROBE_DIR = ROUND20 / "05_probe_lovart"

ADSP_ACTIVE_URL = "http://127.0.0.1:50325/api/v1/browser/active?user_id=k1aph3eg"
ADSP_START_URL = "http://127.0.0.1:50325/api/v1/browser/start?user_id=k1aph3eg"

TASKS = [
    {
        "name": "top03_onbody_round20",
        "board": BOARD_DIR / "top03_onbody_board_round20.jpg",
        "png": RAW_DIR / "top03_onbody_round20.png",
        "jpg": RAW_DIR / "top03_onbody_round20.jpg",
        "prompt": (
            "Match the exact lower-left reference model composition very closely. "
            "Same adult East Asian male torso crop, same studio lighting, same hand-in-pocket posture, same premium Japanese menswear ecommerce styling. "
            "Replace the sweater with our real heavyweight white crew neck short sleeve T-shirt. It must look like a real 300g cotton T-shirt, not knitwear, not sweatshirt, not long sleeve. "
            "Keep beige trousers and neutral background. No text, no logo, no board, no collage, one clean square finished photo only."
        ),
    },
    {
        "name": "top04_onbody_round20",
        "board": BOARD_DIR / "top04_onbody_board_round20.jpg",
        "png": RAW_DIR / "top04_onbody_round20.png",
        "jpg": RAW_DIR / "top04_onbody_round20.jpg",
        "prompt": (
            "Match the exact lower-right reference model composition very closely. "
            "Same adult East Asian male torso crop, same quiet neutral wall, same premium Japanese menswear ecommerce mood. "
            "Replace the knit sweater with our real heavyweight white crew neck short sleeve T-shirt. It must read as a thick real cotton T-shirt with stable collar and believable fabric weight. "
            "Keep the beige trousers and calm mature styling. No text, no logo, no board, no collage, one clean square finished photo only."
        ),
    },
    {
        "name": "top06_structure_round20",
        "board": BOARD_DIR / "top06_structure_board_round20.jpg",
        "png": RAW_DIR / "top06_structure_round20.png",
        "jpg": RAW_DIR / "top06_structure_round20.jpg",
        "prompt": (
            "Match the exact folded edge detail composition very closely. "
            "Create one premium ecommerce still-life close-up for our heavyweight white crew neck T-shirt. Preserve the same camera angle, same shallow depth, same warm neutral scene and same folded edge placement. "
            "The garment must look like a real thick cotton jersey T-shirt, not knitwear, but keep the refined structural feeling of the reference. Show believable collar seam, fabric thickness and edge structure. "
            "No text, no watermark, no board, no collage, one clean square finished photo only."
        ),
    },
]


def ensure_dirs() -> None:
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    PROBE_DIR.mkdir(parents=True, exist_ok=True)


def get_json(url: str) -> dict:
    with urllib.request.urlopen(url) as resp:
        return json.loads(resp.read().decode("utf-8"))


def get_ws() -> str:
    payload = get_json(ADSP_ACTIVE_URL)
    if payload.get("code") == 0 and payload.get("data", {}).get("status") == "Active":
        return payload["data"]["ws"]["puppeteer"]
    payload = get_json(ADSP_START_URL)
    if payload.get("code") == 0:
        return payload["data"]["ws"]["puppeteer"]
    raise RuntimeError(f"AdsPower start failed: {payload}")


def upload_board(page, board: Path) -> None:
    attach = page.locator('[data-testid="agent-attachment-button"]').first
    try:
        with page.expect_file_chooser(timeout=6000) as chooser_info:
            attach.click(force=True)
        chooser_info.value.set_files(str(board))
    except PlaywrightTimeoutError as exc:
        raise RuntimeError("Lovart attachment chooser did not open") from exc


def wait_for_upload(page, board: Path) -> None:
    end = time.time() + 40
    while time.time() < end:
        body = page.locator("body").inner_text()
        if board.stem in body or board.name in body:
            return
        page.wait_for_timeout(600)
    raise TimeoutError(f"Lovart attachment not visible for {board.name}")


def latest_agent_image(page) -> str | None:
    urls: list[str] = []
    for img in page.locator("img").all():
        try:
            src = img.get_attribute("src") or ""
            box = img.bounding_box()
        except Exception:
            continue
        if not box or box["width"] < 120 or box["height"] < 120:
            continue
        if any(token in src for token in [
            "a.lovart.ai/artifacts/agent/",
            "a.lovart.ai/artifacts/generator/",
            "assets-persist.lovart.ai/agent_images/",
        ]):
            urls.append(src)
    return urls[-1] if urls else None


def send_button(page):
    btn = page.locator('[data-testid="agent-send-button"]')
    if btn.count():
        return btn.first
    raise RuntimeError("Lovart send button not found")


def download(url: str, dest: Path) -> None:
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0", "Referer": "https://www.lovart.ai/"})
    with urllib.request.urlopen(req) as resp:
        dest.write_bytes(resp.read())


def convert(src: Path, dest: Path) -> None:
    Image.open(src).convert("RGB").save(dest, quality=96)


def wait_for_result(page, prev_url: str | None) -> str:
    end = time.time() + 480
    while time.time() < end:
        body = page.locator("body").inner_text()
        if "wasn't able to access the image from that URL" in body or "生成中遇到错误" in body:
            raise RuntimeError("Lovart generation failed")
        url = latest_agent_image(page)
        if url and url != prev_url and any(token in body for token in ["Perfect!", "I've created", "created your", "What I delivered", "生成", "ultra-realistic"]):
            return url
        page.wait_for_timeout(3000)
    raise TimeoutError("Lovart wait timeout")


def run_task(page, task: dict[str, Path | str]) -> None:
    page.bring_to_front()
    page.wait_for_timeout(1200)
    new_chat = page.locator('[data-testid="agent-new-chat-button"]')
    if new_chat.count():
        new_chat.first.click(force=True)
        page.wait_for_timeout(1200)

    board = Path(task["board"])
    upload_board(page, board)
    wait_for_upload(page, board)

    editor = page.locator('[contenteditable="true"]').first
    editor.click(force=True)
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")
    page.keyboard.insert_text(str(task["prompt"]))
    page.wait_for_timeout(600)
    prev_url = latest_agent_image(page)

    probe = PROBE_DIR / str(task["name"])
    page.screenshot(path=str(probe.with_suffix(".before.png")), full_page=False)
    send_button(page).click(force=True)
    page.wait_for_timeout(1200)
    page.screenshot(path=str(probe.with_suffix(".after.png")), full_page=False)

    result_url = wait_for_result(page, prev_url)
    png = Path(task["png"])
    jpg = Path(task["jpg"])
    download(result_url, png)
    convert(png, jpg)
    page.screenshot(path=str(probe.with_suffix(".done.png")), full_page=False)
    print(task["name"], result_url)


def run_with_recovery(ctx, task: dict[str, Path | str]) -> None:
    attempts = 0
    last_err: Exception | None = None
    while attempts < 2:
        attempts += 1
        page = ctx.new_page()
        page.goto("https://www.lovart.ai/canvas", wait_until="load")
        page.wait_for_timeout(2500)
        try:
            run_task(page, task)
            return
        except Exception as exc:
            last_err = exc
            try:
                probe = PROBE_DIR / f"{task['name']}_recover_attempt{attempts}.png"
                page.screenshot(path=str(probe), full_page=False)
            except Exception:
                pass
            page.close()
    if last_err:
        raise last_err


def main() -> None:
    ensure_dirs()
    ws = get_ws()
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(ws)
        ctx = browser.contexts[0]
        for task in TASKS:
            run_with_recovery(ctx, task)
        browser.close()


if __name__ == "__main__":
    main()
