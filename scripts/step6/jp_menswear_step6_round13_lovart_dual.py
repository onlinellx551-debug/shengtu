from __future__ import annotations

import base64
import json
import time
import urllib.request
from pathlib import Path

from PIL import Image
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
STEP6 = ROOT / "step6_output" / "material_bundle"
ROUND13 = STEP6 / "T01_round13_2026-03-25"
RAW_DIR = ROUND13 / "03_lovart_raw"
HARVEST_DIR = ROUND13 / "04_harvested"
PROBE_DIR = ROUND13 / "05_probe_lovart"

ADSP_ACTIVE_URL = "http://127.0.0.1:50325/api/v1/browser/active?user_id=k1aph3eg"
ADSP_START_URL = "http://127.0.0.1:50325/api/v1/browser/start?user_id=k1aph3eg"

TASKS = [
    {
        "page_index": 2,
        "name": "stack_multicolor_round13_v4",
        "board": ROUND13 / "01_input_boards_v2" / "G05_stack_multicolor_board_v2.jpg",
        "png": RAW_DIR / "stack_multicolor_round13_v4.png",
        "jpg": HARVEST_DIR / "lovart_stack_multicolor_round13_v4.jpg",
        "prompt": (
            "Use the uploaded board as strict image-to-image guidance. Match the exact right-side still-life composition "
            "very closely, but the garments must clearly be our own heavyweight crew neck short sleeve T-shirts. Never "
            "generate sweaters, knitwear, ribbed knit cuffs, or long sleeves. Keep the yarn cone on the left, same warm "
            "beige tabletop, same camera angle, same crop, same calm premium Japanese menswear ecommerce styling. Show a "
            "neat stacked pile using only our real colors from the board: white, ivory, light gray, taupe-gray, charcoal, "
            "and black. The black T-shirt must stay only as a bottom layer accent, never as a dominant foreground shirt. "
            "Only one top shirt may drape gently across the stack, and it must still clearly read as a short sleeve T-shirt "
            "with real T-shirt hem and sleeve behavior. Keep our own neck label naturally visible on the exposed top shirt. "
            "Real cotton jersey texture, realistic short sleeves, realistic crew neck collar seam, believable folds, soft "
            "natural daylight, subtle imperfections, premium product photography, impossible to notice AI. No text, no logo, "
            "no mockup, no collage, no board, one square finished photo only."
        ),
    },
    {
        "page_index": 3,
        "name": "fold_multicolor_round13_v4",
        "board": ROUND13 / "01_input_boards_v2" / "G07_fold_multicolor_board_v2.jpg",
        "png": RAW_DIR / "fold_multicolor_round13_v4.png",
        "jpg": HARVEST_DIR / "lovart_fold_multicolor_round13_v4.jpg",
        "prompt": (
            "Use the uploaded board as strict image-to-image guidance. Match the exact right-side close-up fold composition "
            "very closely, but make the garments unmistakably our heavyweight crew neck short sleeve T-shirts. Never produce "
            "sweaters, knitwear, or ribbed knit cuffs. Show tightly layered T-shirts in the same crop rhythm and beige "
            "tabletop setup as the reference. Use only our real colors from the board: taupe-gray, white, ivory, and light "
            "gray. Keep our own neck label naturally visible and readable on one exposed shirt. The image must show real "
            "cotton jersey T-shirt texture, real crew neck collar seam, realistic folded short-sleeve garments, soft warm "
            "daylight, believable shadows, subtle imperfections, premium Japanese menswear ecommerce photography, impossible "
            "to notice AI. No text, no logo, no mockup, no collage, no board, one square finished photo only."
        ),
    },
]


def ensure_dirs() -> None:
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    HARVEST_DIR.mkdir(parents=True, exist_ok=True)
    PROBE_DIR.mkdir(parents=True, exist_ok=True)


def get_json(url: str) -> dict:
    with urllib.request.urlopen(url) as resp:
        return json.loads(resp.read().decode("utf-8"))


def get_lovart_ws() -> str:
    payload = get_json(ADSP_ACTIVE_URL)
    if payload.get("code") == 0 and payload.get("data", {}).get("status") == "Active":
        return payload["data"]["ws"]["puppeteer"]
    payload = get_json(ADSP_START_URL)
    if payload.get("code") == 0:
        return payload["data"]["ws"]["puppeteer"]
    raise RuntimeError(f"AdsPower start failed: {payload}")


def paste_board(page, board: Path) -> None:
    payload = base64.b64encode(board.read_bytes()).decode("ascii")
    editor = page.locator('[data-testid="agent-message-input"]').first
    editor.click(force=True)
    page.evaluate(
        """async ({selector, name, b64}) => {
          const el = document.querySelector(selector);
          const dt = new DataTransfer();
          const res = await fetch(`data:image/jpeg;base64,${b64}`);
          const blob = await res.blob();
          const file = new File([blob], name, { type: 'image/jpeg' });
          dt.items.add(file);
          const evt = new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true });
          el.dispatchEvent(evt);
        }""",
        {"selector": '[data-testid="agent-message-input"]', "name": board.name, "b64": payload},
    )


def find_send_button(page):
    btn = page.locator('[data-testid="agent-send-button"]')
    if btn.count() and not btn.first.is_disabled():
        return btn.first
    raise RuntimeError("Lovart send button not found")


def latest_agent_image(page) -> str | None:
    candidates: list[str] = []
    for img in page.locator("img").all():
        try:
            src = img.get_attribute("src") or ""
            box = img.bounding_box()
        except Exception:
            continue
        if not box or box["width"] < 120 or box["height"] < 120:
            continue
        if "a.lovart.ai/artifacts/agent/" in src:
            candidates.append(src)
    return candidates[-1] if candidates else None


def wait_for_uploaded(page) -> None:
    end = time.time() + 30
    while time.time() < end:
        body = page.locator("body").inner_text()
        if "图片上传成功" in body:
            return
        page.wait_for_timeout(500)
    raise TimeoutError("Lovart image paste did not complete")


def wait_for_finished(page, previous_url: str | None) -> str:
    end = time.time() + 420
    while time.time() < end:
        body = page.locator("body").inner_text()
        if "生成中遇到错误" in body or "wasn't able to access the image from that URL" in body:
            raise RuntimeError("Lovart generation failed")
        url = latest_agent_image(page)
        if url and url != previous_url and any(
            token in body for token in ["Perfect!", "I've created", "What I delivered", "What I created", "已生成"]
        ):
            return url
        page.wait_for_timeout(3000)
    raise TimeoutError("Lovart wait timeout")


def download(url: str, dest: Path) -> None:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0",
            "Referer": "https://www.lovart.ai/",
        },
    )
    with urllib.request.urlopen(request) as resp:
        dest.write_bytes(resp.read())


def convert_to_jpg(src: Path, dest: Path) -> None:
    img = Image.open(src).convert("RGB")
    img.save(dest, quality=96)


def run_task(page, task: dict[str, Path | str]) -> None:
    page.bring_to_front()
    page.wait_for_timeout(1500)
    new_chat = page.locator('[data-testid="agent-new-chat-button"]')
    if new_chat.count():
        new_chat.first.click(force=True)
        page.wait_for_timeout(1000)

    board = Path(task["board"])
    paste_board(page, board)
    wait_for_uploaded(page)

    editor = page.locator('[contenteditable="true"]').first
    editor.click(force=True)
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")
    page.keyboard.insert_text(str(task["prompt"]))
    page.wait_for_timeout(600)
    previous_url = latest_agent_image(page)

    probe_base = PROBE_DIR / str(task["name"])
    page.screenshot(path=str(probe_base.with_suffix(".dropped.png")), full_page=False)
    page.screenshot(path=str(probe_base.with_suffix(".before.png")), full_page=False)
    find_send_button(page).click(force=True)
    page.wait_for_timeout(1200)
    page.screenshot(path=str(probe_base.with_suffix(".after.png")), full_page=False)

    result_url = wait_for_finished(page, previous_url)
    png_path = Path(task["png"])
    download(result_url, png_path)
    convert_to_jpg(png_path, Path(task["jpg"]))
    page.screenshot(path=str(probe_base.with_suffix(".done.png")), full_page=False)
    print(task["name"], result_url)


def main() -> None:
    ensure_dirs()
    ws = get_lovart_ws()
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(ws)
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
