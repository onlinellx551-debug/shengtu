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
ROUND17 = STEP6 / "T01_round17_2026-03-25"
RAW_DIR = ROUND17 / "03_lovart_raw"
HARVEST_DIR = ROUND17 / "04_harvested"
PROBE_DIR = ROUND17 / "05_probe_lovart"

ADSP_ACTIVE_URL = "http://127.0.0.1:50325/api/v1/browser/active?user_id=k1aph3eg"
ADSP_START_URL = "http://127.0.0.1:50325/api/v1/browser/start?user_id=k1aph3eg"

TASKS = [
    {
        "name": "top01_onbody_round17",
        "board": ROUND17 / "01_input_boards" / "top01_onbody_board.jpg",
        "png": RAW_DIR / "top01_onbody_round17.png",
        "jpg": HARVEST_DIR / "lovart_top01_onbody_round17.jpg",
        "prompt": (
            "Use the uploaded board as strict image-to-image guidance. Match the exact composition of the reference model shot very closely. "
            "Keep the same adult East Asian male, same studio lighting, same crop, same pose, same hand-in-pocket posture, same premium Japanese menswear ecommerce feel. "
            "Replace the sweater completely with our own heavyweight white crew neck short sleeve T-shirt. It must clearly be a real 300g cotton T-shirt, not knitwear, not a sweatshirt, "
            "not a long sleeve garment. Keep the beige trousers and neutral background, preserve the calm mature menswear styling, realistic body fit, natural cotton folds, slight weight in the fabric, "
            "subtle collar seam, impossible to notice AI. No text, no logo watermark, no mockup, no board, one square finished photo only."
        ),
    },
    {
        "name": "top03_onbody_round17",
        "board": ROUND17 / "01_input_boards" / "top03_onbody_board.jpg",
        "png": RAW_DIR / "top03_onbody_round17.png",
        "jpg": HARVEST_DIR / "lovart_top03_onbody_round17.jpg",
        "prompt": (
            "Use the uploaded board as strict image-to-image guidance. Match the exact composition of the second reference model shot very closely. "
            "Same adult East Asian male, same studio crop, same neutral lighting, same relaxed premium Japanese menswear ecommerce styling. "
            "Replace the sweater completely with our own heavyweight white crew neck short sleeve T-shirt. The garment must read as a real cotton short sleeve T-shirt with a stable collar, "
            "natural drape, believable sleeve opening and subtle fabric weight. Keep the trousers and body posture close to the reference. Impossible to notice AI. "
            "No text, no logo watermark, no mockup, no collage, no board, one square finished photo only."
        ),
    },
    {
        "name": "benefit01_lifestyle_round17",
        "board": ROUND17 / "01_input_boards" / "benefit01_lifestyle_board.jpg",
        "png": RAW_DIR / "benefit01_lifestyle_round17.png",
        "jpg": HARVEST_DIR / "lovart_benefit01_lifestyle_round17.jpg",
        "prompt": (
            "Use the uploaded board as strict image-to-image guidance. Match the exact lifestyle composition very closely: same seated adult East Asian male, same warm interior lighting, "
            "same angle, same calm premium Japanese menswear storytelling style. Replace the sweater completely with our own heavyweight white crew neck short sleeve T-shirt. "
            "The white T-shirt must look real, thick enough, not transparent, natural soft cotton texture, mature understated styling, impossible to notice AI. "
            "Keep the book, chair or ledge relationship, and preserve the warm editorial atmosphere. No text, no logo watermark, no mockup, no board, one square finished photo only."
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


def wait_for_uploaded(page) -> None:
    end = time.time() + 30
    while time.time() < end:
        body = page.locator("body").inner_text()
        if "图片上传成功" in body or "image uploaded" in body.lower():
            return
        page.wait_for_timeout(500)
    raise TimeoutError("Lovart image paste did not complete")


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


def find_send_button(page):
    btn = page.locator('[data-testid="agent-send-button"]')
    if btn.count() and not btn.first.is_disabled():
        return btn.first
    raise RuntimeError("Lovart send button not found")


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
        headers={"User-Agent": "Mozilla/5.0", "Referer": "https://www.lovart.ai/"},
    )
    with urllib.request.urlopen(request) as resp:
        dest.write_bytes(resp.read())


def convert_to_jpg(src: Path, dest: Path) -> None:
    img = Image.open(src).convert("RGB")
    img.save(dest, quality=96)


def run_task(page, task: dict[str, Path | str]) -> None:
    page.bring_to_front()
    page.wait_for_timeout(1200)
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
        page = ctx.pages[1]
        for task in TASKS:
            run_task(page, task)
        browser.close()


if __name__ == "__main__":
    main()
