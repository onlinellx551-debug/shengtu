from __future__ import annotations

import re
import sys
import time
import urllib.request
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
ROUND10 = ROOT / "step6_output" / "素材包" / "T01_round10_2026-03-24"
ROUND11 = ROOT / "step6_output" / "素材包" / "T01_round11_2026-03-24"
BOARD_DIR = ROUND10 / "01_input_boards"
RAW_DIR = ROUND11 / "03_lovart_raw"
PROBE_DIR = ROOT / "step6_output" / "lovart_probe_new"
LOVART_WS = "ws://127.0.0.1:50781/devtools/browser/578a9329-34c2-4548-b6cb-b26bc2052c52"


TASKS = [
    {
        "name": "stack_multicolor_exact",
        "board": BOARD_DIR / "G05_stack_multicolor_board_v1.jpg",
        "outfile": RAW_DIR / "stack_multicolor_exact.png",
        "prompt": (
            "Create one ultra realistic ecommerce product still-life photo using the uploaded board "
            "as strict image-to-image guidance. Match the exact right-side composition very closely. "
            "Show our own heavyweight Japanese menswear crew neck T-shirts folded in a stacked pile on "
            "a warm beige tabletop. Keep the draped top garment, yarn cone on the left, and the same "
            "camera angle and crop. Use our actual product color range from the board: white, ivory, "
            "light gray, taupe-gray, charcoal, and black. The garments must read as our cotton jersey "
            "T-shirt product, not knit sweaters. Keep our own neck label visible on the top folded shirt. "
            "Real studio photography, believable cotton texture, natural daylight, realistic shadows, "
            "subtle imperfections, premium Japanese menswear ecommerce quality. No collage, no board, "
            "no mockup, no text, no logo, one square image only."
        ),
    },
    {
        "name": "fold_multicolor_exact",
        "board": BOARD_DIR / "G07_fold_multicolor_board_v1.jpg",
        "outfile": RAW_DIR / "fold_multicolor_exact.png",
        "prompt": (
            "Create one ultra realistic ecommerce close-up product photo using the uploaded board as strict "
            "image-to-image guidance. Match the exact right-side close-up layered fold composition very closely. "
            "Show several folded heavyweight crew neck T-shirts layered diagonally in a tight close crop on a warm "
            "beige tabletop. Use our actual product colors from the board: white, ivory, light gray, and taupe-gray. "
            "The fabric must be our cotton jersey T-shirt fabric, not sweater knitwear. Keep our own neck label visible "
            "on one exposed shirt. Real camera photograph, soft warm daylight, realistic textile grain, realistic shadow "
            "falloff, subtle natural imperfections, premium Japanese menswear ecommerce styling. No collage, no board, "
            "no mockup, no text, no logo, one square image only."
        ),
    },
]


def ensure_dirs() -> None:
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    PROBE_DIR.mkdir(parents=True, exist_ok=True)


def latest_agent_image(page) -> str | None:
    candidates: list[str] = []
    for img in page.locator("img").all():
        try:
            box = img.bounding_box()
            src = img.get_attribute("src") or ""
        except Exception:
            continue
        if not box or box["width"] < 120 or box["height"] < 120:
            continue
        if "a.lovart.ai/artifacts/agent/" in src:
            candidates.append(src)
    return candidates[-1] if candidates else None


def wait_for_finished(page, timeout_s: int = 240) -> str:
    end = time.time() + timeout_s
    last_state = ""
    while time.time() < end:
        body = page.locator("body").inner_text()
        if "生成中遇到错误" in body or "wasn't able to access the image from that URL" in body:
            raise RuntimeError("Lovart 返回生成错误")
        if any(token in body for token in ["Perfect!", "I've created", "What I delivered", "What I created"]):
            url = latest_agent_image(page)
            if url:
                return url
        if any(token in body for token in ["复用已有结果", "已生成", "已完成"]):
            url = latest_agent_image(page)
            if url:
                return url
        match = re.search(r"(\d+%)", body)
        if match:
            state = match.group(1)
            if state != last_state:
                print("progress:", state)
                last_state = state
        page.wait_for_timeout(3000)
    raise TimeoutError("Lovart 轮询超时")


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


def send_task(page, task: dict[str, object]) -> str:
    # Ensure a blank new dialogue in this page.
    new_chat = page.locator('button[data-testid="agent-new-chat-button"]')
    if new_chat.count() and new_chat.first.is_enabled():
        new_chat.first.click()
        page.wait_for_timeout(1000)

    with page.expect_file_chooser(timeout=5000) as fc:
        page.locator("button").nth(28).click(force=True)
    chooser = fc.value
    chooser.set_files(str(task["board"]))
    basename = Path(str(task["board"])).stem
    uploaded = False
    for _ in range(30):
        body = page.locator("body").inner_text()
        if basename in body and ("图片上传成功" in body or "拖拽至此处添加到对话" in body):
            uploaded = True
            break
        page.wait_for_timeout(500)
    if not uploaded:
        raise RuntimeError(f"附件未完成上传: {basename}")

    editor = page.locator('[contenteditable="true"]').first
    editor.click()
    editor.fill("")
    page.keyboard.insert_text(str(task["prompt"]))
    page.wait_for_timeout(800)

    send_button = page.locator("button").nth(34)
    for _ in range(20):
        if not send_button.is_disabled():
            break
        page.wait_for_timeout(500)
    if send_button.is_disabled():
        raise RuntimeError("发送按钮未激活")

    screenshot_base = PROBE_DIR / f"{task['name']}_before_send.png"
    page.screenshot(path=str(screenshot_base), full_page=False)
    send_button.click(force=True)
    page.wait_for_timeout(1500)
    page.screenshot(path=str(PROBE_DIR / f"{task['name']}_after_send.png"), full_page=False)
    return wait_for_finished(page)


def main() -> None:
    ensure_dirs()
    selected = set(sys.argv[1:])
    tasks = [task for task in TASKS if not selected or task["name"] in selected]
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(LOVART_WS)
        ctx = browser.contexts[0]
        for task in tasks:
            page = ctx.new_page()
            page.goto("https://www.lovart.ai/canvas", wait_until="load")
            page.wait_for_timeout(4000)
            try:
                result_url = send_task(page, task)
                print(task["name"], result_url)
                download(result_url, Path(task["outfile"]))
                page.screenshot(path=str(PROBE_DIR / f"{task['name']}_done.png"), full_page=False)
            finally:
                page.close()
        browser.close()


if __name__ == "__main__":
    main()
