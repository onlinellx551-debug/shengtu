from __future__ import annotations

import argparse
import json
import time
import urllib.request
from pathlib import Path

from PIL import Image, ImageDraw, ImageOps
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
STEP6 = ROOT / "step6_output" / "material_bundle"
ROUND20 = STEP6 / "T01_round20_2026-03-26"
URLS_JSON = ROUND20 / "05_probe_lovart" / "round20_exact_tmpfiles_urls.json"
HARVEST_ROOT = ROUND20 / "05_probe_lovart" / "round21_cleanflow"
USER_ID = "k1at9ows"
ADSP_BASE = "http://127.0.0.1:50325/api/v1/browser"


TASKS = {
    "top02": {
        "label": "top02_onbody",
        "url_key": "top02_exact_board_round20.jpg",
        "wait_seconds": 220,
        "prompt": (
            "Create one ultra realistic premium Japanese menswear ecommerce photo. "
            "The reference board URL contains the exact target shot and our real product details. "
            "Follow the exact target composition almost one-to-one: same adult East Asian male, same crop from below the nose to the hips, "
            "same cooler neutral studio wall with the left vertical wall split visible, same left hand in pocket, same right arm relaxed down, "
            "same centered torso balance, same beige trousers, same restrained premium mood. "
            "Replace the knitwear with our real heavyweight WHITE crew neck short sleeve T-shirt from the board. "
            "It must read as a real 300g cotton tee: short sleeves, cotton jersey, subtle natural wrinkles, believable hem and sleeve folds. "
            "Do not invent long sleeves, knit texture, sweatshirt texture, hoodie structure, text, collage or extra objects. "
            "Return one clean square image only."
        ),
    },
    "top03": {
        "label": "top03_onbody",
        "url_key": "top03_exact_board_round20.jpg",
        "wait_seconds": 220,
        "prompt": (
            "Create one ultra realistic premium Japanese menswear ecommerce photo. "
            "The reference board URL contains the exact target shot and our real product details. "
            "Follow the exact target composition almost one-to-one: same adult East Asian male, same warmer wall tone, same slight body turn, "
            "same left hand in pocket, same right hand hanging and visible, same torso crop, same beige trousers, same softer restrained mood. "
            "Replace the knitwear with our real heavyweight IVORY or LIGHT BEIGE crew neck short sleeve T-shirt from the board. "
            "The shirt color must stay consistent with the flat product on the board. "
            "It must look like a real 300g cotton tee, not knitwear and not a sweatshirt. "
            "Keep the crop clean and realistic, with subtle cotton folds and a believable collar. "
            "No text, no collage, no board, no extra objects. One clean square image only."
        ),
    },
    "top04": {
        "label": "top04_onbody",
        "url_key": "top04_exact_board_round20.jpg",
        "wait_seconds": 220,
        "prompt": (
            "Create one ultra realistic premium Japanese menswear ecommerce photo. "
            "The reference board URL contains the exact target shot and our real product details. "
            "Follow the exact target composition almost one-to-one: same adult East Asian male, same darker textured wall, same straighter front-facing torso, "
            "same firmer silhouette, same beige trousers, same calm premium mood. "
            "Replace the knitwear with our real heavyweight HEATHER GRAY crew neck short sleeve T-shirt from the board. "
            "The shirt color must match the gray flat product on the board. "
            "The shirt must clearly read as a thick cotton tee with short sleeves, a real hem, and realistic shoulder folds. "
            "No knitwear texture, no sweater collar, no hoodie, no extra props, no text. "
            "Return one clean square image only."
        ),
    },
    "top05": {
        "label": "top05_stack",
        "url_key": "top05_exact_board_round20.jpg",
        "wait_seconds": 220,
        "prompt": (
            "Create one ultra realistic premium Japanese menswear ecommerce still-life photo. "
            "The reference board URL contains the exact target shot and our real product details. "
            "Follow the exact target composition almost one-to-one: same warm beige tabletop, same spool position on the left, same tidy stacked pile, "
            "same calm premium mood, same square crop. "
            "But the garments must be our real heavyweight short sleeve T-shirts, not knitwear and not sweatshirts. "
            "Show a folded stack of short sleeve T-shirts in our real color range: white, ivory, gray, black. "
            "The top shirt must be clearly draped over the stack the way the reference does it: one soft top layer folded on top, "
            "with a natural loose edge cascading down the left side of the stack. "
            "Do not make the top shirt flat, stiff or fully folded. "
            "Do not show a long sleeve garment. "
            "Keep the fabric realistic, weighty and photographic. No text, no collage, no extra props beyond the spool. One clean square image only."
        ),
    },
    "top06": {
        "label": "top06_structure",
        "url_key": "top06_exact_board_round20.jpg",
        "wait_seconds": 220,
        "prompt": (
            "Create one ultra realistic premium ecommerce still-life close-up. "
            "The reference board URL contains the exact target shot and our real product details. "
            "Follow the exact composition almost one-to-one: same folded cuff-like edge facing the camera, same close crop, same warm soft lighting, "
            "same quiet Japanese menswear product-photo mood. "
            "Translate the knit reference into our real heavyweight white crew neck short sleeve T-shirt fabric logic. "
            "The main subject must be the folded short sleeve hem / cuff area and the edge thickness, not the collar. "
            "Keep the collar mostly out of frame or secondary. "
            "It must show believable thick cotton jersey, refined folded structure, realistic edge thickness and fine cotton texture. "
            "Do not generate knit rib body fabric, sweater folds, sweatshirt texture or a flat label shot. "
            "No text, no collage, no board, no extra props. One clean square image only."
        ),
    },
    "top07": {
        "label": "top07_multicolor_fold",
        "url_key": "top07_exact_board_round20.jpg",
        "wait_seconds": 220,
        "prompt": (
            "Create one ultra realistic premium ecommerce still-life close-up. "
            "The reference board URL contains the exact target shot and our real product details. "
            "Follow the exact composition almost one-to-one: same folded multicolor stack, same warm tabletop, same crop and same quiet premium Japanese menswear mood. "
            "Use our real heavyweight short sleeve T-shirt logic, not knitwear. "
            "Show believable folded cotton T-shirts with realistic crew-neck collars, natural minimal labels and subtle color differences only from our real product range: white, ivory, gray. "
            "Keep the stack orderly and premium, not messy. No text, no collage, no extra props. One clean square image only."
        ),
    },
    "top08": {
        "label": "top08_collar",
        "url_key": "top08_exact_board_round20.jpg",
        "wait_seconds": 220,
        "prompt": (
            "Create one ultra realistic premium ecommerce detail close-up. "
            "The reference board URL contains the exact target shot and our real product details. "
            "Follow the exact composition almost one-to-one: same tight collar crop, same neck-to-collar relationship, same premium studio light and same quiet Japanese menswear mood. "
            "Translate the knit reference into our real heavyweight white crew neck short sleeve T-shirt collar. "
            "The image must focus on collar structure, ribbing, seam quality and cotton texture. "
            "The collar should look premium and substantial, not thin or cheap, with a clean neck crop and refined shoulder slope. "
            "Do not make it a flat label shot. No text, no collage, no extra props. One clean square image only."
        ),
    },
}


def http_json(url: str) -> dict:
    with urllib.request.urlopen(url, timeout=20) as resp:
        return json.loads(resp.read().decode("utf-8", "ignore"))


def restart_env() -> str:
    try:
        active = http_json(f"{ADSP_BASE}/active?user_id={USER_ID}")
        if active.get("code") == 0 and active.get("data", {}).get("status") == "Active":
            return active["data"]["ws"]["puppeteer"]
    except Exception:
        pass

    last_payload = None
    for _ in range(5):
        payload = http_json(f"{ADSP_BASE}/start?user_id={USER_ID}")
        last_payload = payload
        if payload.get("code") == 0:
            return payload["data"]["ws"]["puppeteer"]
        time.sleep(3)
    raise RuntimeError(f"AdsPower start failed: {last_payload}")


def asset_urls(page) -> list[dict]:
    items: list[dict] = []
    for i in range(page.locator("img").count()):
        img = page.locator("img").nth(i)
        try:
            src = img.get_attribute("src") or ""
            box = img.bounding_box()
        except Exception:
            continue
        if not box:
            continue
        if not any(
            token in src
            for token in [
                "a.lovart.ai/artifacts/agent/",
                "a.lovart.ai/artifacts/generator/",
                "assets-persist.lovart.ai/agent_images/",
            ]
        ):
            continue
        items.append(
            {
                "src": src.split("?")[0],
                "x": box["x"],
                "y": box["y"],
                "width": box["width"],
                "height": box["height"],
                "area": box["width"] * box["height"],
            }
        )
    return items


def visible_canvas_candidates(page) -> list[dict]:
    items = []
    for item in asset_urls(page):
        if item["x"] < 1700 and item["width"] >= 180 and item["height"] >= 180:
            items.append(item)
    items.sort(key=lambda x: x["area"], reverse=True)
    seen = set()
    unique = []
    for item in items:
        if item["src"] in seen:
            continue
        seen.add(item["src"])
        unique.append(item)
    return unique


def click_new_project(page) -> None:
    coords = page.evaluate(
        """() => {
          const exact = ['新项目', '新建项目', '新建對話', '新对话', 'New project', 'New Project', 'New chat', 'New Chat'];
          for (const el of document.querySelectorAll('div,button,a,span')) {
            const t = (el.innerText || '').trim();
            if (exact.includes(t)) {
              const r = el.getBoundingClientRect();
              return { x: r.left + r.width / 2, y: r.top + r.height / 2 };
            }
          }
          return null;
        }"""
    )
    if coords:
        page.mouse.click(coords["x"], coords["y"])
    else:
        page.mouse.click(240, 760)


def ensure_agent_mode(page) -> None:
    try:
        if page.locator("text=Agent").count():
            page.locator("text=Agent").last.click(force=True)
            page.wait_for_timeout(600)
            return
    except Exception:
        pass
    try:
        page.mouse.click(1290, 1246)
        page.wait_for_timeout(600)
    except Exception:
        pass


def ensure_nano_banana_pro(page) -> None:
    try:
        # Open the model preference popup near the lower-right composer controls.
        page.mouse.click(1176, 1244)
        page.wait_for_timeout(1200)
        # Disable auto model routing if it is on.
        auto_on = page.evaluate(
            """() => {
              const track = document.querySelector('.mantine-Switch-track');
              const thumb = document.querySelector('.mantine-Switch-thumb');
              if (!track || !thumb) return false;
              const tr = track.getBoundingClientRect();
              const th = thumb.getBoundingClientRect();
              return th.left > tr.left + 10;
            }"""
        )
        if auto_on:
            page.mouse.click(1156, 846)
            page.wait_for_timeout(800)
        if page.locator("text=Nano Banana Pro").count():
            page.locator("text=Nano Banana Pro").first.click(force=True)
            page.wait_for_timeout(1200)
        # Close the popup so subsequent input is reliable.
        page.mouse.click(640, 640)
        page.wait_for_timeout(500)
    except Exception:
        pass


def ensure_reasoning_mode(page) -> None:
    # Best effort: user prefers the lightbulb "thinking" mode in the composer.
    try:
        page.mouse.click(1415, 1245)
        page.wait_for_timeout(500)
    except Exception:
        pass


def open_clean_project(ctx):
    for pg in list(ctx.pages):
        try:
            if "lovart.ai" in pg.url:
                pg.close()
        except Exception:
            pass

    page = ctx.new_page()
    page.goto("https://www.lovart.ai/zh-TW/home", wait_until="domcontentloaded")
    page.wait_for_timeout(4500)
    click_new_project(page)
    end = time.time() + 35
    while time.time() < end:
        for candidate in reversed(ctx.pages):
            try:
                url = candidate.url
            except Exception:
                continue
            if "/canvas?projectId=" not in url:
                continue
            try:
                if candidate.locator('[data-testid="agent-message-input"]').count():
                    candidate.wait_for_timeout(1500)
                    return candidate
            except Exception:
                continue
        page.wait_for_timeout(500)
    try:
        page.goto("https://www.lovart.ai/canvas", wait_until="domcontentloaded")
        page.wait_for_timeout(4000)
        if page.locator('[data-testid="agent-message-input"]').count():
            return page
    except Exception:
        pass
    raise RuntimeError(f"Lovart clean project not ready: {page.url}")


def wait_for_generation(page, before: set[str], wait_seconds: int) -> tuple[list[str], list[dict]]:
    end = time.time() + wait_seconds
    response_hits: list[str] = []

    def on_resp(resp):
        url = resp.url.split("?")[0]
        if any(
            token in url
            for token in [
                "a.lovart.ai/artifacts/agent/",
                "a.lovart.ai/artifacts/generator/",
                "assets-persist.lovart.ai/agent_images/",
            ]
        ):
            response_hits.append(url)

    page.on("response", on_resp)
    last_change = time.time()
    best_visible: list[dict] = []
    last_count = len(before)

    while time.time() < end:
        visible = [item for item in visible_canvas_candidates(page) if item["src"] not in before]
        if visible:
            best_visible = visible
            last_change = time.time()

        current_count = len({item["src"] for item in asset_urls(page)})
        if current_count != last_count:
            last_count = current_count
            last_change = time.time()

        if best_visible and time.time() - last_change > 10:
            break
        page.wait_for_timeout(2500)

    uniq_hits = []
    seen = set()
    for url in response_hits:
        if url not in seen:
            seen.add(url)
            uniq_hits.append(url)
    return uniq_hits, best_visible


def download(url: str, dest: Path) -> None:
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0",
            "Referer": "https://www.lovart.ai/",
        },
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        dest.write_bytes(resp.read())


def build_contact_sheet(paths: list[Path], dest: Path) -> None:
    if not paths:
        return
    cards = []
    for idx, path in enumerate(paths, start=1):
        im = Image.open(path).convert("RGB")
        small = ImageOps.contain(im, (260, 260))
        card = Image.new("RGB", (280, 300), "white")
        card.paste(small, ((280 - small.width) // 2, 10))
        draw = ImageDraw.Draw(card)
        draw.text((10, 272), f"{idx:02d}", fill="black")
        cards.append(card)

    cols = 3
    rows = (len(cards) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * 280, rows * 300), (245, 245, 245))
    for idx, card in enumerate(cards):
        x = (idx % cols) * 280
        y = (idx // cols) * 300
        sheet.paste(card, (x, y))
    sheet.save(dest, quality=92)


def run_task(task_key: str) -> None:
    config = TASKS[task_key]
    url_map = json.loads(URLS_JSON.read_text(encoding="utf-8"))
    ref_url = url_map[config["url_key"]]["direct_url"]
    prompt = (
        f"Use this exact reference board URL first: {ref_url}\n\n"
        f"{config['prompt']}\n\n"
        "The result must look like a real studio product photograph, not AI art."
    )

    out_dir = HARVEST_ROOT / config["label"]
    out_dir.mkdir(parents=True, exist_ok=True)

    ws = restart_env()
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(ws)
        ctx = browser.contexts[0]
        page = open_clean_project(ctx)
        ensure_agent_mode(page)
        ensure_reasoning_mode(page)
        ensure_nano_banana_pro(page)

        before = {item["src"] for item in asset_urls(page)}
        editor = page.locator('[data-testid="agent-message-input"]').first
        editor.click(force=True)
        page.keyboard.insert_text(prompt)
        page.wait_for_timeout(800)
        page.locator('[data-testid="agent-send-button"]').first.click(force=True)

        hits, visible = wait_for_generation(page, before, config["wait_seconds"])
        page.screenshot(path=str(out_dir / "probe.png"), full_page=False, timeout=60000)
        body = page.locator("body").inner_text()
        (out_dir / "body_head.txt").write_text(body[:4000], encoding="utf-8")

        candidates: list[str] = []
        for item in visible:
            if item["src"] not in candidates:
                candidates.append(item["src"])
        for url in hits:
            if url not in candidates:
                candidates.append(url)

        saved: list[Path] = []
        for idx, url in enumerate(candidates[:12], start=1):
            ext = ".png" if ".png" in url else ".jpg"
            raw = out_dir / f"cand_{idx:02d}{ext}"
            download(url, raw)
            if raw.suffix.lower() == ".png":
                jpg = out_dir / f"cand_{idx:02d}.jpg"
                Image.open(raw).convert("RGB").save(jpg, quality=96)
                saved.append(jpg)
            else:
                saved.append(raw)

        build_contact_sheet(saved, out_dir / "contact_sheet.jpg")
        (out_dir / "urls.json").write_text(json.dumps(candidates, ensure_ascii=False, indent=2), encoding="utf-8")
        browser.close()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("task", choices=sorted(TASKS.keys()))
    args = parser.parse_args()
    run_task(args.task)


if __name__ == "__main__":
    main()
