from __future__ import annotations

from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "step6_output" / "素材包" / "T01_round12_2026-03-25" / "watch_logs"
OUT.mkdir(parents=True, exist_ok=True)

LOVART_WS = "ws://127.0.0.1:50781/devtools/browser/578a9329-34c2-4548-b6cb-b26bc2052c52"
JIMENG_WS = "ws://127.0.0.1:54902/devtools/browser/03b0c484-60cf-482c-b89b-a60766f67580"


def dump_lovart() -> str:
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(LOVART_WS)
        lines = []
        for idx, page in enumerate(browser.contexts[0].pages):
            try:
                body = page.locator("body").inner_text()[:2000]
                lines.append(f"[lovart][page {idx}] {page.url}\n{body}\n")
            except Exception as exc:
                lines.append(f"[lovart][page {idx}] ERROR {exc}\n")
        browser.close()
    return "\n".join(lines)


def dump_jimeng() -> str:
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(JIMENG_WS)
        lines = []
        for idx, page in enumerate(browser.contexts[0].pages):
            if idx not in (9, 14):
                continue
            try:
                body = page.locator("body").inner_text()[:2400]
                lines.append(f"[jimeng][page {idx}] {page.url}\n{body}\n")
            except Exception as exc:
                lines.append(f"[jimeng][page {idx}] ERROR {exc}\n")
        browser.close()
    return "\n".join(lines)


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    try:
        lovart = dump_lovart()
    except Exception as exc:
        lovart = f"[lovart] ERROR {exc}\n"
    try:
        jimeng = dump_jimeng()
    except Exception as exc:
        jimeng = f"[jimeng] ERROR {exc}\n"
    text = f"timestamp={stamp}\n\n{lovart}\n\n{jimeng}\n"
    out = OUT / f"session_status_{stamp}.txt"
    out.write_text(text, encoding="utf-8")
    print(out)


if __name__ == "__main__":
    main()
