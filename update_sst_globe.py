"""
update_sst_globe.py
-------------------
Downloads the latest SST Anomaly orthographic globe image from
Climate Reanalyzer and saves it to this folder as `map.png`.

That's it. No Excel, no PDF, no extras.

Strategy
========
1. Try a direct HTTPS download of the known PNG URL with a browser
   User-Agent + Referer (fast, no browser required).
2. If that fails, fall back to Playwright: load the page, read the
   <img id="img_world1"> src that the page's JavaScript injects, then
   download that URL inside the browser context. Last resort: screenshot
   the element.

Install
=======
    pip install requests pillow
    # Only needed if the direct download ever stops working:
    pip install playwright
    python -m playwright install chromium

Run
===
    python update_sst_globe.py
"""

from __future__ import annotations

import shutil
import sys
import traceback
from pathlib import Path

# ---------------------------------------------------------------------------
# CONFIG -- edit if needed.
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent

# Output file (lives in the same folder as this script).
OUTPUT_PNG = SCRIPT_DIR / "map.png"

# Climate Reanalyzer source.
PAGE_URL = "https://climatereanalyzer.org/wx/todays-weather/?var_id=sstanom&ortho=6&wt=1"

# The page renders TWO image sections:
#   * Top: 8 orthographic globe views (img_sat1..img_sat8). The URL
#     parameter ortho=N selects which one is shown. ortho=6 -> img_sat6
#     -> gfs_pacific-sat_*.png  (Pacific-centered globe, what we want).
#   * Bottom: flat world maps (img_world1/img_world2). NOT what we want.
#
# Map of ortho -> filename token (from tw_v11.min.js):
#   1 nh-sat1   2 samer-sat   3 euroafr-sat   4 asia-sat
#   5 ausnz-sat 6 pacific-sat 7 spole-sat     8 npole-sat
DIRECT_PNG_URL = (
    "https://climatereanalyzer.org/wx/todays-weather/maps/"
    "gfs_pacific-sat_sstanom_d1.png"
)
# DOM id of the corresponding <img> on the page (used by the Playwright
# fallback). ortho=6 -> #img_sat6.
TARGET_IMG_ID = "img_sat6"

# A modern desktop UA -- the CDN 403's bare curl/python defaults.
BROWSER_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

REQUEST_TIMEOUT = 30          # seconds, for direct HTTP
PLAYWRIGHT_TIMEOUT_MS = 45000


# ---------------------------------------------------------------------------
# Path 1 -- Direct HTTPS download with proper headers.
# ---------------------------------------------------------------------------
def download_direct(url: str, dest: Path) -> bool:
    try:
        import requests
    except ImportError:
        print("[direct] 'requests' not installed; skipping direct download.")
        return False

    headers = {
        "User-Agent": BROWSER_UA,
        "Referer": PAGE_URL,
        "Accept": "image/avif,image/webp,image/png,image/*,*/*;q=0.8",
    }
    try:
        print(f"[direct] GET {url}")
        r = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT, stream=True)
        if r.status_code != 200:
            print(f"[direct] HTTP {r.status_code} -- falling back to Playwright.")
            return False
        ctype = r.headers.get("Content-Type", "")
        if "image" not in ctype:
            print(f"[direct] Unexpected Content-Type {ctype!r} -- falling back.")
            return False
        with open(dest, "wb") as f:
            for chunk in r.iter_content(8192):
                f.write(chunk)
        if dest.stat().st_size < 10_000:
            print(f"[direct] File suspiciously small ({dest.stat().st_size} bytes).")
            return False
        print(f"[direct] OK ({dest.stat().st_size:,} bytes)")
        return True
    except Exception as e:
        print(f"[direct] Error: {e}")
        return False


# ---------------------------------------------------------------------------
# Path 2 -- Playwright fallback (real browser).
# ---------------------------------------------------------------------------
def download_via_playwright(page_url: str, dest: Path) -> bool:
    try:
        from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
    except ImportError:
        print("[playwright] Not installed. Run: pip install playwright && "
              "python -m playwright install chromium")
        return False

    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            ctx = browser.new_context(user_agent=BROWSER_UA,
                                       viewport={"width": 1280, "height": 1024})
            page = ctx.new_page()
            print(f"[playwright] Loading {page_url}")
            page.goto(page_url, timeout=PLAYWRIGHT_TIMEOUT_MS, wait_until="domcontentloaded")

            # The page's JS sets imgs' .src after window 'load'. Wait for the
            # specific orthographic globe (#img_sat6 by default) to finish loading.
            try:
                page.wait_for_function(
                    f"""() => {{
                        const el = document.getElementById('{TARGET_IMG_ID}');
                        return el && el.src && el.src.indexOf('sstanom') !== -1
                               && el.complete && el.naturalWidth > 0;
                    }}""",
                    timeout=PLAYWRIGHT_TIMEOUT_MS,
                )
            except PWTimeout:
                print(f"[playwright] Timed out waiting for #{TARGET_IMG_ID} to load.")

            img_src = page.evaluate(
                f"() => {{ const e = document.getElementById('{TARGET_IMG_ID}'); "
                "return e ? e.src : null; }"
            )
            print(f"[playwright] {TARGET_IMG_ID}.src = {img_src}")

            # 1) Try downloading the resolved URL via the browser context.
            if img_src:
                try:
                    resp = ctx.request.get(img_src, headers={"Referer": page_url})
                    if resp.ok and "image" in resp.headers.get("content-type", ""):
                        dest.write_bytes(resp.body())
                        if dest.stat().st_size > 10_000:
                            print(f"[playwright] Downloaded via ctx.request "
                                  f"({dest.stat().st_size:,} bytes)")
                            browser.close()
                            return True
                except Exception as e:
                    print(f"[playwright] ctx.request.get failed: {e}")

            # 2) Last resort: screenshot the element.
            try:
                locator = page.locator(f"#{TARGET_IMG_ID}")
                locator.scroll_into_view_if_needed(timeout=5000)
                locator.screenshot(path=str(dest))
                print(f"[playwright] Element screenshot saved.")
                browser.close()
                return dest.exists() and dest.stat().st_size > 5_000
            except Exception as e:
                print(f"[playwright] Element screenshot failed: {e}")
                browser.close()
                return False
    except Exception as e:
        print(f"[playwright] Fatal: {e}")
        traceback.print_exc()
        return False


# ---------------------------------------------------------------------------
# Orchestrator.
# ---------------------------------------------------------------------------
def acquire_image(dest: Path) -> None:
    """Try direct, then Playwright. Atomic write via .part file."""
    tmp = dest.with_suffix(dest.suffix + ".part")
    if tmp.exists():
        tmp.unlink()

    ok = download_direct(DIRECT_PNG_URL, tmp) or download_via_playwright(PAGE_URL, tmp)
    if not ok:
        tmp.unlink(missing_ok=True)
        raise RuntimeError(
            "Could not download the SST anomaly globe image via either the "
            "direct URL or Playwright. Check your internet connection and "
            "that the Climate Reanalyzer page is reachable."
        )

    # Validate it's a real image.
    try:
        from PIL import Image
        with Image.open(tmp) as im:
            im.verify()
    except ImportError:
        # Pillow optional -- skip verification if not installed.
        pass
    except Exception as e:
        tmp.unlink(missing_ok=True)
        raise RuntimeError(f"Downloaded file is not a valid image: {e}")

    # Atomic move into place (replaces any prior map.png).
    shutil.move(str(tmp), str(dest))


def main() -> int:
    print("=" * 60)
    print("Climate Reanalyzer SST Anomaly Globe -> map.png")
    print("=" * 60)
    try:
        acquire_image(OUTPUT_PNG)
    except Exception as e:
        print(f"ERROR: {e}")
        return 2
    print(f"Saved: {OUTPUT_PNG}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
