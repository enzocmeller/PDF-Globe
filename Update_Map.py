"""
Update_Map.py
-------------
Portable script. Downloads the Pacific-centered SST Anomaly orthographic
globe from Climate Reanalyzer and saves it next to this script as `map.png`.

Source image (same one shown at the top of the page):
  https://climatereanalyzer.org/wx/todays-weather/?var_id=sstanom&ortho=6&wt=1

Why this version is robust
==========================
* Pure Python 3.8+. No virtualenv or pip required.
* Tries 4 strategies in order, so it works even if one fails:
    1. urllib (stdlib)  -- zero dependencies
    2. requests         -- if available; auto-installed if pip works
    3. Playwright       -- if installed, uses a real Chromium browser
    4. Element screenshot via Playwright as a last resort
* Atomic write to map.png so a mid-download failure can't leave a
  half-written file lying around.
* Output is written next to this script so you can copy the file to any
  machine and run it from any working directory.

Usage
=====
    python Update_Map.py
"""

from __future__ import annotations

import os
import shutil
import ssl
import subprocess
import sys
import time
import traceback
import urllib.request
import urllib.error
from pathlib import Path

# ---------------------------------------------------------------------------
# CONFIG -- you usually don't need to edit anything here.
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_PNG = SCRIPT_DIR / "map.png"

# Climate Reanalyzer source page (ortho=6 -> Pacific-centered globe).
PAGE_URL = (
    "https://climatereanalyzer.org/wx/todays-weather/"
    "?var_id=sstanom&ortho=6&wt=1"
)

# Direct URL of the globe PNG (set by the page's JS:
#   document.img_sat6.src = "maps/gfs_pacific-sat_sstanom_d1.png").
#
# ortho -> filename token mapping (from js/tw_v11.min.js), in case you
# ever want a different orthographic view:
#   1 nh-sat1     2 samer-sat    3 euroafr-sat   4 asia-sat
#   5 ausnz-sat   6 pacific-sat  7 spole-sat     8 npole-sat
DIRECT_PNG_URL = (
    "https://climatereanalyzer.org/wx/todays-weather/maps/"
    "gfs_pacific-sat_sstanom_d1.png"
)
TARGET_IMG_ID = "img_sat6"

# Browser-like headers (the CDN 403's bare Python User-Agents).
BROWSER_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {
    "User-Agent": BROWSER_UA,
    "Referer": PAGE_URL,
    "Accept": "image/avif,image/webp,image/png,image/*,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

REQUEST_TIMEOUT = 30          # seconds, per HTTP attempt
HTTP_MAX_TRIES = 3            # retry policy for transient ECONNRESETs
PLAYWRIGHT_TIMEOUT_MS = 60000 # generous; corporate networks can be slow


# ===========================================================================
# Strategy 1 -- stdlib urllib (no dependencies at all).
# ===========================================================================
def download_with_urllib(url: str, dest: Path) -> bool:
    last_err = None
    for attempt in range(1, HTTP_MAX_TRIES + 1):
        try:
            print(f"[urllib] GET {url}  (attempt {attempt}/{HTTP_MAX_TRIES})")
            req = urllib.request.Request(url, headers=HEADERS)
            ctx = ssl.create_default_context()
            with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT, context=ctx) as resp:
                status = getattr(resp, "status", 200)
                if status != 200:
                    print(f"[urllib] HTTP {status}.")
                    return False
                ctype = resp.headers.get("Content-Type", "")
                if "image" not in ctype:
                    print(f"[urllib] Unexpected Content-Type {ctype!r}.")
                    return False
                with open(dest, "wb") as f:
                    shutil.copyfileobj(resp, f)
            return True
        except (urllib.error.URLError, TimeoutError, ConnectionResetError, OSError) as e:
            last_err = e
            print(f"[urllib] error on attempt {attempt}: {e}")
            time.sleep(1.5 * attempt)   # gentle backoff
        except Exception as e:
            print(f"[urllib] unexpected error: {e}")
            return False
    print(f"[urllib] giving up after {HTTP_MAX_TRIES} attempts ({last_err}).")
    return False


# ===========================================================================
# Strategy 2 -- requests (auto-installed if needed and pip is reachable).
# ===========================================================================
def _try_get_requests():
    try:
        import requests  # noqa: F401
        return requests
    except ImportError:
        pass

    if os.environ.get("GET_GLOBE_NO_PIP"):
        return None

    try:
        print("[setup] 'requests' not found; attempting pip install...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--quiet",
             "--disable-pip-version-check", "requests"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.STDOUT,
        )
        import requests  # noqa: F401
        return requests
    except Exception as e:
        print(f"[setup] pip install failed ({e}); skipping requests.")
        return None


def download_with_requests(url: str, dest: Path) -> bool:
    requests_mod = _try_get_requests()
    if requests_mod is None:
        return False

    last_err = None
    for attempt in range(1, HTTP_MAX_TRIES + 1):
        try:
            print(f"[requests] GET {url}  (attempt {attempt}/{HTTP_MAX_TRIES})")
            r = requests_mod.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT, stream=True)
            if r.status_code != 200:
                print(f"[requests] HTTP {r.status_code}.")
                return False
            ctype = r.headers.get("Content-Type", "")
            if "image" not in ctype:
                print(f"[requests] Unexpected Content-Type {ctype!r}.")
                return False
            with open(dest, "wb") as f:
                for chunk in r.iter_content(8192):
                    f.write(chunk)
            return True
        except Exception as e:
            last_err = e
            print(f"[requests] error on attempt {attempt}: {e}")
            time.sleep(1.5 * attempt)
    print(f"[requests] giving up after {HTTP_MAX_TRIES} attempts ({last_err}).")
    return False


# ===========================================================================
# Strategy 3 -- Playwright (real browser). Honours your env's proxy/SSL.
# ===========================================================================
def download_with_playwright(page_url: str, dest: Path) -> bool:
    try:
        from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
    except ImportError:
        print("[playwright] Not installed; skipping.")
        return False

    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            ctx = browser.new_context(
                user_agent=BROWSER_UA,
                viewport={"width": 1280, "height": 1024},
            )
            page = ctx.new_page()
            print(f"[playwright] Loading {page_url}")
            page.goto(page_url, timeout=PLAYWRIGHT_TIMEOUT_MS,
                      wait_until="domcontentloaded")

            # Wait for the page's JS to populate #img_sat6.
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
                print(f"[playwright] Timed out waiting for #{TARGET_IMG_ID}.")

            img_src = page.evaluate(
                f"() => {{ const e = document.getElementById('{TARGET_IMG_ID}');"
                "  return e ? e.src : null; }"
            )
            print(f"[playwright] {TARGET_IMG_ID}.src = {img_src}")

            # 3a) Try downloading the resolved URL through the browser context
            #     (correct cookies/Referer). Retry once on ECONNRESET.
            if img_src:
                for attempt in range(1, 3):
                    try:
                        resp = ctx.request.get(
                            img_src,
                            headers={"Referer": page_url, "User-Agent": BROWSER_UA},
                            timeout=PLAYWRIGHT_TIMEOUT_MS,
                        )
                        if resp.ok and "image" in resp.headers.get("content-type", ""):
                            dest.write_bytes(resp.body())
                            print(f"[playwright] ctx.request OK "
                                  f"({dest.stat().st_size:,} bytes)")
                            browser.close()
                            return True
                        print(f"[playwright] ctx.request status="
                              f"{getattr(resp,'status',None)}")
                    except Exception as e:
                        print(f"[playwright] ctx.request attempt {attempt} failed: {e}")
                        time.sleep(1.5)

            # 3b) Last resort: screenshot the element. The destination must end
            #     in .png/.jpeg/.jpg or Playwright refuses; pass a real .png
            #     path and copy onto dest afterwards.
            try:
                ss_path = dest.with_suffix(".png") if dest.suffix.lower() != ".png" \
                          else dest
                locator = page.locator(f"#{TARGET_IMG_ID}")
                locator.scroll_into_view_if_needed(timeout=5000)
                # Force PNG via 'type=' so Playwright doesn't reject .part files.
                shot_bytes = locator.screenshot(type="png")
                ss_path.write_bytes(shot_bytes)
                print(f"[playwright] Element screenshot saved "
                      f"({ss_path.stat().st_size:,} bytes)")
                browser.close()
                return ss_path.exists() and ss_path.stat().st_size > 5_000
            except Exception as e:
                print(f"[playwright] Element screenshot failed: {e}")
                browser.close()
                return False
    except Exception as e:
        print(f"[playwright] Fatal: {e}")
        traceback.print_exc()
        return False


# ===========================================================================
# Helpers.
# ===========================================================================
def _looks_like_png(path: Path) -> bool:
    """Cheap PNG validation; no Pillow needed."""
    try:
        if path.stat().st_size < 10_000:
            return False
        with open(path, "rb") as f:
            magic = f.read(8)
        return magic == b"\x89PNG\r\n\x1a\n"
    except Exception:
        return False


def acquire_image(dest: Path) -> None:
    """Run strategies in order until one succeeds."""
    tmp = dest.with_suffix(".png.part")   # always .png.part to avoid
                                          # weird suffix stacking
    if tmp.exists():
        try:
            tmp.unlink()
        except Exception:
            pass

    strategies = [
        ("urllib",     lambda: download_with_urllib(DIRECT_PNG_URL, tmp)),
        ("requests",   lambda: download_with_requests(DIRECT_PNG_URL, tmp)),
        ("playwright", lambda: download_with_playwright(PAGE_URL, tmp)),
    ]

    for name, fn in strategies:
        print(f"--- trying strategy: {name} ---")
        try:
            ok = fn()
        except Exception as e:
            print(f"[{name}] crashed: {e}")
            ok = False
        if ok and _looks_like_png(tmp):
            os.replace(tmp, dest)
            return
        # tidy up between attempts
        if tmp.exists():
            try:
                tmp.unlink()
            except Exception:
                pass

    raise RuntimeError(
        "All download strategies failed. Possible causes:\n"
        "  * No internet on this machine.\n"
        "  * Corporate firewall/proxy blocking climatereanalyzer.org.\n"
        "  * The site temporarily down.\n"
        "Try opening this URL in a browser on this same PC to confirm:\n"
        f"  {DIRECT_PNG_URL}"
    )


def main() -> int:
    print("=" * 60)
    print("Climate Reanalyzer SST Anomaly Globe -> map.png")
    print("=" * 60)
    print(f"Python : {sys.version.split()[0]} ({sys.executable})")
    print(f"Output : {OUTPUT_PNG}")
    try:
        acquire_image(OUTPUT_PNG)
    except Exception as e:
        print(f"ERROR: {e}")
        return 2
    print(f"Saved: {OUTPUT_PNG} ({OUTPUT_PNG.stat().st_size:,} bytes)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
