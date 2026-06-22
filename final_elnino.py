"""
final_elnino.py
---------------
Downloads the Nino 3.4 forecast image from NOAA CPC, trims a little
off the sides, and saves it next to this script as elnino.png.

Source:
  https://www.cpc.ncep.noaa.gov/products/CFSv2/cfsv2fcst/imagesInd3/nino34Mon.gif

Output (in the same folder as this script):
  elnino.png   - trimmed image, ready to insert into Excel

PNG is chosen over PDF because Excel can embed PNG natively via
Pictures.Insert / Shapes.AddPicture. PDF is not a supported image
format for Excel cells and would require external conversion.

Portable: pure Python 3.8+. Uses urllib (stdlib) first, then falls back
to requests if available, then Playwright as a last resort.
Auto-installs Pillow if pip works.

Run:
    python "final_elnino.py"
"""

from __future__ import annotations

import gzip
import io
import os
import shutil
import ssl
import subprocess
import sys
import time
import traceback
import urllib.request
import urllib.error
import zlib
from pathlib import Path

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
SCRIPT_DIR  = Path(__file__).resolve().parent
OUTPUT_PNG  = SCRIPT_DIR / "elnino.png"

SOURCE_URL = (
    "https://www.cpc.ncep.noaa.gov/products/CFSv2/cfsv2fcst/imagesInd3/"
    "nino34Mon.gif"
)
SOURCE_REFERER = "https://www.cpc.ncep.noaa.gov/products/CFSv2/CFSv2seasonal.shtml"

# Trim as a fraction of the original image dimensions. 0.05 == 5%.
# These are independent so you can shave more off one side if needed.
TRIM_LEFT_FRAC   = 0.10
TRIM_RIGHT_FRAC  = 0.11
TRIM_TOP_FRAC    = 0.075
TRIM_BOTTOM_FRAC = 0.05

BROWSER_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {
    "User-Agent": BROWSER_UA,
    "Referer": SOURCE_REFERER,
    "Accept": "image/avif,image/webp,image/png,image/gif,image/*,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Cache-Control": "no-cache",
    "Pragma": "no-cache",
    "Sec-Fetch-Dest": "image",
    "Sec-Fetch-Mode": "no-cors",
    "Sec-Fetch-Site": "same-origin",
    "Connection": "close",
}

REQUEST_TIMEOUT = 45
HTTP_MAX_TRIES  = 4
PLAYWRIGHT_TIMEOUT_MS = 60000


# ---------------------------------------------------------------------------
# Bootstrap: make sure Pillow is available (needed for trim + PDF).
# ---------------------------------------------------------------------------
def _ensure_pillow():
    try:
        from PIL import Image  # noqa: F401
        return True
    except ImportError:
        pass
    if os.environ.get("NO_PIP"):
        print("ERROR: Pillow is required. Install with: pip install pillow")
        return False
    try:
        print("[setup] Installing Pillow ...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--quiet",
             "--disable-pip-version-check", "pillow"],
            stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT,
        )
        from PIL import Image  # noqa: F401
        return True
    except Exception as e:
        print(f"ERROR: could not install Pillow ({e}). "
              "Please run: pip install pillow")
        return False


# ---------------------------------------------------------------------------
# Download strategies. NOAA CPC sometimes RSTs Python-style traffic on the
# first attempt; we try several techniques in order until one succeeds.
# ---------------------------------------------------------------------------
def _decompress(payload: bytes, encoding: str) -> bytes:
    """Handle gzip/deflate/br Content-Encoding ourselves (urllib doesn't)."""
    enc = (encoding or "").lower().strip()
    if not enc or enc == "identity":
        return payload
    if enc == "gzip":
        return gzip.decompress(payload)
    if enc == "deflate":
        # Some servers send raw deflate, some zlib-wrapped; try both.
        try:
            return zlib.decompress(payload)
        except zlib.error:
            return zlib.decompress(payload, -zlib.MAX_WBITS)
    if enc == "br":
        try:
            import brotli  # type: ignore
            return brotli.decompress(payload)
        except ImportError:
            # We sent Accept-Encoding: br, but if it actually comes back br
            # and brotli isn't installed we'll have to retry without it.
            raise RuntimeError("server returned brotli; install 'brotli' "
                               "(pip install brotli) or use another strategy")
    return payload


def download_with_urllib(url: str, dest: Path) -> bool:
    last_err = None
    for attempt in range(1, HTTP_MAX_TRIES + 1):
        try:
            print(f"[urllib] GET {url}  (attempt {attempt}/{HTTP_MAX_TRIES})")
            # On retries after a reset, drop br from Accept-Encoding to avoid
            # the rare 'brotli not installed' path biting us.
            headers = dict(HEADERS)
            if attempt > 1:
                headers["Accept-Encoding"] = "gzip, deflate"
            req = urllib.request.Request(url, headers=headers)
            ctx = ssl.create_default_context()
            with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT, context=ctx) as resp:
                if getattr(resp, "status", 200) != 200:
                    print(f"[urllib] HTTP {resp.status}.")
                    return False
                raw = resp.read()
                enc = resp.headers.get("Content-Encoding", "")
                try:
                    body = _decompress(raw, enc)
                except Exception as e:
                    print(f"[urllib] decompress failed ({enc}): {e}")
                    return False
                with open(dest, "wb") as f:
                    f.write(body)
            return True
        except (urllib.error.URLError, TimeoutError, ConnectionResetError, OSError) as e:
            last_err = e
            print(f"[urllib] error on attempt {attempt}: {e}")
            time.sleep(1.5 * attempt)
        except Exception as e:
            print(f"[urllib] unexpected error: {e}")
            return False
    print(f"[urllib] giving up after {HTTP_MAX_TRIES} attempts ({last_err}).")
    return False


def download_with_requests(url: str, dest: Path) -> bool:
    try:
        import requests
        from requests.adapters import HTTPAdapter
    except ImportError:
        return False
    # Use a Session with a fresh adapter on each attempt and Connection: close
    # to defeat any keep-alive state that's tripping the remote RST.
    last_err = None
    for attempt in range(1, HTTP_MAX_TRIES + 1):
        try:
            print(f"[requests] GET {url}  (attempt {attempt}/{HTTP_MAX_TRIES})")
            sess = requests.Session()
            sess.mount("https://", HTTPAdapter(max_retries=0, pool_connections=1))
            headers = dict(HEADERS)
            if attempt > 1:
                headers["Accept-Encoding"] = "gzip, deflate"  # drop br
            r = sess.get(url, headers=headers, timeout=REQUEST_TIMEOUT, stream=True)
            if r.status_code != 200:
                print(f"[requests] HTTP {r.status_code}.")
                sess.close()
                return False
            with open(dest, "wb") as f:
                for chunk in r.iter_content(8192):
                    f.write(chunk)
            sess.close()
            return True
        except Exception as e:
            last_err = e
            print(f"[requests] error on attempt {attempt}: {e}")
            time.sleep(2.0 * attempt)
    print(f"[requests] giving up after {HTTP_MAX_TRIES} attempts ({last_err}).")
    return False


def download_with_playwright(url: str, dest: Path) -> bool:
    """Last resort: drive real Chromium, which uses Windows' own networking
    stack and proxy settings. Survives firewalls that RST Python traffic."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("[playwright] Not installed; skipping. "
              "To enable: pip install playwright && python -m playwright install chromium")
        return False

    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            ctx = browser.new_context(
                user_agent=BROWSER_UA,
                viewport={"width": 1280, "height": 1024},
            )
            # Visit the parent page first so Referer/cookies are right.
            page = ctx.new_page()
            try:
                print(f"[playwright] Warming up: {SOURCE_REFERER}")
                page.goto(SOURCE_REFERER, timeout=PLAYWRIGHT_TIMEOUT_MS,
                          wait_until="domcontentloaded")
            except Exception as e:
                print(f"[playwright] (referer page failed, continuing): {e}")

            # 1) ctx.request - uses Chromium's network stack.
            for attempt in range(1, 4):
                try:
                    print(f"[playwright] ctx.request GET {url} (attempt {attempt})")
                    resp = ctx.request.get(
                        url,
                        headers={"Referer": SOURCE_REFERER, "User-Agent": BROWSER_UA},
                        timeout=PLAYWRIGHT_TIMEOUT_MS,
                    )
                    if resp.ok:
                        body = resp.body()
                        if body and len(body) > 1000:
                            dest.write_bytes(body)
                            print(f"[playwright] ctx.request OK "
                                  f"({dest.stat().st_size:,} bytes)")
                            browser.close()
                            return True
                    print(f"[playwright] ctx.request status="
                          f"{getattr(resp, 'status', None)}")
                except Exception as e:
                    print(f"[playwright] ctx.request attempt {attempt} failed: {e}")
                    time.sleep(2.0)

            # 2) In-page fetch() - uses the very same network as a real visit.
            try:
                print("[playwright] trying in-page fetch() ...")
                b64 = page.evaluate(
                    """async (url) => {
                        const r = await fetch(url, { credentials: 'include' });
                        if (!r.ok) return null;
                        const buf = await r.arrayBuffer();
                        const bytes = new Uint8Array(buf);
                        let bin = '';
                        for (let i = 0; i < bytes.length; i++)
                            bin += String.fromCharCode(bytes[i]);
                        return btoa(bin);
                    }""",
                    url,
                )
                if b64:
                    import base64
                    dest.write_bytes(base64.b64decode(b64))
                    size = dest.stat().st_size
                    if size > 1000:
                        print(f"[playwright] in-page fetch OK ({size:,} bytes)")
                        browser.close()
                        return True
                    print(f"[playwright] in-page fetch returned only {size} bytes")
                else:
                    print("[playwright] in-page fetch returned null")
            except Exception as e:
                print(f"[playwright] in-page fetch failed: {e}")

            # 3) Last resort: navigate to the image URL and screenshot it.
            try:
                print("[playwright] navigating to image URL and screenshotting ...")
                page.goto(url, timeout=PLAYWRIGHT_TIMEOUT_MS,
                          wait_until="load")
                img = page.locator("img").first
                shot = img.screenshot(type="png")
                dest.write_bytes(shot)
                size = dest.stat().st_size
                print(f"[playwright] screenshot saved ({size:,} bytes)")
                browser.close()
                return size > 1000
            except Exception as e:
                print(f"[playwright] screenshot failed: {e}")
                browser.close()
                return False
    except Exception as e:
        print(f"[playwright] Fatal: {e}")
        traceback.print_exc()
        return False


def acquire_raw(dest: Path) -> None:
    """Download the raw image (may be GIF or PNG, server decides)."""
    tmp = dest.with_suffix(dest.suffix + ".part")
    if tmp.exists():
        try:
            tmp.unlink()
        except Exception:
            pass

    strategies = [
        ("urllib",     lambda: download_with_urllib(SOURCE_URL, tmp)),
        ("requests",   lambda: download_with_requests(SOURCE_URL, tmp)),
        ("playwright", lambda: download_with_playwright(SOURCE_URL, tmp)),
    ]
    for name, fn in strategies:
        print(f"--- trying strategy: {name} ---")
        try:
            ok = fn()
        except Exception as e:
            print(f"[{name}] crashed: {e}")
            ok = False
        if ok and tmp.exists() and tmp.stat().st_size > 1000:
            os.replace(tmp, dest)
            return
        if tmp.exists():
            try:
                tmp.unlink()
            except Exception:
                pass

    raise RuntimeError(
        "Could not download the Nino 3.4 image. Possible causes:\n"
        "  * No internet on this machine.\n"
        "  * Corporate firewall/proxy blocking cpc.ncep.noaa.gov.\n"
        "  * The site temporarily down.\n"
        f"Try opening this URL in a browser on this same PC:\n  {SOURCE_URL}"
    )


# ---------------------------------------------------------------------------
# Trim + save outputs.
# ---------------------------------------------------------------------------
def trim_and_save(raw_path: Path, png_out: Path) -> None:
    from PIL import Image

    with Image.open(raw_path) as im:
        # If animated, take the first frame.
        try:
            im.seek(0)
        except Exception:
            pass

        # Convert to RGB up-front so Excel inserts a clean opaque image.
        im = im.convert("RGB")

        w, h = im.size
        left   = int(w * TRIM_LEFT_FRAC)
        right  = w - int(w * TRIM_RIGHT_FRAC)
        top    = int(h * TRIM_TOP_FRAC)
        bottom = h - int(h * TRIM_BOTTOM_FRAC)

        print(f"[trim] original {w}x{h} -> crop box "
              f"(left={left}, top={top}, right={right}, bottom={bottom}) "
              f"= {right-left}x{bottom-top}")

        if right <= left or bottom <= top:
            raise RuntimeError("Trim percentages are too aggressive; "
                               "nothing left of the image.")

        cropped = im.crop((left, top, right, bottom))

    cropped.save(png_out, "PNG", optimize=True)
    print(f"[save] PNG -> {png_out} ({png_out.stat().st_size:,} bytes)")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def _purge_old_outputs() -> None:
    """Delete previous outputs and any leftover temp files (including any
    legacy elnino.pdf from earlier versions) so each run starts clean."""
    candidates = [
        OUTPUT_PNG,
        OUTPUT_PNG.with_suffix(OUTPUT_PNG.suffix + ".part"),
        SCRIPT_DIR / "elnino.pdf",          # legacy from earlier version
        SCRIPT_DIR / "elnino.pdf.part",     # legacy
        SCRIPT_DIR / "_elnino_raw.bin",
        SCRIPT_DIR / "_elnino_raw.bin.part",
    ]
    for p in candidates:
        try:
            if p.exists():
                p.unlink()
                print(f"[cleanup] removed {p.name}")
        except Exception as e:
            print(f"[cleanup] could not remove {p.name}: {e}")


def main() -> int:
    print("=" * 60)
    print("CPC Nino 3.4 forecast -> trimmed elnino.png")
    print("=" * 60)
    print(f"Python : {sys.version.split()[0]} ({sys.executable})")
    print(f"PNG    : {OUTPUT_PNG}")

    # Wipe previous outputs first so a failed run can't leave stale files.
    _purge_old_outputs()

    if not _ensure_pillow():
        return 3

    raw_path = SCRIPT_DIR / "_elnino_raw.bin"
    try:
        acquire_raw(raw_path)
        trim_and_save(raw_path, OUTPUT_PNG)
    except Exception as e:
        print(f"ERROR: {e}")
        return 2
    finally:
        if raw_path.exists():
            try:
                raw_path.unlink()
            except Exception:
                pass

    print("Done.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
