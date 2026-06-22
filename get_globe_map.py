"""
get_globe_map.py
----------------
Standalone, portable script. Downloads the Pacific-centered SST Anomaly
orthographic globe from Climate Reanalyzer and saves it next to this
script as `map.png`.

This is the SAME image as the one shown at the top of:
  https://climatereanalyzer.org/wx/todays-weather/?var_id=sstanom&ortho=6&wt=1

Portable design
===============
* Pure Python 3.8+; no virtualenv required.
* Auto-installs its only optional dependency (`requests`) on first run
  if it's missing. If `pip install` fails (e.g. offline), it falls back
  to the standard library's urllib so it still works.
* Output is written next to the script, so copying the .py file to any
  machine and running `python get_globe_map.py` is enough.

Usage
=====
    python get_globe_map.py
"""

from __future__ import annotations

import os
import shutil
import ssl
import subprocess
import sys
import urllib.request
from pathlib import Path

# ---------------------------------------------------------------------------
# CONFIG -- you usually don't need to edit anything here.
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_PNG = SCRIPT_DIR / "map.png"

# Climate Reanalyzer source page (ortho=6 -> Pacific-centered globe).
PAGE_URL = "https://climatereanalyzer.org/wx/todays-weather/?var_id=sstanom&ortho=6&wt=1"

# Direct URL of the globe PNG (set by the page's JS:
#   document.img_sat6.src = "maps/gfs_pacific-sat_sstanom_d1.png").
#
# ortho -> filename token mapping (from js/tw_v11.min.js), if you ever
# want a different orthographic view:
#   1 nh-sat1     2 samer-sat    3 euroafr-sat   4 asia-sat
#   5 ausnz-sat   6 pacific-sat  7 spole-sat     8 npole-sat
DIRECT_PNG_URL = (
    "https://climatereanalyzer.org/wx/todays-weather/maps/"
    "gfs_pacific-sat_sstanom_d1.png"
)

# Browser-like headers (the CDN 403's bare User-Agents).
BROWSER_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {
    "User-Agent": BROWSER_UA,
    "Referer": PAGE_URL,
    "Accept": "image/avif,image/webp,image/png,image/*,*/*;q=0.8",
}

REQUEST_TIMEOUT = 30  # seconds


# ---------------------------------------------------------------------------
# Tiny self-bootstrap: try to ensure `requests` is available. If it's
# missing, attempt `pip install requests` once. If that fails (offline,
# restricted environment, etc.) we silently fall through to urllib.
# ---------------------------------------------------------------------------
def _try_get_requests():
    try:
        import requests  # noqa: F401
        return requests
    except ImportError:
        pass

    # Don't try to install if we're explicitly told not to.
    if os.environ.get("GET_GLOBE_NO_PIP"):
        return None

    try:
        print("[setup] 'requests' not found; attempting 'pip install requests'...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--quiet",
             "--disable-pip-version-check", "requests"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.STDOUT,
        )
        import requests  # noqa: F401
        return requests
    except Exception as e:
        print(f"[setup] pip install failed ({e}); will use stdlib urllib instead.")
        return None


# ---------------------------------------------------------------------------
# Download helpers.
# ---------------------------------------------------------------------------
def _download_with_requests(requests_mod, url: str, dest: Path) -> bool:
    try:
        print(f"[download] GET {url}")
        r = requests_mod.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT, stream=True)
        if r.status_code != 200:
            print(f"[download] HTTP {r.status_code}.")
            return False
        ctype = r.headers.get("Content-Type", "")
        if "image" not in ctype:
            print(f"[download] Unexpected Content-Type {ctype!r}.")
            return False
        with open(dest, "wb") as f:
            for chunk in r.iter_content(8192):
                f.write(chunk)
        return True
    except Exception as e:
        print(f"[download] requests error: {e}")
        return False


def _download_with_urllib(url: str, dest: Path) -> bool:
    """Standard-library fallback so the script works even with zero deps."""
    try:
        print(f"[download/urllib] GET {url}")
        req = urllib.request.Request(url, headers=HEADERS)
        ctx = ssl.create_default_context()
        with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT, context=ctx) as resp:
            if resp.status != 200:
                print(f"[download/urllib] HTTP {resp.status}.")
                return False
            ctype = resp.headers.get("Content-Type", "")
            if "image" not in ctype:
                print(f"[download/urllib] Unexpected Content-Type {ctype!r}.")
                return False
            with open(dest, "wb") as f:
                shutil.copyfileobj(resp, f)
        return True
    except Exception as e:
        print(f"[download/urllib] Error: {e}")
        return False


def _looks_like_png(path: Path) -> bool:
    """Cheap PNG validation -- no Pillow required."""
    try:
        if path.stat().st_size < 10_000:
            return False
        with open(path, "rb") as f:
            magic = f.read(8)
        return magic == b"\x89PNG\r\n\x1a\n"
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Orchestrator.
# ---------------------------------------------------------------------------
def acquire_image(dest: Path) -> None:
    """Download with requests if available, otherwise urllib. Atomic write."""
    tmp = dest.with_suffix(dest.suffix + ".part")
    if tmp.exists():
        try:
            tmp.unlink()
        except Exception:
            pass

    ok = False
    requests_mod = _try_get_requests()
    if requests_mod is not None:
        ok = _download_with_requests(requests_mod, DIRECT_PNG_URL, tmp)

    if not ok:
        ok = _download_with_urllib(DIRECT_PNG_URL, tmp)

    if not ok or not _looks_like_png(tmp):
        try:
            tmp.unlink()
        except Exception:
            pass
        raise RuntimeError(
            "Could not download the SST anomaly globe image. Check your "
            "internet connection and that climatereanalyzer.org is reachable."
        )

    # Atomic replace (works on Windows + POSIX).
    os.replace(tmp, dest)


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
    size = OUTPUT_PNG.stat().st_size
    print(f"Saved: {OUTPUT_PNG} ({size:,} bytes)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
