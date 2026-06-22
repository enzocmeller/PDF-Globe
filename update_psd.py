from __future__ import annotations

import io
import json
import sys
import zipfile
from datetime import date, datetime, timezone
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

import pandas as pd

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------
ARCHIVE_URL = (
    "https://apps.fas.usda.gov/psdonline/downloads/archives/"
    "{year}/{month:02d}/psd_grains_pulses_csv.zip"
)

COMMODITIES = ["Barley", "Corn", "Wheat"]   # matched on Commodity_Description

SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIR = SCRIPT_DIR / "data"
LATEST_CSV = DATA_DIR / "psd_grains_latest.csv"
PRIOR_CSV = DATA_DIR / "psd_grains_prior.csv"
STATE_FILE = DATA_DIR / "_state.json"

HEADERS = {"User-Agent": "Mozilla/5.0 (PSD grains updater)"}

# How many months back to search for available releases before giving up.
MAX_LOOKBACK_MONTHS = 24


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def log(msg: str) -> None:
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}")


def month_url(year: int, month: int) -> str:
    return ARCHIVE_URL.format(year=year, month=month)


def prev_month(year: int, month: int) -> tuple[int, int]:
    return (year - 1, 12) if month == 1 else (year, month - 1)


def release_exists(year: int, month: int) -> str | None:
    """Return Last-Modified if the archived grains file exists, else None."""
    req = Request(month_url(year, month), headers=HEADERS, method="HEAD")
    try:
        with urlopen(req, timeout=60) as resp:
            return resp.headers.get("Last-Modified", "")
    except HTTPError as e:
        if e.code == 404:
            return None
        raise


def find_recent_releases(count: int = 2) -> list[tuple[int, int, str]]:
    """Walk backwards from this month, returning the newest `count` releases
    that are actually posted, as (year, month, last_modified), newest first."""
    found: list[tuple[int, int, str]] = []
    y, m = date.today().year, date.today().month
    for _ in range(MAX_LOOKBACK_MONTHS):
        last_mod = release_exists(y, m)
        if last_mod is not None:
            found.append((y, m, last_mod))
            log(f"Found release {y}-{m:02d} (USDA Last-Modified: {last_mod or 'n/a'})")
            if len(found) == count:
                break
        y, m = prev_month(y, m)
    return found


def download_filtered(year: int, month: int) -> pd.DataFrame:
    """Download one release zip, extract its CSV, filter to the commodities."""
    url = month_url(year, month)
    log(f"Downloading {year}-{month:02d}: {url}")
    req = Request(url, headers=HEADERS)
    with urlopen(req, timeout=180) as resp:
        zip_bytes = resp.read()

    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        csv_names = [n for n in zf.namelist() if n.lower().endswith(".csv")]
        if not csv_names:
            raise RuntimeError(f"No CSV inside zip for {year}-{month:02d}.")
        with zf.open(csv_names[0]) as fh:
            csv_text = fh.read().decode("utf-8-sig")

    df = pd.read_csv(io.StringIO(csv_text))
    out = df[df["Commodity_Description"].str.strip().isin(COMMODITIES)].copy()
    if out.empty:
        raise RuntimeError(
            f"{year}-{month:02d}: filter produced 0 rows -- USDA may have "
            "changed the Commodity_Description values."
        )
    counts = out["Commodity_Description"].value_counts().to_dict()
    log(f"  {year}-{month:02d}: kept {len(out):,} rows {counts}")
    return out


def save_state(state: dict) -> None:
    STATE_FILE.write_text(json.dumps(state, indent=2))


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def main() -> int:
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    releases = find_recent_releases(count=2)
    if not releases:
        raise RuntimeError("No USDA releases found in the archive -- check the URL/network.")

    # ----- latest -----------------------------------------------------------
    ly, lm, l_mod = releases[0]
    download_filtered(ly, lm).to_csv(LATEST_CSV, index=False)
    log(f"Wrote {LATEST_CSV.name}  <- release {ly}-{lm:02d}")

    state = {
        "latest": {"year": ly, "month": lm, "last_modified": l_mod,
                   "label": f"{ly}-{lm:02d}"},
        "updated": datetime.now(timezone.utc).isoformat(),
    }

    # ----- prior ------------------------------------------------------------
    if len(releases) >= 2:
        py, pm, p_mod = releases[1]
        download_filtered(py, pm).to_csv(PRIOR_CSV, index=False)
        log(f"Wrote {PRIOR_CSV.name}   <- release {py}-{pm:02d}")
        state["prior"] = {"year": py, "month": pm, "last_modified": p_mod,
                          "label": f"{py}-{pm:02d}"}
    else:
        log("Only one release available; prior-month file not written yet.")

    save_state(state)
    log("Done.")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (HTTPError, URLError) as exc:
        log(f"NETWORK ERROR: {exc}")
        sys.exit(1)
    except Exception as exc:  # noqa: BLE001
        log(f"ERROR: {exc}")
        sys.exit(1)
        