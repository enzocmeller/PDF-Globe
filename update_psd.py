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

# --- Corporate-firewall friendliness ---------------------------------------
# Trust the operating system's certificate store (Windows) so HTTPS works
# behind SSL-inspection proxies that re-sign traffic with a company root CA.
# Safe no-op if truststore isn't installed.
try:
    import truststore as _truststore
    _truststore.inject_into_ssl()
except Exception:
    pass

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------
ARCHIVE_BASE = (
    "https://apps.fas.usda.gov/psdonline/downloads/archives/"
    "{year}/{month:02d}/{zip}"
)

# Each dataset = one USDA category file + the commodities to keep from it.
# Outputs go to  data/psd_<name>_latest.csv  and  data/psd_<name>_prior.csv.
#
# To add a commodity, append its EXACT Commodity_Description (as USDA spells it)
# to the right list -- or add a whole new dataset block.
DATASETS = [
    {
        "name": "grains",
        "zip": "psd_grains_pulses_csv.zip",
        "commodities": ["Barley", "Corn", "Wheat"],
    },
    {
        "name": "oilseeds",
        "zip": "psd_oilseeds_csv.zip",
        # USDA spelling:  the bean / the meal / the oil
        "commodities": ["Oilseed, Soybean", "Meal, Soybean", "Oil, Soybean"],
    },
]

# Which file decides which monthly releases are posted (all categories publish
# on the same monthly cycle, so any always-present file works).
RELEASE_DETECT_ZIP = "psd_grains_pulses_csv.zip"

SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIR = SCRIPT_DIR / "data"
STATE_FILE = DATA_DIR / "_state.json"

HEADERS = {"User-Agent": "Mozilla/5.0 (PSD updater)"}

# How many months back to search for available releases before giving up.
MAX_LOOKBACK_MONTHS = 24


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def log(msg: str) -> None:
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}")


def month_url(year: int, month: int, zip_name: str) -> str:
    return ARCHIVE_BASE.format(year=year, month=month, zip=zip_name)


def prev_month(year: int, month: int) -> tuple[int, int]:
    return (year - 1, 12) if month == 1 else (year, month - 1)


def release_exists(year: int, month: int) -> str | None:
    """Return Last-Modified if the release is posted, else None."""
    req = Request(month_url(year, month, RELEASE_DETECT_ZIP),
                  headers=HEADERS, method="HEAD")
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


def download_filtered(year: int, month: int, zip_name: str,
                      commodities: list[str]) -> pd.DataFrame:
    """Download one category zip for a release, extract its CSV, filter to the
    requested commodities."""
    url = month_url(year, month, zip_name)
    log(f"Downloading {year}-{month:02d}: {url}")
    req = Request(url, headers=HEADERS)
    with urlopen(req, timeout=180) as resp:
        zip_bytes = resp.read()

    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        csv_names = [n for n in zf.namelist() if n.lower().endswith(".csv")]
        if not csv_names:
            raise RuntimeError(f"No CSV inside {zip_name} for {year}-{month:02d}.")
        with zf.open(csv_names[0]) as fh:
            csv_text = fh.read().decode("utf-8-sig")

    df = pd.read_csv(io.StringIO(csv_text))
    out = df[df["Commodity_Description"].str.strip().isin(commodities)].copy()
    if out.empty:
        raise RuntimeError(
            f"{zip_name} {year}-{month:02d}: filter produced 0 rows -- check the "
            f"Commodity_Description values {commodities}."
        )
    counts = out["Commodity_Description"].value_counts().to_dict()
    log(f"  {zip_name} {year}-{month:02d}: kept {len(out):,} rows {counts}")
    return out


def write_release(year: int, month: int, suffix: str) -> None:
    """Download + write every dataset for one release ('latest' or 'prior').
    A single dataset's network failure is non-fatal -- the existing CSV is kept."""
    for ds in DATASETS:
        out_path = DATA_DIR / f"psd_{ds['name']}_{suffix}.csv"
        try:
            df = download_filtered(year, month, ds["zip"], ds["commodities"])
            df.to_csv(out_path, index=False)
            log(f"Wrote {out_path.name}  <- {ds['name']} release {year}-{month:02d}")
        except (HTTPError, URLError) as e:
            log(f"WARNING: {ds['name']} {year}-{month:02d} download failed ({e}); "
                f"keeping existing {out_path.name} if present.")


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
    write_release(ly, lm, "latest")
    state = {
        "latest": {"year": ly, "month": lm, "last_modified": l_mod,
                   "label": f"{ly}-{lm:02d}"},
        "datasets": [d["name"] for d in DATASETS],
        "updated": datetime.now(timezone.utc).isoformat(),
    }

    # ----- prior ------------------------------------------------------------
    if len(releases) >= 2:
        py, pm, p_mod = releases[1]
        write_release(py, pm, "prior")
        state["prior"] = {"year": py, "month": pm, "last_modified": p_mod,
                          "label": f"{py}-{pm:02d}"}
    else:
        log("Only one release available; prior-month files not written yet.")

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
