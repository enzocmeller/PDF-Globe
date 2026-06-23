# Setup & firewall notes

## TL;DR — how to run it on any machine

1. Make sure **Python 3.10+ (64-bit)** and **Microsoft Office** are installed.
2. Double-click **`run_update.bat`**.
   - On the first run it auto-runs **`setup.bat`**, which creates `.venv` and
     installs all dependencies. No manual steps.
3. The deck lands on your Desktop: `AUS SnD Commodities 2026 - <date>.pdf`.

You can also run `setup.bat` by itself once, ahead of time.

## How dependencies install (firewall-tolerant)

`setup.bat` installs in this order, stopping at the first that works:

1. **Offline** — from the bundled `wheels\` folder (`pip --no-index`).
   Needs **no internet at all**. This is the firewall-proof path.
2. **Online fallback** — normal `pip` with `--trusted-host` flags so it still
   works behind a proxy that re-signs HTTPS (SSL inspection).

The bundled `wheels\` were built for **Python 3.13, 64-bit**. If the target
machine has a *different* Python version, the compiled wheels (pandas, numpy,
Pillow, pywin32) won't match and setup falls back to online pip. To make a new
offline cache, on a machine **with internet** that has the same Python:

```
pip download -r requirements.txt -d wheels
```

then copy the `wheels\` folder next to these scripts.

## What still needs network at RUN time (and the firewall reality)

The pipeline downloads fresh data from three public hosts each run:

| What | Host |
|------|------|
| USDA PSD grain data | `apps.fas.usda.gov` |
| SST anomaly globe | `climatereanalyzer.org` |
| El Niño forecast | `cpc.ncep.noaa.gov` |

- **SSL inspection** (JBS re-signing HTTPS with a company root CA): handled.
  All scripts call `truststore`, which trusts the **Windows certificate store**,
  so the company CA is accepted instead of throwing `CERTIFICATE_VERIFY_FAILED`.
- **Proxy**: honored automatically via the Windows system proxy settings.
- **Outright domain blocking**: if JBS's firewall *blocks* one of those three
  hosts entirely, **no code can work around that** — the data simply can't be
  fetched. The fix is to ask IT to allow-list the host, or run on a network
  that can reach it. The scripts fail with a clear message naming the host and
  URL so you know exactly what to request.

**Quick test for a blocked host:** open the URL in a browser on the same PC. If
the browser can't reach it, the script can't either, and it's a firewall block
to raise with IT — for example:
`https://apps.fas.usda.gov/psdonline/downloads/archives/2026/06/psd_grains_pulses_csv.zip`

## What must already be on the machine (can't be bootstrapped)

- **Python 3.10+ 64-bit** — the scripts run on it; `setup.bat` only builds the
  venv, it can't install Python itself.
- **Microsoft Excel + the workbook** — the deck is built by automating real
  Excel (refresh, image insert, PDF export).

## Honest bottom line

- Dependency install: **solved** (offline wheels + online fallback).
- SSL inspection / proxy: **solved** (truststore + system proxy).
- Hard domain blocks at runtime: **cannot be solved in code** — needs an IT
  allow-list. This is the one thing to verify with a quick browser test before
  relying on it at JBS.
