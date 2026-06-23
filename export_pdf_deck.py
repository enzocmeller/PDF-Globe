"""
export_pdf_deck.py
------------------
Build ONE high-quality PDF "slide deck" from USDA_PSD.xlsx.

How it differs from the old pptx route (export_pdfs.py):
  old:  Excel -> PDF (vector) -> PNG at 400 DPI (RASTER) -> PowerPoint
  new:  Excel -> PDF (vector) -> placed AS VECTORS onto uniform slide pages
        -> one combined PDF

Because nothing is rasterized, text and charts stay perfectly crisp at any
zoom, the file is smaller, and there's no PowerPoint dependency.

Each worksheet in SLIDE_ORDER becomes one slide (a uniform-size PDF page with
the sheet centered on it, scaled to fit -- "contain").

Usage:
    python export_pdf_deck.py [workbook.xlsx] [output.pdf]

Defaults:
    workbook = USDA_PSD.xlsx  (next to this script)
    output   = Desktop\\AUS SnD Commodities 2026 - YYYY-MM-DD.pdf

Requires: PyMuPDF (fitz) and pywin32.
"""
from __future__ import annotations

import datetime
import os
import re
import sys
import tempfile
import time
from pathlib import Path

import fitz  # PyMuPDF


# Order of sheets -> slides. (Same list the pptx exporter used.)
SLIDE_ORDER = [
    "Intro Slide",
    "Globe & ElNino",
    "AUS Wheat",
    "AUS Wheat Chart",
    "Wheat S&D",
    "WRD Wheat Chart",
    "AUS Barley",
    "AUS Barley Chart",
    "Barley S&D",
    "Corn S&D",
    "WRD Corn Chart",
    # --- Soybean complex -------------------------------------------------
    # These slides are skipped (with a warning) until you create the matching
    # sheets in USDA_PSD.xlsx. Rename here to match whatever you name them.
    "Soybean S&D",
    "Soybean Meal S&D",
    "Soybean Oil S&D",
]

# Uniform slide canvas:
#   "auto"  -> pick the exported page size that wastes the least blank space
#              when every other page is contained inside it (matches old logic)
#   "16:9"  -> 13.333in x 7.5in widescreen
#   "first" -> use the first slide's page size for all
CANVAS_MODE = "auto"
MARGIN_PT = 0.0           # inner margin around each slide, in points (1pt = 1/72in)
WHITE_BACKGROUND = True   # paint each slide white before placing content
OPEN_FULLSCREEN = True    # ask PDF viewers to open in presentation mode

# Force each worksheet onto exactly ONE page when exporting, so every sheet
# becomes exactly one slide and nothing is cut off (wide S&D tables otherwise
# spill onto 2-4 pages). It stays vector, so you can zoom in for detail.
# Set False to keep each sheet's natural page count (multi-page sheets -> multi
# slides). This only affects the temporary export; the workbook is not modified.
FIT_SHEET_TO_ONE_PAGE = True

# Crop each slide to its actual (non-white) content before placing it, so the
# page hugs the design with no white border -> fills the screen edge to edge.
CROP_TO_CONTENT = True
CROP_WHITE_THRESHOLD = 248   # pixels >= this (0-255 gray) count as "white"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def normalize_sheet_name(name: str) -> str:
    s = str(name or "").strip().lower()
    s = re.sub(r"\b(page|sheet)\b", "", s)
    s = re.sub(r"[^a-z0-9]+", "", s)
    return s


def choose_canvas(sizes: list[tuple[float, float]]) -> tuple[float, float]:
    """Pick the slide canvas size from a list of (width, height) page sizes."""
    if not sizes:
        return (960.0, 540.0)
    if CANVAS_MODE == "first":
        return sizes[0]
    if CANVAS_MODE == "16:9":
        return (960.0, 540.0)  # 13.333" x 7.5" at 72pt/in
    # "auto": minimize total blank area when containing every page.
    best, best_blank = None, None
    for cw, ch in sizes:
        blank = 0.0
        for w, h in sizes:
            scale = min(cw / w, ch / h)
            blank += cw * ch - (w * h * scale * scale)
        if best_blank is None or blank < best_blank:
            best, best_blank = (cw, ch), blank
    return best


# ---------------------------------------------------------------------------
# PDF merge core  (no Excel here -> unit-testable on its own)
# ---------------------------------------------------------------------------
def content_clip(page) -> "fitz.Rect | None":
    """Return a Rect (page coords) tightly bounding the non-white content, or
    None if it can't be determined (then the full page is used)."""
    if not CROP_TO_CONTENT:
        return None
    try:
        from PIL import Image
    except Exception:
        return None
    try:
        zoom = 100 / 72.0
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom),
                              colorspace=fitz.csGRAY, alpha=False)
        img = Image.frombytes("L", (pix.width, pix.height), pix.samples)
        # non-white -> 255, white -> 0, then bbox of the non-white region
        mask = img.point(lambda p: 0 if p >= CROP_WHITE_THRESHOLD else 255)
        bbox = mask.getbbox()
        if not bbox:
            return None
        l, t, r, b = bbox
        return fitz.Rect(l / zoom, t / zoom, r / zoom, b / zoom)
    except Exception:
        return None


def merge_pdfs_to_deck(pdf_paths: list[Path], out_path: Path) -> Path:
    """Place every page of every input PDF, as vectors, onto uniform slide
    pages in one output PDF. Returns the path actually written."""
    src_docs = [fitz.open(str(p)) for p in pdf_paths]
    try:
        pages = []  # (doc, page_index, clip_or_None, width, height)
        for d in src_docs:
            for i in range(d.page_count):
                clip = content_clip(d[i])
                if clip is None:
                    clip = d[i].rect
                pages.append((d, i, clip, clip.width, clip.height))
        if not pages:
            raise RuntimeError("No pages found in the exported PDFs.")

        canvas_w, canvas_h = choose_canvas([(w, h) for _, _, _, w, h in pages])
        print(f"Slide canvas: {canvas_w:.0f} x {canvas_h:.0f} pt "
              f"({canvas_w / 72:.2f} x {canvas_h / 72:.2f} in), {len(pages)} slides")

        out = fitz.open()
        for d, i, clip, w, h in pages:
            page = out.new_page(width=canvas_w, height=canvas_h)
            if WHITE_BACKGROUND:
                try:
                    page.draw_rect(page.rect, color=None, fill=(1, 1, 1))
                except Exception:
                    pass
            avail_w = canvas_w - 2 * MARGIN_PT
            avail_h = canvas_h - 2 * MARGIN_PT
            scale = min(avail_w / w, avail_h / h)        # contain
            tw, th = w * scale, h * scale
            x0 = (canvas_w - tw) / 2                      # center
            y0 = (canvas_h - th) / 2
            rect = fitz.Rect(x0, y0, x0 + tw, y0 + th)
            page.show_pdf_page(rect, d, i, clip=clip)    # vector embed -> crisp

        if OPEN_FULLSCREEN:
            try:
                out.set_pagemode("FullScreen")
            except Exception:
                pass

        out_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            out.save(str(out_path), deflate=True, garbage=4)
        except Exception:
            ts = int(time.time())
            out_path = out_path.with_name(f"{out_path.stem}_{ts}{out_path.suffix}")
            out.save(str(out_path), deflate=True, garbage=4)
        out.close()
        return out_path
    finally:
        for d in src_docs:
            try:
                d.close()
            except Exception:
                pass


# ---------------------------------------------------------------------------
# Excel export  (drives Excel via COM -- unchanged approach from the old code)
# ---------------------------------------------------------------------------
def export_sheets_to_pdfs(wb, temp_dir: Path) -> list[Path]:
    pdfs = []
    for desired in SLIDE_ORDER:
        target = normalize_sheet_name(desired)
        found = None
        for sh in wb.Sheets:
            try:
                if normalize_sheet_name(sh.Name) == target:
                    found = sh
                    break
            except Exception:
                continue
        if not found:
            print(f"  ! sheet not found for '{desired}', skipping.")
            continue

        safe = "".join(c if c not in '\\/:*?"<>|' else "_" for c in found.Name).strip()
        safe = safe or f"slide_{len(pdfs) + 1}"
        pdf_path = temp_dir / f"{len(pdfs):02d}_{safe}.pdf"

        if FIT_SHEET_TO_ONE_PAGE:
            try:
                ps = found.PageSetup
                ps.Zoom = False            # required before FitToPages takes effect
                ps.FitToPagesWide = 1
                ps.FitToPagesTall = 1
                # Zero the page margins so content fills the page (kills the
                # big white border); CROP_TO_CONTENT removes the rest.
                ps.LeftMargin = 0
                ps.RightMargin = 0
                ps.TopMargin = 0
                ps.BottomMargin = 0
                ps.HeaderMargin = 0
                ps.FooterMargin = 0
                ps.CenterHorizontally = True
                ps.CenterVertically = True
            except Exception:
                pass

        try:
            found.ExportAsFixedFormat(
                Type=0,                  # 0 = PDF
                Filename=str(pdf_path),
                Quality=0,               # standard quality (keeps vectors)
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
        except Exception as e:
            print(f"  ! failed to export '{found.Name}': {e}")
            continue
        if pdf_path.exists():
            print(f"  + {found.Name}")
            pdfs.append(pdf_path)
    return pdfs


def build_deck(xlsx_path: Path, out_path: Path):
    import win32com.client as win32  # imported here so the merge core needs no pywin32

    try:
        excel = win32.GetObject(Class="Excel.Application")
    except Exception:
        excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.ScreenUpdating = False

    wb = None
    try:
        wb = excel.Workbooks.Open(str(xlsx_path.absolute()))
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            print("Exporting sheets to PDF ...")
            pdfs = export_sheets_to_pdfs(wb, td)
            if not pdfs:
                print("No sheets exported; aborting.")
                return None
            print("Merging into one vector PDF deck ...")
            written = merge_pdfs_to_deck(pdfs, out_path)
            print(f"Saved PDF deck: {written}")
            return written
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    here = Path(__file__).parent
    xlsx = here / "USDA_PSD.xlsx"
    out = None
    if len(sys.argv) > 1:
        xlsx = Path(sys.argv[1])
    if len(sys.argv) > 2:
        out = Path(sys.argv[2])
    if out is None:
        today = datetime.date.today().strftime("%Y-%m-%d")
        out = Path.home() / "Desktop" / f"AUS SnD Commodities 2026 - {today}.pdf"

    if not xlsx.exists():
        print(f"Error: workbook not found - {xlsx}")
        sys.exit(1)

    result = build_deck(xlsx, out)
    if result:
        try:
            os.startfile(str(result.parent))
        except Exception:
            pass
    else:
        sys.exit(1)
