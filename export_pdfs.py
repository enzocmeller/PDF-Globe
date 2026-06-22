import datetime
import os
import re
import sys
import time
from pathlib import Path
import tempfile
import fitz
from PIL import Image
import win32com.client as win32
import pywintypes


def normalize_sheet_name(name: str) -> str:
    normalized = str(name or "").strip().lower()
    normalized = re.sub(r'\b(page|sheet)\b', '', normalized)
    normalized = re.sub(r'[^a-z0-9]+', '', normalized)
    return normalized


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
]

# Options: 'match_pdf' = set slide size to PDF page size (not recommended if pages vary)
#          'cover' = scale image to cover entire slide (may crop edges)
#          'contain' = scale image to fit inside slide (no cropping) — recommended when you need all content visible
SLIDE_FILL_MODE = "contain"


def export_excel_to_ppt(excel_file_path, ppt_output_path=None):
    """Create a PowerPoint with slides in the specified order using Excel sheets.

    Flow:
      - Export each sheet in SLIDE_ORDER to a PDF (single temp folder)
      - Measure PDF page sizes and choose a presentation canvas size (max width/height)
      - Render each PDF page to a high-DPI PNG, auto-crop whitespace, then insert into a slide
      - Use contain scaling by default so all content stays visible
    """
    if ppt_output_path is None:
        desktop = Path.home() / "Desktop"
        today = datetime.date.today().strftime("%Y-%m-%d")
        ppt_output_path = desktop / f"AUS SnD Commodities 2026 - {today}.pptx"
    else:
        ppt_output_path = Path(ppt_output_path)

    excel_file_path = Path(excel_file_path)
    if not excel_file_path.exists():
        raise FileNotFoundError(str(excel_file_path))

    # Start Excel
    try:
        excel = win32.GetObject(Class="Excel.Application")
    except Exception:
        excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.ScreenUpdating = False

    # Start PowerPoint
    try:
        ppt = win32.GetObject(Class="PowerPoint.Application")
    except Exception:
        ppt = win32.Dispatch("PowerPoint.Application")
    ppt.Visible = True

    wb = None
    pres = None
    try:
        wb = excel.Workbooks.Open(str(excel_file_path.absolute()))

        jobs = []  # (sheet_name, pdf_path, page_w_pts, page_h_pts)
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_dir = Path(temp_dir)
            # Export PDFs
            for desired in SLIDE_ORDER:
                found = None
                desired_norm = normalize_sheet_name(desired)
                for sh in wb.Sheets:
                    try:
                        sheet_name = sh.Name
                        sheet_norm = normalize_sheet_name(sheet_name)
                        if sheet_norm == desired_norm:
                            found = sh
                            break
                    except Exception:
                        continue

                if not found:
                    print(f"Warning: sheet not found for '{desired}', skipping.")
                    continue

                safe_name = "".join(c if c not in "\\/:*?\"<>|" else "_" for c in found.Name).strip()
                if not safe_name:
                    safe_name = f"slide_{len(jobs)+1}"
                pdf_path = temp_dir / f"{safe_name}.pdf"

                try:
                    found.ExportAsFixedFormat(
                        Type=0,
                        Filename=str(pdf_path),
                        Quality=0,
                        IncludeDocProperties=True,
                        IgnorePrintAreas=False,
                        OpenAfterPublish=False,
                    )
                except Exception as e:
                    print(f"Failed to export '{found.Name}' to PDF: {e}")
                    continue

                if not pdf_path.exists():
                    print(f"PDF not created for '{found.Name}'")
                    continue

                try:
                    with fitz.open(str(pdf_path)) as doc:
                        page = doc[0]
                        page_w = page.rect.width
                        page_h = page.rect.height
                except Exception as e:
                    print(f"Failed reading PDF page for '{found.Name}': {e}")
                    continue

                jobs.append((found.Name, pdf_path, page_w, page_h))

            if not jobs:
                print("No PDFs exported; aborting.")
                return

            # Choose a slide size that minimizes total blank area across all pages.
            # This keeps all content visible while reducing the empty borders.
            best_size = None
            best_blank = None
            for candidate_w, candidate_h in [(p[2], p[3]) for p in jobs]:
                total_blank = 0.0
                for _, _, page_w, page_h in jobs:
                    scale = min(candidate_w / page_w, candidate_h / page_h)
                    used_area = page_w * page_h * (scale ** 2)
                    total_blank += candidate_w * candidate_h - used_area
                if best_blank is None or total_blank < best_blank:
                    best_blank = total_blank
                    best_size = (candidate_w, candidate_h)

            slide_width, slide_height = best_size
            try:
                pres = ppt.Presentations.Add()
                pres.PageSetup.SlideWidth = slide_width
                pres.PageSetup.SlideHeight = slide_height
            except Exception:
                pres = ppt.Presentations.Add()

            dpi = 400
            # Insert slides with the selected scaling mode
            for idx, (sheet_name, pdf_path, page_w, page_h) in enumerate(jobs, start=1):
                slide = pres.Slides.Add(idx, 12)
                png_path = temp_dir / f"{sheet_name}.png"
                try:
                    with fitz.open(str(pdf_path)) as doc:
                        page = doc[0]
                        zoom = dpi / 72.0
                        mat = fitz.Matrix(zoom, zoom)
                        pix = page.get_pixmap(matrix=mat, alpha=False)
                        pix.save(str(png_path))
                except Exception as e:
                    print(f"Failed to render PDF for '{sheet_name}' to image: {e}")
                    continue

                # Auto-crop
                try:
                    img = Image.open(png_path).convert('RGB')
                    gray = img.convert('L')
                    bw = gray.point(lambda p: 0 if p > 250 else 255, mode='L')
                    bbox = bw.getbbox()
                    if bbox:
                        img = img.crop(bbox)
                        img.save(str(png_path))
                except Exception as e:
                    print(f"Warning: could not auto-crop image for '{sheet_name}': {e}")

                img_px_w, img_px_h = img.size
                image_width = img_px_w * 72.0 / dpi
                image_height = img_px_h * 72.0 / dpi

                slide_w = pres.PageSetup.SlideWidth
                slide_h = pres.PageSetup.SlideHeight

                if SLIDE_FILL_MODE == 'match_pdf':
                    # (not recommended when pages vary) resize presentation to page size
                    try:
                        pres.PageSetup.SlideWidth = page_w
                        pres.PageSetup.SlideHeight = page_h
                        slide_w = page_w
                        slide_h = page_h
                    except Exception:
                        pass

                # Choose scaling based on mode
                if SLIDE_FILL_MODE == 'contain':
                    # allow upscaling so the image fills the slide as much as possible
                    scale = min(slide_w / image_width, slide_h / image_height)
                else:
                    scale = max(slide_w / image_width, slide_h / image_height)

                target_w = image_width * scale
                target_h = image_height * scale

                pic = slide.Shapes.AddPicture(
                    FileName=str(png_path),
                    LinkToFile=False,
                    SaveWithDocument=True,
                    Left=0,
                    Top=0,
                    Width=target_w,
                    Height=target_h,
                )
                try:
                    pic.LockAspectRatio = -1
                except Exception:
                    pass
                pic.Left = int((slide_w - pic.Width) / 2)
                pic.Top = int((slide_h - pic.Height) / 2)

            if SLIDE_FILL_MODE == 'cover':
                # Final pass only for cover mode; contain mode must preserve all content
                try:
                    slide_w = pres.PageSetup.SlideWidth
                    slide_h = pres.PageSetup.SlideHeight
                    for s in pres.Slides:
                        for i in range(1, s.Shapes.Count + 1):
                            shp = s.Shapes.Item(i)
                            try:
                                if shp.Type == 13:  # msoPicture
                                    img_w = shp.Width
                                    img_h = shp.Height
                                    if img_w <= 0 or img_h <= 0:
                                        continue
                                    scale = max(slide_w / img_w, slide_h / img_h)
                                    new_w = img_w * scale
                                    new_h = img_h * scale
                                    try:
                                        shp.LockAspectRatio = -1
                                    except Exception:
                                        pass
                                    shp.Width = new_w
                                    shp.Height = new_h
                                    shp.Left = int((slide_w - new_w) / 2)
                                    shp.Top = int((slide_h - new_h) / 2)
                            except Exception:
                                continue
                except Exception:
                    pass

            # Save presentation with fallback
            try:
                pres.SaveAs(str(ppt_output_path))
                print(f"Saved PowerPoint: {ppt_output_path}")
            except pywintypes.com_error as e:
                try:
                    ts = int(time.time())
                    alt_path = ppt_output_path.with_name(ppt_output_path.stem + f"_{ts}" + ppt_output_path.suffix)
                    pres.SaveAs(str(alt_path))
                    print(f"Target locked; saved PowerPoint to: {alt_path}")
                    ppt_output_path = alt_path
                except Exception:
                    print(f"Failed to save presentation: {e}")

            # Open folder
            try:
                os.startfile(str(ppt_output_path.parent))
            except Exception:
                pass

    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
        try:
            ppt.Quit()
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    current_dir = Path(__file__).parent
    excel_file = current_dir / "USDA_PSD.xlsx"
    ppt_output = None

    if len(sys.argv) > 1:
        excel_file = Path(sys.argv[1])
    if len(sys.argv) > 2:
        ppt_output = Path(sys.argv[2])

    if not excel_file.exists():
        print("Error: File not found - " + str(excel_file))
        sys.exit(1)

    export_excel_to_ppt(str(excel_file), str(ppt_output) if ppt_output else None)
