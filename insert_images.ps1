# =============================================================================
# insert_images.ps1
# -----------------
# Inserts (or refreshes) the latest globe + El Nino PNGs into the
# "Globe & ElNino" worksheet of USDA_PSD.xlsx via Excel COM automation.
#
# Why COM and not openpyxl?
#   USDA_PSD.xlsx contains charts, Power Query / queryTables, tables,
#   rich-data web images, customXml, and a classification label. openpyxl
#   cannot safely round-trip those. Driving real Excel preserves every
#   feature because Excel itself does the save.
#
# Idempotent: each run deletes the shapes named CR_Globe / CPC_ElNino if
# present, then re-inserts. Other shapes (charts, logos) are untouched.
#
# Inputs (must already exist in the same folder as this script):
#   - USDA_PSD.xlsx
#   - map.png      (from "final map.py" / Update_Map.py)
#   - elnino.png   (from final_elnino.py)
# =============================================================================

param(
    [string]$Workbook = (Join-Path $PSScriptRoot "USDA_PSD.xlsx"),
    [string]$MapPng   = (Join-Path $PSScriptRoot "map.png"),
    [string]$NinoPng  = (Join-Path $PSScriptRoot "elnino.png")
)

# -----------------------------------------------------------------------------
# CONFIG -- adjust here after the first run.
# -----------------------------------------------------------------------------
$SheetName       = "Globe & ElNino"

$GlobeAnchorCell = "AU80"
$NinoAnchorCell  = "BC80"

$GlobeShapeName  = "CR_Globe"
$NinoShapeName   = "CPC_ElNino"

# Starting small per your request. Width in *points* (1 pt = 1/72 inch).
# Excel works in points; height auto-adjusts from aspect ratio. Bump these
# after the first run if you want larger.
$GlobeWidthPt    = 180     # ~2.5 inches wide
$NinoWidthPt     = 220     # ~3.05 inches wide

# Set $true on first run if you want to also remove the older untitled
# images that currently sit at AS74 (800x774) and AX74 (1100x850). After
# they're gone, set this back to $false so subsequent runs only touch
# the named shapes.
$DeleteLegacyPositionalImages = $false


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
function Fail($msg, $code = 1) {
    Write-Host "ERROR: $msg" -ForegroundColor Red
    exit $code
}

function Test-FileLocked($path) {
    if (-not (Test-Path $path)) { return $false }
    try {
        $fs = [System.IO.File]::Open($path, 'Open', 'ReadWrite', 'None')
        $fs.Close(); $fs.Dispose()
        return $false
    } catch {
        return $true
    }
}

# -----------------------------------------------------------------------------
# Pre-flight checks
# -----------------------------------------------------------------------------
if (-not (Test-Path $Workbook)) { Fail "Workbook not found: $Workbook" 2 }
if (-not (Test-Path $MapPng))   { Fail "Globe PNG not found: $MapPng (run 'final map.py' first)" 3 }
if (-not (Test-Path $NinoPng))  { Fail "El Nino PNG not found: $NinoPng (run final_elnino.py first)" 3 }

if (Test-FileLocked $Workbook) {
    Fail "Workbook appears to be open in Excel (file locked). Close it and try again: $Workbook" 4
}

# Resolve to absolute paths -- AddPicture requires an absolute path.
$Workbook = (Resolve-Path $Workbook).Path
$MapPng   = (Resolve-Path $MapPng).Path
$NinoPng  = (Resolve-Path $NinoPng).Path

Write-Host "Workbook : $Workbook"
Write-Host "Globe PNG: $MapPng"
Write-Host "Nino  PNG: $NinoPng"
Write-Host "Sheet    : $SheetName"
Write-Host "Anchors  : $GlobeAnchorCell (globe), $NinoAnchorCell (El Nino)"
Write-Host ""

# -----------------------------------------------------------------------------
# Drive Excel
# -----------------------------------------------------------------------------
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false

# Office constants we need (so the script works without referencing the
# Excel/Office object library).
$msoTrue                  = -1
$msoFalse                 =  0
$msoPicture               = 13
$msoCTrue                 = -1
$xlPicture                = -4147   # not used; AddPicture path is fine

$wb = $null
try {
    Write-Host "Opening workbook ..."
    $wb = $excel.Workbooks.Open($Workbook, 0, $false)

    # Find the target worksheet.
    $ws = $null
    foreach ($s in $wb.Worksheets) {
        if ($s.Name -eq $SheetName) { $ws = $s; break }
    }
    if ($null -eq $ws) {
        $names = ($wb.Worksheets | ForEach-Object { $_.Name }) -join ", "
        throw "Worksheet '$SheetName' not found. Sheets present: $names"
    }

    # ------------------------------------------------------------------
    # Optional one-time cleanup of legacy positional images.
    # ------------------------------------------------------------------
    if ($DeleteLegacyPositionalImages) {
        Write-Host "Removing legacy positional images (col AS / AX, row 74) ..."
        # Iterate by index in reverse so deletions don't shift the iterator.
        for ($i = $ws.Shapes.Count; $i -ge 1; $i--) {
            $shp = $ws.Shapes.Item($i)
            try {
                # Only act on actual pictures (type 13), and only ones
                # anchored very close to AS74 or AX74.
                if ($shp.Type -eq $msoPicture) {
                    $tlc = $shp.TopLeftCell
                    if ($tlc.Row -eq 74 -and ($tlc.Column -eq 45 -or $tlc.Column -eq 50)) {
                        Write-Host "  deleting legacy picture '$($shp.Name)' at $($tlc.Address($false,$false))"
                        $shp.Delete()
                    }
                }
            } catch { }
        }
    }

    # ------------------------------------------------------------------
    # Helper: delete any shape with a given name (idempotent insert).
    # ------------------------------------------------------------------
    function Remove-ShapeByName($worksheet, $name) {
        for ($i = $worksheet.Shapes.Count; $i -ge 1; $i--) {
            $shp = $worksheet.Shapes.Item($i)
            try {
                if ($shp.Name -eq $name) {
                    Write-Host "  removed existing shape '$name'"
                    $shp.Delete()
                }
            } catch { }
        }
    }

    # ------------------------------------------------------------------
    # Helper: insert picture at a cell anchor, aspect-locked.
    # ------------------------------------------------------------------
    function Insert-Picture($worksheet, $pngPath, $anchorCell, $shapeName, $widthPt) {
        Write-Host "Inserting '$shapeName' from $([System.IO.Path]::GetFileName($pngPath)) at $anchorCell ..."

        Remove-ShapeByName $worksheet $shapeName

        $anchor = $worksheet.Range($anchorCell)
        $left   = $anchor.Left
        $top    = $anchor.Top

        # AddPicture(Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
        # Pass -1 for Width/Height to use native, then resize while preserving ratio.
        $shp = $worksheet.Shapes.AddPicture(
            $pngPath,
            $msoFalse,        # LinkToFile = no -> embed bytes
            $msoCTrue,        # SaveWithDocument = yes
            $left,
            $top,
            -1,               # native width
            -1                # native height
        )
        $shp.Name = $shapeName
        $shp.LockAspectRatio = $msoTrue
        $shp.Width = [double]$widthPt   # height follows aspect ratio
        # Pin top-left exactly to the anchor cell (AddPicture sometimes
        # nudges by sub-pixel rounding).
        $shp.Left = $left
        $shp.Top  = $top

        Write-Host "  -> placed at $anchorCell  size = $([math]::Round($shp.Width,1)) x $([math]::Round($shp.Height,1)) pt"
    }

    # ------------------------------------------------------------------
    # Do the work
    # ------------------------------------------------------------------
    Insert-Picture $ws $MapPng  $GlobeAnchorCell $GlobeShapeName $GlobeWidthPt
    Insert-Picture $ws $NinoPng $NinoAnchorCell  $NinoShapeName  $NinoWidthPt

    Write-Host ""
    Write-Host "Saving workbook ..."
    $wb.Save()
    Write-Host "Done."
}
catch {
    Write-Host ""
    Fail $_.Exception.Message 5
}
finally {
    if ($null -ne $wb) {
        try { $wb.Close($true) } catch { }
    }
    try { $excel.Quit() } catch { }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
