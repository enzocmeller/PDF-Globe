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
#   - map.png      (from "final map.py")
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

# Each image is sized to FILL the cell range below (preserving aspect ratio)
# and centered inside it. Left half of the slide = globe, right half = El Nino.
# These ranges sit below the two captions (row 72) and above the footer (row 91).
$GlobeRange      = "AT74:AY89"
$NinoRange       = "AZ74:BG89"

$GlobeShapeName  = "CR_Globe"
$NinoShapeName   = "CPC_ElNino"

# Fraction of the range each image fills (1.0 = touch the range edges).
$FillFactor      = 0.97

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
Write-Host "Ranges   : $GlobeRange (globe), $NinoRange (El Nino)"
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
    # Helper: insert a picture sized to FILL a cell range (aspect-locked)
    # and centered inside it.
    # ------------------------------------------------------------------
    function Insert-PictureInRange($worksheet, $pngPath, $rangeAddr, $shapeName, $fill) {
        Write-Host "Inserting '$shapeName' from $([System.IO.Path]::GetFileName($pngPath)) into $rangeAddr ..."

        Remove-ShapeByName $worksheet $shapeName

        $rng = $worksheet.Range($rangeAddr)
        $rL = $rng.Left; $rT = $rng.Top; $rW = $rng.Width; $rH = $rng.Height

        # AddPicture(Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
        # Insert at native size (-1,-1) first so we can read the true aspect.
        $shp = $worksheet.Shapes.AddPicture(
            $pngPath, $msoFalse, $msoCTrue, $rL, $rT, -1, -1)
        $shp.Name = $shapeName
        $shp.LockAspectRatio = $msoTrue
        $nativeW = $shp.Width
        $nativeH = $shp.Height

        # Scale to fill the range (contain), preserving aspect.
        $availW = $rW * $fill
        $availH = $rH * $fill
        $scale  = [math]::Min($availW / $nativeW, $availH / $nativeH)
        $newW   = $nativeW * $scale
        $shp.Width = [double]$newW         # height follows via locked aspect
        $newH   = $shp.Height

        # Center within the range, both axes.
        $shp.Left = $rL + ($rW - $newW) / 2.0
        $shp.Top  = $rT + ($rH - $newH) / 2.0

        Write-Host ("  -> {0}: {1:N0} x {2:N0} pt, centered in {3}" -f `
                    $shapeName, $newW, $newH, $rangeAddr)
    }

    # ------------------------------------------------------------------
    # Do the work
    # ------------------------------------------------------------------
    Insert-PictureInRange $ws $MapPng  $GlobeRange $GlobeShapeName $FillFactor
    Insert-PictureInRange $ws $NinoPng $NinoRange  $NinoShapeName  $FillFactor

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
