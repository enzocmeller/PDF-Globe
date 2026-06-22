param(
  # Defaults to USDA_PSD.xlsx sitting next to this script. Override if yours differs.
  [string]$Workbook = (Join-Path $PSScriptRoot "USDA_PSD.xlsx")
)

if (-not (Test-Path $Workbook)) {
  Write-Host "No workbook found at: $Workbook"
  Write-Host "Skipping Excel refresh. Create it once (see README), name it USDA_PSD.xlsx,"
  Write-Host "and keep it in this folder."
  exit 0
}

Write-Host "Refreshing $Workbook ..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($Workbook)

  # Make all queries refresh synchronously so RefreshAll actually waits.
  foreach ($c in $wb.Connections) {
    try { $c.OLEDBConnection.BackgroundQuery = $false } catch {}
    try { $c.ODBCConnection.BackgroundQuery  = $false } catch {}
  }

  $wb.RefreshAll()
  # Power Query queries run async; this blocks until they finish.
  $excel.CalculateUntilAsyncQueriesDone()

  $wb.Save()
  $wb.Close($true)
  Write-Host "Excel refreshed and saved."
}
finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}