param([Parameter(Mandatory=$true)][string]$InputPath,
      [Parameter(Mandatory=$true)][string]$OutputPath,
      [Parameter(Mandatory=$true)][string]$PdfPath)
$ErrorActionPreference='Stop'
$source=(Resolve-Path -LiteralPath $InputPath).Path
$dest=[IO.Path]::GetFullPath($OutputPath)
$pdf=[IO.Path]::GetFullPath($PdfPath)
if ((Test-Path -LiteralPath $dest) -or (Test-Path -LiteralPath $pdf)) { throw 'Use new output paths.' }
$hwpBudget=New-Object -ComObject HWPFrame.HwpObject
try {
    $hwpBudget.XHwpWindows.Item(0).Visible=$false
    if (-not $hwpBudget.Open($source,'','')) { throw 'Hancom open failed.' }
    if (-not $hwpBudget.SaveAs($dest,'HWPX','')) { throw 'HWPX save failed.' }
    if (-not $hwpBudget.SaveAs($pdf,'PDF','')) { throw 'PDF save failed.' }
    @{hwpx=$dest;pdf=$pdf;pages=$hwpBudget.PageCount} | ConvertTo-Json
} finally {
    $hwpBudget.Clear(1)
    $hwpBudget.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($hwpBudget)
}
