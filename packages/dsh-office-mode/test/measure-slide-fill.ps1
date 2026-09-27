# Measure how much of each rendered slide's body area carries no text.
#
# Purpose: the deck engine sizes helper containers (cards / kpi / steps / ...) to
# fill the body box, so pages whose text is one or two lines end up mostly empty.
# This script turns "looks empty" into a number, per slide.
#
# Usage:
#   powershell -File test/measure-slide-fill.ps1 -Dir <rendered-png-dir> [-DarkThreshold 120] [-Step 6]
#
# Rendering the PNGs in the first place: test/export-pptx-png.vbs
#
# NOTE: this file is deliberately ASCII-only. Windows PowerShell 5.1 reads a
# UTF-8 script *without* a BOM as ANSI, which would garble CJK comments and
# string literals. Keep it ASCII unless you also add a BOM.

param(
    [Parameter(Mandatory = $true)][string]$Dir,
    [int]$DarkThreshold = 120,
    [int]$Step = 6
)

Add-Type -AssemblyName System.Drawing

$bands = 40           # vertical resolution: each band is 2.5% of the page height
$bodyFrom = 5         # 12.5%  (skip the header band)
$bodyTo = 36          # 92.5%  (skip the footer band)

$rows = @()
foreach ($file in Get-ChildItem $Dir -Filter 'slide-*.png' | Sort-Object Name) {
    $bmp = [System.Drawing.Bitmap]::FromFile($file.FullName)
    $w = $bmp.Width
    $h = $bmp.Height
    $dark = New-Object 'int[]' $bands
    $total = New-Object 'int[]' $bands
    for ($y = 0; $y -lt $h; $y += $Step) {
        $band = [int][Math]::Floor($y * $bands / $h)
        if ($band -gt $bands - 1) { $band = $bands - 1 }
        for ($x = 0; $x -lt $w; $x += $Step) {
            $p = $bmp.GetPixel($x, $y)
            $lum = 0.299 * $p.R + 0.587 * $p.G + 0.114 * $p.B
            $total[$band]++
            if ($lum -lt $DarkThreshold) { $dark[$band]++ }
        }
    }
    $bmp.Dispose()

    $empty = 0
    $run = 0
    $longest = 0
    for ($b = $bodyFrom; $b -le $bodyTo; $b++) {
        $ratio = if ($total[$b] -gt 0) { $dark[$b] / $total[$b] } else { 0 }
        if ($ratio -lt 0.002) {
            $empty++
            $run++
            if ($run -gt $longest) { $longest = $run }
        } else {
            $run = 0
        }
    }
    $rows += [pscustomobject]@{
        Slide = $file.BaseName
        EmptyBands = $empty
        BodyBands = ($bodyTo - $bodyFrom + 1)
        LongestEmptyRun = $longest
    }
}

$rows | Format-Table -AutoSize | Out-String -Width 80 | Write-Host
$half = $rows | Where-Object { $_.EmptyBands -ge ($_.BodyBands / 2) }
Write-Host ("slides with >= 50% empty body: {0} / {1}" -f $half.Count, $rows.Count)
if ($half.Count -gt 0) {
    Write-Host ("  " + (($half | ForEach-Object { $_.Slide }) -join ', '))
}
