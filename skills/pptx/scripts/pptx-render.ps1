#!/usr/bin/env pwsh
# Export every slide of a .pptx as PNG via PowerPoint COM (best rendering
# fidelity). Requires desktop PowerPoint. For LibreOffice-only environments
# use the Bash version (scripts/pptx-render).
$ErrorActionPreference = "Stop"

if ($args.Count -lt 1 -or $args.Count -gt 3) {
    Write-Host "Usage: pptx-render.ps1 <pptx_file> [output_dir] [width_px]"
    Write-Host ""
    Write-Host "Arguments:"
    Write-Host "  pptx_file  - Path to the PowerPoint file (.pptx)"
    Write-Host "  output_dir - Optional: directory for slide PNGs (default: <source dir>\render)"
    Write-Host "  width_px   - Optional: export width in pixels (default: 1600)"
    exit 1
}

$OriginalPwd = Get-Location
$PptxFile = $args[0]
if (-not [System.IO.Path]::IsPathRooted($PptxFile)) {
    $PptxFile = Join-Path $OriginalPwd $PptxFile
}
if (-not (Test-Path $PptxFile)) {
    Write-Error "PowerPoint file not found: $PptxFile"
    exit 1
}
$OutDir = if ($args.Count -gt 1) { $args[1] } else { Join-Path (Split-Path -Parent $PptxFile) "render" }
if (-not [System.IO.Path]::IsPathRooted($OutDir)) {
    $OutDir = Join-Path $OriginalPwd $OutDir
}
$Width = if ($args.Count -gt 2) { [int]$args[2] } else { 1600 }
New-Item -ItemType Directory -Force $OutDir | Out-Null

# PowerPoint COM Open/Export fails (E_FAIL) on non-ASCII paths; stage through
# an ASCII temp directory and move results back.
$NonAscii = "[^\x00-\x7F]"
$OpenPath = $PptxFile
$ExportDir = $OutDir
$TempBase = $null
if (($PptxFile -match $NonAscii) -or ($OutDir -match $NonAscii)) {
    $TempBase = Join-Path ([System.IO.Path]::GetTempPath()) `
        ("pptx_render_" + [System.Guid]::NewGuid().ToString("N").Substring(0, 8))
    New-Item -ItemType Directory -Force $TempBase | Out-Null
    if ($PptxFile -match $NonAscii) {
        $OpenPath = Join-Path $TempBase "deck.pptx"
        Copy-Item $PptxFile $OpenPath
    }
    if ($OutDir -match $NonAscii) {
        $ExportDir = Join-Path $TempBase "out"
    }
}

$app = $null
try {
    $app = New-Object -ComObject PowerPoint.Application
    # ReadOnly=-1, Untitled=0, WithWindow=-1 (WithWindow=0 fails on some setups)
    $pres = $app.Presentations.Open($OpenPath, -1, 0, -1)
    # PageSetup dimensions are in points; 1 cm = 28.3465 pt
    $wPts = $pres.PageSetup.SlideWidth
    $hPts = $pres.PageSetup.SlideHeight
    $wCm = [Math]::Round($wPts / 28.3465, 2)
    $Height = [int][Math]::Round($Width * $hPts / $wPts)

    New-Item -ItemType Directory -Force $ExportDir | Out-Null
    $count = $pres.Slides.Count
    for ($i = 1; $i -le $count; $i++) {
        $pres.Slides.Item($i).Export(
            (Join-Path $ExportDir "slide$i.png"), 'PNG', $Width, $Height)
    }
    $pres.Close()

    if ($ExportDir -ne $OutDir) {
        Move-Item (Join-Path $ExportDir "slide*.png") $OutDir -Force
    }

    $ratio = [Math]::Round($Width / $wCm, 3)
    Write-Host "Success: $PptxFile"
    Write-Host "Slides: $count"
    Write-Host "PNG: $OutDir (slide1.png..slide$count.png, ${Width}x${Height}px)"
    Write-Host "Ratio: $ratio px/cm (${Width}px / ${wCm}cm) - use for cm<->px conversion"
}
finally {
    if ($app) {
        $app.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
    }
    if ($TempBase) {
        Remove-Item -Recurse -Force $TempBase -ErrorAction SilentlyContinue
    }
}
