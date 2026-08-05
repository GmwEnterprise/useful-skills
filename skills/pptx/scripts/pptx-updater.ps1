#!/usr/bin/env pwsh
$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PackageDir = Join-Path (Split-Path $ScriptDir -Parent) "packages\pptx"

$OriginalPwd = Get-Location

if ($args.Count -lt 2) {
    Write-Host "Usage: pptx-updater.ps1 <pptx_file> <changes.json> [output.pptx]"
    Write-Host ""
    Write-Host "Arguments:"
    Write-Host "  pptx_file    - Path to the source PowerPoint file (.pptx)"
    Write-Host "  changes.json - JSON file describing update operations"
    Write-Host "  output.pptx  - Optional: output path (default: <source>_updated.pptx)"
    exit 1
}

$PptxFile = $args[0]
$ChangesFile = $args[1]
$OutputFile = if ($args.Count -gt 2) { $args[2] } else { "" }

if (-not [System.IO.Path]::IsPathRooted($PptxFile)) {
    $PptxFile = Join-Path $OriginalPwd $PptxFile
}

if (-not [System.IO.Path]::IsPathRooted($ChangesFile)) {
    $ChangesFile = Join-Path $OriginalPwd $ChangesFile
}

if ($OutputFile -and -not [System.IO.Path]::IsPathRooted($OutputFile)) {
    $OutputFile = Join-Path $OriginalPwd $OutputFile
}

Push-Location $PackageDir

if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    Write-Error "Error: uv is not installed. Please install uv first:"
    Write-Error "  irm https://astral.sh/uv/install.ps1 | iex"
    Pop-Location
    exit 1
}

uv sync --quiet
if ($OutputFile) {
    uv run python updater.py $PptxFile $ChangesFile $OutputFile
} else {
    uv run python updater.py $PptxFile $ChangesFile
}

Pop-Location
