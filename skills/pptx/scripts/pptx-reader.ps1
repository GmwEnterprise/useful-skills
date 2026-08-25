#!/usr/bin/env pwsh
$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PackageDir = Join-Path (Split-Path $ScriptDir -Parent) "packages\pptx"

$OriginalPwd = Get-Location

if ($args.Count -eq 0) {
    Write-Host "Usage: pptx-reader.ps1 <pptx_file> [output_dir]"
    Write-Host ""
    Write-Host "Arguments:"
    Write-Host "  pptx_file  - Path to the PowerPoint file (.pptx)"
    Write-Host "  output_dir - Optional: directory for output files (default: same dir as input)"
    Write-Host ""
    Write-Host "Output:"
    Write-Host "  Markdown: <output_dir>/<filename>.pptx_reader.md"
    Write-Host "  JSON:     <output_dir>/<filename>.pptx_reader.json"
    exit 1
}

$PptxFile = $args[0]
$OutputDir = if ($args.Count -gt 1) { $args[1] } else { "" }

if (-not [System.IO.Path]::IsPathRooted($PptxFile)) {
    $PptxFile = Join-Path $OriginalPwd $PptxFile
}

if ($OutputDir -and -not [System.IO.Path]::IsPathRooted($OutputDir)) {
    $OutputDir = Join-Path $OriginalPwd $OutputDir
}

Push-Location $PackageDir

if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    Write-Error "Error: uv is not installed. Please install uv first:"
    Write-Error "  irm https://astral.sh/uv/install.ps1 | iex"
    Pop-Location
    exit 1
}

uv sync --quiet
if ($OutputDir) {
    uv run python main.py $PptxFile $OutputDir
} else {
    uv run python main.py $PptxFile
}
$Code = $LASTEXITCODE

Pop-Location
exit $Code
