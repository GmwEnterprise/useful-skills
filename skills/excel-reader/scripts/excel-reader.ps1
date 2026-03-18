#!/usr/bin/env pwsh
$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PackageDir = Join-Path (Split-Path $ScriptDir -Parent) "packages\excel-reader"

if ($args.Count -eq 0) {
    Write-Host "Usage: excel-reader.ps1 <excel_file> [sheet_name]"
    Write-Host ""
    Write-Host "Arguments:"
    Write-Host "  excel_file  - Path to the Excel file (.xlsx or .xlsm)"
    Write-Host "  sheet_name  - Optional: Name of specific sheet to read"
    Write-Host ""
    Write-Host "Output:"
    Write-Host "  JSON format with all sheets and cell data"
    exit 1
}

$ExcelFile = $args[0]
$SheetName = if ($args.Count -gt 1) { $args[1] } else { "" }

if (-not [System.IO.Path]::IsPathRooted($ExcelFile)) {
    $ExcelFile = (Resolve-Path -Path $ExcelFile -ErrorAction Stop).Path
}

Push-Location $PackageDir

if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    Write-Error "Error: uv is not installed. Please install uv first:"
    Write-Error "  irm https://astral.sh/uv/install.ps1 | iex"
    exit 1
}

uv sync --quiet
if ($SheetName) {
    uv run python main.py $ExcelFile $SheetName
} else {
    uv run python main.py $ExcelFile
}

Pop-Location
