#!/usr/bin/env pwsh
$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PackageDir = Join-Path (Split-Path $ScriptDir -Parent) "packages\pptx"

$OriginalPwd = Get-Location

if ($args.Count -lt 2) {
    Write-Host "Usage: pptx-icon.ps1 <out_dir> <icon_name>... [--color RRGGBB] [--size N] [--prefix set]"
    Write-Host ""
    Write-Host "Arguments:"
    Write-Host "  out_dir     - output directory for PNG files"
    Write-Host "  icon_name   - one or more iconify icon names (e.g. alert target)"
    Write-Host ""
    Write-Host "Options:"
    Write-Host "  --color RRGGBB  tint color (default FFFFFF)"
    Write-Host "  --size N        square canvas px (default 256)"
    Write-Host "  --prefix set    iconify prefix (default mdi)"
    exit 1
}

$OutDir = $args[0]
$Rest = $args[1..($args.Count - 1)]

if (-not [System.IO.Path]::IsPathRooted($OutDir)) {
    $OutDir = Join-Path $OriginalPwd $OutDir
}

Push-Location $PackageDir

if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    Write-Error "Error: uv is not installed. Please install uv first:"
    Write-Error "  irm https://astral.sh/uv/install.ps1 | iex"
    Pop-Location
    exit 1
}

uv sync --quiet
uv run python icon.py $OutDir @Rest
$Code = $LASTEXITCODE
Pop-Location
exit $Code
