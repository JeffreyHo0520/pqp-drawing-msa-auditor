# ==============================================================================
# PQP Drawing & MSA Quality Inspection Suite - Windows Installer
# Supported Environments: Antigravity / Gemini CLI / Claude Code / Cursor / Codex
# Skills: pqp-drawing-msa-auditor & drawing-spec-pqp-comparison
# ==============================================================================

param(
    [string]$TargetDir = "",
    [switch]$Workspace,
    [switch]$Claude
)

$ErrorActionPreference = "Stop"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "  PQP Drawing & MSA Quality Inspection Suite Installer    " -ForegroundColor Yellow
Write-Host "==========================================================" -ForegroundColor Cyan

# Determine installation directory
if ($TargetDir -ne "") {
    $destBase = $TargetDir
} elseif ($Workspace) {
    $destBase = Join-Path (Get-Location) ".agents\skills"
} elseif ($Claude) {
    $destBase = Join-Path $HOME ".claude\skills"
} else {
    # Default Antigravity / Gemini CLI global skills directory
    $destBase = Join-Path $HOME ".gemini\config\skills"
}

Write-Host "[1/3] Destination: $destBase" -ForegroundColor Gray

if (-not (Test-Path $destBase)) {
    New-Item -ItemType Directory -Path $destBase -Force | Out-Null
}

$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) {
    $scriptRoot = (Get-Location).Path
}

$skills = @(
    "pqp-drawing-msa-auditor",
    "drawing-spec-pqp-comparison"
)

Write-Host "[2/3] Installing skills..." -ForegroundColor Gray

foreach ($skill in $skills) {
    $srcPath = Join-Path $scriptRoot $skill
    $targetPath = Join-Path $destBase $skill

    if (-not (Test-Path $srcPath)) {
        Write-Error "Source directory not found: $srcPath"
        continue
    }

    if (Test-Path $targetPath) {
        Write-Host "  -> Updating existing: $skill" -ForegroundColor DarkYellow
        Remove-Item -Recurse -Force $targetPath
    }

    Copy-Item -Recurse -Path $srcPath -Destination $targetPath -Force
    Write-Host "  + Installed: $skill" -ForegroundColor Green
}

# Verification
Write-Host "[3/3] Verifying installation..." -ForegroundColor Gray
$allOk = $true
foreach ($skill in $skills) {
    $skillMd = Join-Path $destBase "$skill\SKILL.md"
    if (Test-Path $skillMd) {
        Write-Host "  + Verified: $skillMd" -ForegroundColor Green
    } else {
        Write-Host "  x Verification failed: $skillMd missing" -ForegroundColor Red
        $allOk = $false
    }
}

Write-Host "==========================================================" -ForegroundColor Cyan
if ($allOk) {
    Write-Host "SUCCESS: PQP Drawing & MSA Skills installed successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "How to use in Antigravity / Claude Code / Codex:" -ForegroundColor Yellow
    Write-Host "  1. Slash commands:" -ForegroundColor White
    Write-Host "     /pqp-drawing-msa-auditor" -ForegroundColor Cyan
    Write-Host "     /drawing-spec-pqp-comparison" -ForegroundColor Cyan
    Write-Host "  2. CLI direct execution via uv:" -ForegroundColor White
    Write-Host "     uv run ~/.gemini/config/skills/pqp-drawing-msa-auditor/scripts/pqp_dwg_msa_auditor.py all --pqp PQP.xlsx --dwgs Drawing.pdf" -ForegroundColor Cyan
} else {
    Write-Host "WARNING: Some skill files could not be verified." -ForegroundColor Red
}
Write-Host "==========================================================" -ForegroundColor Cyan
