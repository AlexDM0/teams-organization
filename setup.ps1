<#
.SYNOPSIS
  setup.ps1 - install the `teams-org` CLI on Windows using native PowerShell.

.DESCRIPTION
  A PowerShell-native alternative to setup.sh for Windows users who don't have a
  bash shell (Git Bash / MSYS2 / WSL). It verifies Bun is installed and working,
  installs dependencies, links the CLI into Bun's global bin directory, and makes
  sure that directory is on your user PATH.

  Works in both Windows PowerShell 5.1 and PowerShell 7+. This file is
  deliberately pure ASCII so it parses correctly under 5.1 even without a BOM.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File .\setup.ps1
#>

# NOTE: we deliberately do NOT set $ErrorActionPreference = 'Stop' globally.
# Under Windows PowerShell 5.1, native commands (bun.exe) that write progress to
# stderr get wrapped as terminating errors when EAP is Stop, which would abort
# the script on a perfectly healthy `bun install`. Instead we check
# $LASTEXITCODE after each native call and wrap .NET calls in try/catch.

# Run from the folder this script lives in (guard against being dot-sourced or
# pasted via iex, where $PSScriptRoot is empty).
if ($PSScriptRoot) { Set-Location -Path $PSScriptRoot }

# --- Pretty output helpers -------------------------------------------------
# Glyphs are built from code points at runtime so the source stays pure ASCII
# (a BOM-less .ps1 with literal non-ASCII would render as mojibake under 5.1).
$script:BulletGlyph = [char]0x2022  # bullet
$script:CheckGlyph  = [char]0x2713  # check mark
$script:CrossGlyph  = [char]0x2717  # ballot X

function Write-Bold($message) { Write-Host $message -ForegroundColor White }
function Write-Info($message) { Write-Host "$script:BulletGlyph $message" -ForegroundColor Cyan }
function Write-Ok($message)   { Write-Host "$script:CheckGlyph $message" -ForegroundColor Green }
function Write-Warn($message) { Write-Host "! $message" -ForegroundColor Yellow }
function Write-Err($message)   { Write-Host "$script:CrossGlyph $message" -ForegroundColor Red }

Write-Bold 'Setting up the teams-org CLI (windows)'
Write-Host ''

function Show-BunInstallHelp {
    Write-Host ''
    Write-Err 'Bun is required but was not found (or is not working).'
    Write-Host ''
    Write-Info 'Install Bun on Windows (run in PowerShell):'
    Write-Host '    powershell -c "irm bun.sh/install.ps1 | iex"'
    Write-Host ''
    Write-Info 'More info: https://bun.sh/'
    Write-Host ''
    Write-Warn 'After installing, open a new terminal (so PATH updates) and re-run .\setup.ps1'
}

# --- 1. Verify Bun is installed AND working --------------------------------
Write-Info 'Checking for Bun...'
if (-not (Get-Command bun -ErrorAction SilentlyContinue)) {
    Show-BunInstallHelp
    exit 1
}

# Confirm it actually runs. Merge stderr and rely on $LASTEXITCODE - a broken
# install exits non-zero (or fails to launch, landing in catch).
$bunVersion = $null
try {
    $bunOutput = (& bun --version 2>&1 | Out-String)
    if ($LASTEXITCODE -eq 0 -and $bunOutput) {
        $bunVersion = ($bunOutput -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -First 1).Trim()
    }
} catch { }

if (-not $bunVersion) {
    Write-Err "Found 'bun' on PATH but it failed to run - the install looks broken."
    Show-BunInstallHelp
    exit 1
}
Write-Ok "Bun $bunVersion is installed and working."

# --- 2. Install dependencies ----------------------------------------------
Write-Info 'Installing dependencies (bun install)...'
& bun install
if ($LASTEXITCODE -ne 0) { Write-Err 'bun install failed.'; exit 1 }
Write-Ok 'Dependencies installed.'

# --- 3. Link the CLI globally ---------------------------------------------
Write-Info 'Linking the teams-org CLI (bun link)...'
& bun link
if ($LASTEXITCODE -ne 0) { Write-Err 'bun link failed.'; exit 1 }
Write-Ok 'CLI linked.'

# --- 4. Ensure Bun's bin directory is on the user PATH --------------------
# Bun places global bins in $env:BUN_INSTALL\bin (defaults to %USERPROFILE%\.bun\bin).
$bunInstallRoot = if ($env:BUN_INSTALL) { $env:BUN_INSTALL } else { Join-Path $env:USERPROFILE '.bun' }
$bunBinDirectory = Join-Path $bunInstallRoot 'bin'

# Persistently add it to the *user* PATH if it isn't already there.
# [Environment]::SetEnvironmentVariable also broadcasts WM_SETTINGCHANGE, so a
# newly-opened terminal picks up the change without a full sign-out (a raw
# registry write would not). It preserves your existing PATH length (unlike
# setx, which truncates at 1024 chars). Idempotent and safe to re-run.
function Add-BunToUserPath($binDirectory) {
    try {
        $normalizedTarget = $binDirectory.TrimEnd('\')
        $userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
        $existingEntries = @()
        if ($userPath) { $existingEntries = $userPath -split ';' | Where-Object { $_ } }

        # -ieq (explicit case-insensitive) plus trailing-slash trim so we don't
        # append a near-duplicate that differs only by case or a trailing '\'.
        if ($existingEntries | Where-Object { $_.TrimEnd('\') -ieq $normalizedTarget }) {
            Write-Info "Bun's bin directory is already on your user PATH."
            return
        }

        $newPath = if ([string]::IsNullOrEmpty($userPath)) {
            $normalizedTarget
        } else {
            ($userPath.TrimEnd(';') + ';' + $normalizedTarget)
        }
        [Environment]::SetEnvironmentVariable('Path', $newPath, 'User')
        Write-Ok "Added Bun to your user PATH ($normalizedTarget)."
    } catch {
        Write-Warn "Couldn't update your PATH automatically: $($_.Exception.Message)"
        Write-Info 'Add this folder to your PATH (System Environment Variables), then open a new terminal:'
        Write-Host "    $binDirectory"
    }
}

Write-Host ''
# Make this session able to find the just-linked shim for the check below,
# without clobbering process-only PATH entries (nvm, conda, dev shells).
if (($env:Path -split ';') -notcontains $bunBinDirectory) {
    $env:Path = "$env:Path;$bunBinDirectory"
}

# Verify the shim bun just linked actually exists, rather than trusting a
# possibly-stale `teams-org` from some other install elsewhere on PATH.
$linkedShimExists = (Test-Path (Join-Path $bunBinDirectory 'teams-org')) -or
                    (Test-Path (Join-Path $bunBinDirectory 'teams-org.exe')) -or
                    (Test-Path (Join-Path $bunBinDirectory 'teams-org.cmd')) -or
                    (Test-Path (Join-Path $bunBinDirectory 'teams-org.bunx'))

if ($linkedShimExists -and (Get-Command teams-org -ErrorAction SilentlyContinue)) {
    Write-Ok "Success! 'teams-org' is ready to use."
    Write-Host ''
    Write-Info 'Try it now:'
    Write-Host '    teams-org extract'
} else {
    Write-Warn "'teams-org' was linked but isn't on your PATH yet."
    Write-Host ''
    Add-BunToUserPath $bunBinDirectory
    Write-Host ''
    Write-Info 'Open a new terminal to pick up the updated PATH.'
    Write-Host ''
    Write-Info 'Meanwhile you can always run it from this folder with:'
    Write-Host '    bun run bin/cli.ts extract'
}
