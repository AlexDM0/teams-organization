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

# --- Pretty output helpers -------------------------------------------------
# Glyphs are built from code points at runtime so the source stays pure ASCII
# (a BOM-less .ps1 with literal non-ASCII would render as mojibake under 5.1).
$script:BulletGlyph = [char]0x2022  # bullet
$script:CheckGlyph  = [char]0x2713  # check mark
$script:CrossGlyph  = [char]0x2717  # ballot X

# Exit code the launcher at the bottom hands back (only when run as a script).
$script:SetupExitCode = 0

function Write-Bold($message) { Write-Host $message -ForegroundColor White }
function Write-Info($message) { Write-Host "$script:BulletGlyph $message" -ForegroundColor Cyan }
function Write-Ok($message)   { Write-Host "$script:CheckGlyph $message" -ForegroundColor Green }
function Write-Warn($message) { Write-Host "! $message" -ForegroundColor Yellow }
function Write-Err($message)  { Write-Host "$script:CrossGlyph $message" -ForegroundColor Red }

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

# Persistently add $binDirectory to the *user* PATH if it isn't already there.
# Reads/writes the raw registry value with DoNotExpandEnvironmentNames so
# existing %VAR%-style entries (e.g. the %USERPROFILE%\.bun\bin that bun's own
# installer may write) are preserved rather than frozen to their current
# expansion. Idempotent and safe to re-run. We deliberately do not broadcast the
# change to running processes - the CLI becomes available on the next terminal
# open, which keeps this pure .NET and portable (no user32 P/Invoke).
function Add-BunToUserPath($binDirectory) {
    try {
        $normalizedTarget = $binDirectory.TrimEnd('\')

        $environmentKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Environment', $true)
        if (-not $environmentKey) {
            $environmentKey = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey('Environment')
        }
        try {
            $userPath = [string]$environmentKey.GetValue(
                'Path', '', [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            $existingEntries = @()
            if ($userPath) { $existingEntries = $userPath -split ';' | Where-Object { $_ } }

            # -ieq (explicit case-insensitive) plus trailing-slash trim so we
            # don't append a near-duplicate that differs only by case or a '\'.
            if ($existingEntries | Where-Object { $_.TrimEnd('\') -ieq $normalizedTarget }) {
                Write-Info "Bun's bin directory is already on your user PATH."
                return
            }

            $newPath = if ([string]::IsNullOrWhiteSpace($userPath)) {
                $normalizedTarget
            } else {
                ($userPath.TrimEnd(';') + ';' + $normalizedTarget)
            }
            # ExpandString => REG_EXPAND_SZ, so %VAR% entries keep expanding.
            $environmentKey.SetValue('Path', $newPath, [Microsoft.Win32.RegistryValueKind]::ExpandString)
        } finally {
            $environmentKey.Close()
        }

        Write-Ok "Added Bun to your user PATH ($normalizedTarget)."
    } catch {
        Write-Warn "Couldn't update your PATH automatically: $($_.Exception.Message)"
        Write-Info 'Add this folder to your PATH (System Environment Variables), then open a new terminal:'
        Write-Host "    $binDirectory"
    }
}

# Everything runs inside this function so early-outs use `return` (not `exit`):
# dot-sourcing this file therefore can't close the caller's shell. The launcher
# at the bottom maps the result to a real exit code only when run as a script.
function Invoke-TeamsOrgSetup {
    # Run from the folder this script lives in. When $PSScriptRoot is empty (the
    # script was dot-sourced or pasted via iex), fall back to the current dir but
    # only if it really is this project, so `bun link` can't register the wrong
    # package or scaffold stray files in an unrelated folder.
    if ($PSScriptRoot) {
        Set-Location -Path $PSScriptRoot
    } else {
        $packageName = $null
        if (Test-Path 'package.json') {
            try { $packageName = (Get-Content 'package.json' -Raw | ConvertFrom-Json).name } catch { }
        }
        if ($packageName -ne 'msteams-org-tree') {
            Write-Err "Run this from the teams-org project folder (current directory isn't it)."
            $script:SetupExitCode = 1; return
        }
    }

    Write-Bold 'Setting up the teams-org CLI (windows)'
    Write-Host ''

    # --- 1. Verify Bun is installed AND working ----------------------------
    Write-Info 'Checking for Bun...'
    if (-not (Get-Command bun -ErrorAction SilentlyContinue)) {
        Show-BunInstallHelp
        $script:SetupExitCode = 1; return
    }

    # Confirm it actually runs. Reset $LASTEXITCODE first so a stale 0 from an
    # earlier call can't be mistaken for success if `bun` fails to launch.
    $bunVersion = $null
    try {
        $global:LASTEXITCODE = 0
        $bunOutput = (& bun --version 2>&1 | Out-String)
        if ($LASTEXITCODE -eq 0 -and $bunOutput) {
            $bunVersion = ($bunOutput -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -First 1).Trim()
        }
    } catch { }

    if (-not $bunVersion) {
        Write-Err "Found 'bun' on PATH but it failed to run - the install looks broken."
        Show-BunInstallHelp
        $script:SetupExitCode = 1; return
    }
    Write-Ok "Bun $bunVersion is installed and working."

    # --- 2. Install dependencies -------------------------------------------
    Write-Info 'Installing dependencies (bun install)...'
    # Pre-set a non-zero code so a native call that fails to even launch (leaving
    # $LASTEXITCODE untouched) is treated as a failure, not a stale success.
    $global:LASTEXITCODE = 1
    & bun install
    if ($LASTEXITCODE -ne 0) { Write-Err 'bun install failed.'; $script:SetupExitCode = 1; return }
    Write-Ok 'Dependencies installed.'

    # --- 3. Link the CLI globally ------------------------------------------
    Write-Info 'Linking the teams-org CLI (bun link)...'
    $global:LASTEXITCODE = 1
    & bun link
    if ($LASTEXITCODE -ne 0) { Write-Err 'bun link failed.'; $script:SetupExitCode = 1; return }
    Write-Ok 'CLI linked.'

    # --- 4. Ensure Bun's bin directory is on the user PATH -----------------
    # Bun places global bins in $env:BUN_INSTALL\bin (default %USERPROFILE%\.bun\bin).
    $bunInstallRoot = if ($env:BUN_INSTALL) { $env:BUN_INSTALL } else { Join-Path $env:USERPROFILE '.bun' }
    $bunBinDirectory = Join-Path $bunInstallRoot 'bin'

    Write-Host ''
    # Persist to the user PATH *unconditionally* (idempotent) so the next
    # terminal always finds it. We intentionally leave the current session's PATH
    # untouched - the CLI becomes available on the next terminal open, which is
    # the only guarantee a subprocess installer can make portably anyway.
    Add-BunToUserPath $bunBinDirectory

    # Success is gated on the freshly-linked shim actually existing in the bun
    # bin dir - a filesystem check, independent of PATH (so a stale teams-org
    # elsewhere on PATH can't fake success).
    $linkedShimExists = (Test-Path (Join-Path $bunBinDirectory 'teams-org.exe')) -or
                        (Test-Path (Join-Path $bunBinDirectory 'teams-org.cmd')) -or
                        (Test-Path (Join-Path $bunBinDirectory 'teams-org.bunx')) -or
                        (Test-Path (Join-Path $bunBinDirectory 'teams-org'))

    Write-Host ''
    if ($linkedShimExists) {
        Write-Ok "Success! 'teams-org' is installed."
        Write-Host ''
        Write-Info 'Open a new terminal (to pick up PATH), then run:'
        Write-Host '    teams-org extract'
    } else {
        Write-Warn "'teams-org' was linked but no shim was found in $bunBinDirectory."
        Write-Host ''
        Write-Info 'This usually means bun linked into a different directory. Check with:'
        Write-Host '    bun pm bin -g'
        Write-Host ''
        Write-Info 'Ensure that directory is on your PATH, or re-run:  bun link'
        Write-Host ''
        Write-Info 'Meanwhile you can always run it from this folder with:'
        Write-Host '    bun run bin/cli.ts extract'
    }
}

# Run it. Because bun install/link stream their output uncaptured here, the
# function returns nothing - the exit code travels via $script:SetupExitCode.
Invoke-TeamsOrgSetup

# Only hard-exit when launched as a script (-File / & .\setup.ps1). If someone
# dot-sources this file, `exit` would close their whole shell, so we skip it.
if ($MyInvocation.InvocationName -ne '.') { exit $script:SetupExitCode }
