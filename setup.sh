#!/usr/bin/env bash
#
# setup.sh — install the `teams-org` CLI so it's available system-wide.
#
# Works on macOS, Linux, and Windows (via Git Bash / MSYS2 / WSL).
# It verifies Bun is installed and working, installs dependencies, and links
# the CLI into Bun's global bin directory so `teams-org` is on your PATH.
#
set -euo pipefail

# Resolve the directory this script lives in (so it works from anywhere).
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" >/dev/null 2>&1 && pwd)"
cd "$SCRIPT_DIR"

# --- Pretty output helpers -------------------------------------------------
bold()  { printf '\033[1m%s\033[0m\n' "$1"; }
info()  { printf '\033[36m•\033[0m %s\n' "$1"; }
ok()    { printf '\033[32m✓\033[0m %s\n' "$1"; }
warn()  { printf '\033[33m!\033[0m %s\n' "$1"; }
err()   { printf '\033[31m✗\033[0m %s\n' "$1" >&2; }

# --- Detect OS -------------------------------------------------------------
OS="$(uname -s 2>/dev/null || echo unknown)"
case "$OS" in
  Darwin*)                       PLATFORM="macos" ;;
  Linux*)                        PLATFORM="linux" ;;
  MINGW*|MSYS*|CYGWIN*|Windows*) PLATFORM="windows" ;;
  *)                             PLATFORM="unknown" ;;
esac

bold "Setting up the teams-org CLI ($PLATFORM)"
echo

# --- Bun install instructions ---------------------------------------------
print_bun_install_help() {
  echo
  err "Bun is required but was not found (or is not working)."
  echo
  if [ "$PLATFORM" = "windows" ]; then
    info "Install Bun on Windows (run in PowerShell):"
    echo "    powershell -c \"irm bun.sh/install.ps1 | iex\""
  else
    info "Install Bun on macOS/Linux:"
    echo "    curl -fsSL https://bun.sh/install | bash"
  fi
  echo
  info "More info: https://bun.sh/"
  echo
  warn "After installing, open a new terminal (so PATH updates) and re-run ./setup.sh"
}

# --- 1. Verify Bun is installed AND working --------------------------------
info "Checking for Bun..."
if ! command -v bun >/dev/null 2>&1; then
  print_bun_install_help
  exit 1
fi

# Confirm it actually runs (a broken/partial install fails here).
if ! BUN_VERSION="$(bun --version 2>/dev/null)"; then
  err "Found 'bun' on PATH but it failed to run — the install looks broken."
  print_bun_install_help
  exit 1
fi
ok "Bun $BUN_VERSION is installed and working."

# --- 2. Install dependencies ----------------------------------------------
info "Installing dependencies (bun install)..."
bun install
ok "Dependencies installed."

# --- 3. Link the CLI globally ---------------------------------------------
info "Linking the teams-org CLI (bun link)..."
bun link
ok "CLI linked."

# --- 4. Verify it's available on PATH -------------------------------------
# Bun places global bins in $BUN_INSTALL/bin (defaults to ~/.bun/bin).
BUN_BIN_DIR="${BUN_INSTALL:-$HOME/.bun}/bin"

# Pick the shell rc file to edit, based on the user's login shell. On macOS the
# default is zsh (~/.zshrc); most Linux distros default to bash (~/.bashrc).
# Bun's own installer only edits an rc file that already exists, so on a fresh
# machine with no ~/.zshrc the PATH line is never added — we create it here.
shell_rc_file() {
  case "${SHELL:-}" in
    */zsh)  echo "$HOME/.zshrc" ;;
    */bash)
      # macOS starts Terminal/iTerm bash as a *login* shell, which reads
      # ~/.bash_profile (not ~/.bashrc). Most Linux terminals start a non-login
      # interactive bash, which reads ~/.bashrc.
      if [ "$PLATFORM" = "macos" ]; then echo "$HOME/.bash_profile"; else echo "$HOME/.bashrc"; fi ;;
    *)      [ "$PLATFORM" = "macos" ] && echo "$HOME/.zshrc" || echo "$HOME/.bashrc" ;;
  esac
}

# Persistently add Bun's bin directory to the Windows *user* PATH (registry),
# using PowerShell so the full PATH is preserved (setx truncates at 1024 chars
# and is unsafe). Idempotent: does nothing if the entry is already present.
add_bun_to_windows_path() {
  if ! command -v powershell.exe >/dev/null 2>&1; then
    warn "Couldn't find powershell.exe to update your PATH automatically."
    info "Add this folder to your PATH (System Environment Variables), then open a new terminal:"
    echo "        %USERPROFILE%\\.bun\\bin"
    return 0
  fi

  # `|| result=""` keeps `set -euo pipefail` from aborting the whole script (and
  # skipping the manual-fallback message below) if PowerShell fails or is blocked.
  local result
  result="$(powershell.exe -NoProfile -Command '
    # Honor a Windows-style $BUN_INSTALL (mirrors setup.ps1); else the default.
    $bunBin = if ($env:BUN_INSTALL -and $env:BUN_INSTALL -match "^[A-Za-z]:\\") {
      Join-Path $env:BUN_INSTALL "bin"
    } else {
      Join-Path $env:USERPROFILE ".bun\bin"
    }
    $target = $bunBin.TrimEnd("\")
    $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
    $entries = @()
    if ($userPath) { $entries = $userPath -split ";" | Where-Object { $_ } }
    # Case-insensitive, trailing-backslash-insensitive match avoids duplicates.
    if ($entries | Where-Object { $_.TrimEnd("\") -ieq $target }) {
      Write-Output "present"
    } else {
      $newPath = if ([string]::IsNullOrEmpty($userPath)) { $target } else { ($userPath.TrimEnd(";") + ";" + $target) }
      [Environment]::SetEnvironmentVariable("Path", $newPath, "User")
      Write-Output "added"
    }
  ' 2>/dev/null | tr -d "[:space:]")" || result=""

  case "$result" in
    added)   ok "Added Bun to your Windows user PATH (%USERPROFILE%\\.bun\\bin)." ;;
    present) info "Bun's bin directory is already on your Windows user PATH." ;;
    *)
      warn "Couldn't update your PATH automatically."
      info "Add this folder to your PATH (System Environment Variables), then open a new terminal:"
      echo "        %USERPROFILE%\\.bun\\bin"
      ;;
  esac
}

# Append the Bun PATH export to the given rc file (creating it if absent),
# unless an entry for Bun's bin directory is already present.
add_bun_to_path() {
  local rc_file="$1"
  local export_line="export PATH=\"$BUN_BIN_DIR:\$PATH\""

  # Match the exact line we would write, so re-runs are perfectly idempotent and
  # we never false-positive on a comment or a longer path that merely contains
  # this one as a substring. (A differently-formatted bun line just means a
  # harmless duplicate export — last one wins.)
  if [ -f "$rc_file" ] && grep -qF "$export_line" "$rc_file"; then
    info "Bun's bin directory is already referenced in $rc_file."
    return 0
  fi

  if [ ! -f "$rc_file" ]; then
    info "No $rc_file found — creating one."
  fi

  {
    printf '\n# Added by teams-org setup.sh — Bun global bin directory\n'
    printf '%s\n' "$export_line"
  } >> "$rc_file"
  ok "Added Bun to your PATH in $rc_file."
}

echo
if command -v teams-org >/dev/null 2>&1; then
  ok "Success! 'teams-org' is ready to use."
  echo
  info "Try it now:"
  echo "    teams-org extract"
else
  warn "'teams-org' was linked but isn't on your PATH yet."
  echo
  if [ "$PLATFORM" = "windows" ]; then
    add_bun_to_windows_path
    echo
    info "Open a new terminal to pick up the updated PATH."
  else
    RC_FILE="$(shell_rc_file)"
    add_bun_to_path "$RC_FILE"
    echo
    info "Restart your shell (or run 'source $RC_FILE') to pick it up."
  fi
  echo
  info "Meanwhile you can always run it from this folder with:"
  echo "    bun run bin/cli.ts extract"
fi
