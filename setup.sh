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

echo
if command -v teams-org >/dev/null 2>&1; then
  ok "Success! 'teams-org' is ready to use."
  echo
  info "Try it now:"
  echo "    teams-org extract"
else
  warn "'teams-org' was linked but isn't on your PATH yet."
  echo
  info "Add Bun's global bin directory to your PATH:"
  if [ "$PLATFORM" = "windows" ]; then
    echo "    Add this folder to your PATH (System Environment Variables):"
    echo "        %USERPROFILE%\\.bun\\bin"
    echo "    ...then open a new terminal."
  else
    echo "    Add this line to your ~/.bashrc or ~/.zshrc, then restart your shell:"
    echo "        export PATH=\"$BUN_BIN_DIR:\$PATH\""
  fi
  echo
  info "Meanwhile you can always run it from this folder with:"
  echo "    bun run bin/cli.ts extract"
fi
