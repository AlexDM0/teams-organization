#!/usr/bin/env bash
#
# setup.sh — install the `teams-org` CLI so it's available system-wide.
#
# Works on macOS, Linux, and Windows (via Git Bash / MSYS2). On WSL it takes the
# normal Linux path (Windows users without a bash shell should use setup.ps1).
# It verifies Bun is installed and working, installs dependencies, and links the
# CLI into Bun's global bin directory so `teams-org` is on your PATH.
#

# This script relies on bashisms (BASH_SOURCE, [[ ]], `local`, pipefail). Refuse
# to run under a POSIX /bin/sh (dash), where `set -o pipefail` and ${BASH_SOURCE}
# would otherwise abort with an opaque error before any friendly output.
if [ -z "${BASH_VERSION:-}" ]; then
  echo "Please run this script with bash:  bash setup.sh" >&2
  exit 1
fi

set -euo pipefail

# Resolve the directory this script lives in (so it works from anywhere).
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]:-$0}")" >/dev/null 2>&1 && pwd)"
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

# On Windows (Git Bash / MSYS2 / Cygwin) hand the whole install off to the native
# PowerShell installer — it is the single source of truth for the Windows
# user-PATH registry write, so that logic isn't duplicated here in a second
# language. (WSL reports "Linux" from uname and takes the normal path below.)
if [ "$PLATFORM" = "windows" ] && command -v powershell.exe >/dev/null 2>&1; then
  info "Windows detected — handing off to the native setup.ps1..."
  PS1_PATH="$SCRIPT_DIR/setup.ps1"
  if command -v cygpath >/dev/null 2>&1; then PS1_PATH="$(cygpath -w "$PS1_PATH")"; fi
  exec powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$PS1_PATH"
fi

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

# Identify the user's login-shell family so we edit the right startup file — and
# never write bash syntax into a fish/csh config that can't read it.
detect_shell_family() {
  case "${SHELL:-}" in
    */zsh)        echo "zsh" ;;
    */bash)       echo "bash" ;;
    */fish)       echo "fish" ;;
    */tcsh|*/csh) echo "csh" ;;
    *)            echo "other" ;;
  esac
}

# Pick the startup file for a zsh/bash shell that we can safely append an
# `export PATH=...` line to. Only called for the zsh/bash families.
posix_rc_file() {
  case "$(detect_shell_family)" in
    zsh) echo "$HOME/.zshrc" ;;
    bash)
      if [ "$PLATFORM" = "macos" ]; then
        # macOS starts Terminal/iTerm bash as a *login* shell, which reads the
        # first of ~/.bash_profile, ~/.bash_login, ~/.profile that exists. Prefer
        # whichever already exists so we don't create a fresh ~/.bash_profile
        # that would shadow (and silently disable) an existing ~/.profile.
        for candidate in "$HOME/.bash_profile" "$HOME/.bash_login" "$HOME/.profile"; do
          if [ -f "$candidate" ]; then echo "$candidate"; return; fi
        done
        echo "$HOME/.bash_profile"
      else
        # Most Linux terminals start a non-login interactive bash → ~/.bashrc.
        echo "$HOME/.bashrc"
      fi ;;
  esac
}

# Compose the `export PATH` line. Use the literal $HOME token when the bin dir is
# under $HOME so the line survives a changed $HOME / synced dotfiles instead of
# baking in an absolute per-user path, and matches on re-runs.
bun_path_export_line() {
  local path_value="$BUN_BIN_DIR"
  case "$BUN_BIN_DIR" in
    "$HOME/"*) path_value="\$HOME/${BUN_BIN_DIR#"$HOME"/}" ;;
  esac
  printf 'export PATH="%s:$PATH"' "$path_value"
}

# Print copy-paste PATH guidance for shells we won't edit automatically (fish,
# csh/tcsh, or an unrecognized login shell) so we never write a startup file the
# shell can't read and then wrongly claim success.
print_manual_path_help() {
  warn "Your login shell ($(detect_shell_family)) isn't auto-configured — add Bun's bin dir yourself:"
  case "$(detect_shell_family)" in
    fish) echo "    fish_add_path \"$BUN_BIN_DIR\"" ;;
    csh)  echo "    setenv PATH \"$BUN_BIN_DIR:\$PATH\"     # add to ~/.tcshrc or ~/.cshrc" ;;
    *)    echo "    export PATH=\"$BUN_BIN_DIR:\$PATH\"     # add to your shell's startup file" ;;
  esac
  echo
  info "Then open a new terminal to pick it up."
}

# Append the Bun PATH export to the given rc file (creating it if absent), unless
# an entry for Bun's bin directory is already present.
add_bun_to_path() {
  local rc_file="$1"
  local export_line
  export_line="$(bun_path_export_line)"

  # Whole-line match (grep -x) so a commented-out or otherwise-embedded copy of
  # this string never counts as "already present" — a plain -F is a *substring*
  # match and would false-positive on `# export PATH="…/.bun/bin:$PATH"`.
  if [ -f "$rc_file" ] && grep -qxF "$export_line" "$rc_file" 2>/dev/null; then
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

# Treat the install as successful only when the shim bun just linked actually
# exists in Bun's bin dir — not merely when *some* `teams-org` resolves on PATH
# (which could be a stale npm-global install shadowing it) and not when it only
# resolves via an ephemeral PATH entry that a fresh terminal won't have.
linked_shim_present() {
  [ -e "$BUN_BIN_DIR/teams-org" ]
}

echo
if linked_shim_present && command -v teams-org >/dev/null 2>&1; then
  ok "Success! 'teams-org' is ready to use."
  echo
  info "Try it now:"
  echo "    teams-org extract"
else
  if ! linked_shim_present; then
    warn "'teams-org' was linked but no shim was found in $BUN_BIN_DIR."
    info "Check where bun linked it with:  bun pm bin -g   (then add that dir to PATH)."
    echo
  else
    warn "'teams-org' was linked but isn't on your PATH yet."
    echo
  fi
  if [ "$PLATFORM" = "windows" ]; then
    # Only reached when powershell.exe was unavailable for the handoff above, so
    # we can't automate the registry PATH write — give manual instructions.
    warn "Couldn't run PowerShell to update your PATH automatically."
    info "Add this folder to your PATH (System Environment Variables), then open a new terminal:"
    echo "        $BUN_BIN_DIR"
  else
    case "$(detect_shell_family)" in
      zsh|bash)
        RC_FILE="$(posix_rc_file)"
        add_bun_to_path "$RC_FILE"
        echo
        info "Restart your shell (or run 'source $RC_FILE') to pick it up." ;;
      *)
        print_manual_path_help ;;
    esac
  fi
  echo
  info "Meanwhile you can always run it from this folder with:"
  echo "    bun run bin/cli.ts extract"
fi
