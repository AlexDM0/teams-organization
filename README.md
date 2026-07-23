# Microsoft Teams Organization Tree Extractor

A fast, Typescript/Bun-based script that extracts the organization tree from Microsoft Teams (Entra ID) using the Microsoft Graph API and generates a beautiful, interactive D3.js visualization.

## Quick Start (CLI)

The project ships a `teams-org` CLI. It works out of the box using Microsoft's
public app registration — no configuration needed. (Only if your IT department
has disabled that public app do you need your own IDs — see
[CUSTOM_TENANT.md](CUSTOM_TENANT.md).)

1. **Install Dependencies & link the CLI** (once)

   Run the setup script — it checks that Bun is installed and working, installs
   dependencies, and links `teams-org` onto your PATH:
   ```bash
   ./setup.sh            # macOS / Linux / Windows (Git Bash, MSYS2, or WSL)
   ```
   The script also makes sure `teams-org` ends up on your PATH: on macOS/Linux it
   adds Bun's bin directory to your shell startup file (e.g. `~/.zshrc`,
   **creating it if it doesn't exist**), and for shells it can't edit
   automatically (fish, csh) it prints the exact line to add; on Windows it adds
   it to your user PATH via PowerShell. Open a new terminal afterwards so the
   change takes effect.

   Prefer to do it manually?
   ```bash
   bun install
   bun link          # makes `teams-org` available on your PATH
   ```
   > No global link? You can always run it via `bun run extract`.

   #### Windows

   Two options — pick whichever shell you use:

   **Native PowerShell** (no bash needed) — run `setup.ps1`:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\setup.ps1
   ```

   **Git Bash / MSYS2** — run `setup.sh` (it won't run directly in PowerShell or
   CMD; needs a bash shell such as the one bundled with
   [Git for Windows](https://git-scm.com/download/win)):
   ```bash
   ./setup.sh
   ```

   Both `setup.ps1` and `setup.sh` (under Git Bash/MSYS2) install dependencies,
   link the CLI, and add `%USERPROFILE%\.bun\bin` to your **Windows user PATH**
   automatically (preserving your existing PATH; safe to re-run). **Open a new
   terminal** afterwards so the updated PATH is picked up, then `teams-org
   extract` works from any shell — PowerShell, CMD, or Git Bash.

   > Under **WSL** you're in a Linux environment: run `./setup.sh` and install a
   > Linux Bun — it's treated exactly like the macOS/Linux flow above (it edits
   > `~/.bashrc`), independent of your Windows Bun/PATH.

   Prefer to do it manually? From PowerShell:
   ```powershell
   bun install
   bun link
   ```
   Bun's installer normally puts `%USERPROFILE%\.bun\bin` on your PATH already; if
   `teams-org` still isn't found, add that folder under **System Environment
   Variables → Path** and open a new terminal. (Meanwhile `bun run extract` always
   works from the project folder.)

2. **Extract**
   ```bash
   teams-org extract
   ```
   The command will:
   - Check that required tools (like Bun) are installed and, if not, tell you how
     to install them.
   - Ask which folder to extract into (defaults to the current directory).
   - Create/read `tenant.json` in that folder and, on a fresh folder, ask whether
     you want to use custom client/tenant IDs (almost always: no — see
     [CUSTOM_TENANT.md](CUSTOM_TENANT.md)).
   - Authenticate through your browser (interactive sign-in by default; pass
     `--device-code` for the device-code flow) and download the whole org.
   - Write these files into the chosen folder:
     - `tenant.json` — the client/tenant IDs used for this folder.
     - `org_graph.json` — the readable org graph: `{ nodes, edges }`, where nodes
       are people and edges are typed relationships (`IS_MANAGER_OF`, manager →
       report). No image blobs.
     - `photos.json` — a separate `id → base64` map of profile pictures.
     - `index.html` — a **fully portable**, self-contained visualization with the
       data (and vendor libraries) embedded. Just double-click it — no web server
       needed.
     - `custom_tenant_id.html` — *(only if you opt into custom IDs)* a guide for
       setting them up.

### Useful options

```bash
teams-org extract --folder ./my-org      # choose output folder non-interactively
teams-org extract --yes                  # accept all defaults, no prompts
teams-org extract --no-photos            # skip profile pictures
teams-org extract --no-inline            # keep CDN <script> links (smaller file, needs internet)
teams-org extract --device-code          # use the device-code flow instead of the browser
teams-org extract --tenant-id <id> --client-id <id>   # use your own app registration
```

### Rebuilding without re-downloading

Already have an extraction and just want a fresh `index.html` (for example after
pulling a new template)? Use `update` — it reuses the existing `org_graph.json`
and `photos.json` in the folder and skips authentication and downloading:

```bash
teams-org update                 # rebuild index.html in the current folder
teams-org update --folder ./my-org
teams-org update --no-inline     # keep CDN <script> links instead of inlining
```

### Viewing the raw JSON (dev mode)

The generated `index.html` is portable and needs no server. If you instead want to
open the template in `templates/index.html` against the raw `org_graph.json` /
`photos.json` files, copy those files next to the template (or serve them from the
same folder) and serve locally to satisfy browser CORS rules:

```bash
bun run serve
```
Open the printed URL in your browser (navigate into `templates/`) to explore the
interactive chart!

## Features
- **Interactive Browser Sign-In**: Signs in through your system browser by default (no passwords stored); a device-code flow is available via `--device-code` for headless machines.
- **Deep Graph API Integration**: Automatically resolves managers and fetches profile-picture thumbnails encoded directly to Base64, with throttling-aware retries and silent token refresh for large tenants.
- **Interactive Visualization**: Uses `d3-org-chart` to render an aesthetic, self-contained interactive web UI for exploring the organizational hierarchy. Front-end libraries are version-pinned and verified with Subresource Integrity.

## Prerequisites
- **Bun**: Install [Bun](https://bun.sh/) if you haven't already (the `./setup.sh` script checks this for you).
- **Azure AD App Registration** *(rarely needed)*: The tool works out of the box using Microsoft's public app registration. You only need your own registration if your organization has disabled it — see [CUSTOM_TENANT.md](CUSTOM_TENANT.md).

## Authentication

By default `teams-org` signs in with Microsoft's public **"Microsoft Graph
PowerShell"** app (`clientId 14d82eec-…`, `tenantId common`) — multi-tenant and
enabled in most organizations, so **no configuration is required**.

Each extraction folder gets a `tenant.json` recording the IDs it used. If you need
a dedicated Azure AD app (only when the public one is disabled), see
**[CUSTOM_TENANT.md](CUSTOM_TENANT.md)** for the `tenant.json` flow and a full
step-by-step registration guide.

