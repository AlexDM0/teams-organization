# Microsoft Teams Organization Tree Extractor

A fast, Typescript/Bun-based script that extracts the organization tree from Microsoft Teams (Entra ID) using the Microsoft Graph API and generates a beautiful, interactive D3.js visualization.

## Features
- **Interactive Browser Sign-In**: Signs in through your system browser by default (no passwords stored); a device-code flow is available via `--device-code` for headless machines.
- **Deep Graph API Integration**: Automatically resolves managers and fetches profile-picture thumbnails encoded directly to Base64, with throttling-aware retries and silent token refresh for large tenants.
- **Interactive Visualization**: Uses `d3-org-chart` to render an aesthetic, self-contained interactive web UI for exploring the organizational hierarchy. Front-end libraries are version-pinned and verified with Subresource Integrity.

## Prerequisites
- **Bun**: Install [Bun](https://bun.sh/) if you haven't already (the `./setup.sh` script checks this for you).
- **Azure AD App Registration** *(optional)*: The tool works out of the box using Microsoft's public app registration. You only need your own registration if your organization has disabled it — see [Using your own app registration](#using-your-own-app-registration-optional).

## Authentication: Client ID & Tenant ID

`teams-org` signs in with an Azure AD (Microsoft Entra) application, identified by
a **client ID** and **tenant ID**.

By default it uses Microsoft's public **"Microsoft Graph PowerShell"** application:

| Setting | Default value |
| --- | --- |
| `clientId` | `14d82eec-204b-4c2f-b7e8-296a70dab67e` |
| `tenantId` | `common` |

This app is multi-tenant and enabled in many organizations, so **no configuration
is required** — it usually works out of the box.

### How IDs are chosen (`tenant.json`)

Each extraction folder gets a `tenant.json` that controls which IDs are used:

1. **First run in a folder** — `teams-org` creates `tenant.json` seeded with the
   defaults, then asks whether you'd like to use your **own** client/tenant ID.
   - Answer **No** → it extracts immediately using the defaults.
   - Answer **Yes** → it writes a `custom_tenant_id.html` guide into the folder and
     exits so you can fill in your IDs. Edit `tenant.json`, then re-run.
2. **Re-run in the same folder** — `teams-org` reads `tenant.json`:
   - If the values **differ from the defaults**, it uses them and extracts right away.
   - If they're **still the defaults**, it asks whether to proceed with the defaults.

> Tip: pass `--yes` to skip all prompts and use whatever is in `tenant.json`
> (defaults on a fresh folder). You can also set the IDs up front with
> `--client-id` / `--tenant-id`, which are written straight into `tenant.json`.

## Using your own app registration (optional)

Follow these steps only if you want a dedicated Azure AD app (the generated
`custom_tenant_id.html` contains the same guide):

### 1. Register the Application
1. Go to the [Microsoft Entra admin center](https://entra.microsoft.com/) (or the [Azure Portal](https://portal.azure.com/)) and log in with your company account.
2. In the left-hand menu, expand **Applications** and click on **App registrations**.
3. Click **New registration** at the top.
4. Name it something recognizable, like *"Teams Org Tree Extractor"*.
5. Under **Supported account types**, leave the default ("Accounts in this organizational directory only").
6. Click **Register** at the bottom.

### 2. Copy your IDs
Immediately after registering, you will be taken to the app's **Overview** page. Look at the "Essentials" section at the top:
- Copy the **Application (client) ID** — this is your `clientId`.
- Copy the **Directory (tenant) ID** — this is your `tenantId`.

Put these into the `tenant.json` in your extraction folder:
```json
{
  "clientId": "YOUR-APPLICATION-CLIENT-ID",
  "tenantId": "YOUR-DIRECTORY-TENANT-ID"
}
```

### 3. Enable public client flows (Crucial!)
Both the interactive browser sign-in and the device-code flow are public-client
flows, so we must explicitly allow them.
1. On the left menu of your App Registration, click **Authentication**.
2. Scroll all the way down to **Advanced settings**.
3. Look for **Allow public client flows** (or "Enable the following mobile and desktop flows") and toggle it to **Yes**.
4. Click **Save** at the top.

### 4. Grant API Permissions
Lastly, we need to give the app permission to read the organization's user profiles.
1. On the left menu, click **API permissions**.
2. Click **Add a permission** -> **Microsoft Graph** -> **Delegated permissions**.
3. In the search box, type `User.Read.All` and check the box next to it.
4. Click **Add permissions** at the bottom.
5. **Important:** Back on the API permissions screen, click the button that says **Grant admin consent for [Your Organization]** (you may need a company administrator to click this for you if you don't have admin rights).

## Quick Start (CLI)

The project ships a `teams-org` CLI. It works out of the box using Microsoft's
public app registration — see [Authentication](#authentication-client-id--tenant-id)
for details and how to use your own IDs via `tenant.json`.

1. **Install Dependencies & link the CLI** (once)

   Run the setup script — it checks that Bun is installed and working, installs
   dependencies, and links `teams-org` onto your PATH:
   ```bash
   ./setup.sh            # macOS / Linux / Windows (Git Bash, MSYS2, or WSL)
   ```
   Prefer to do it manually?
   ```bash
   bun install
   bun link          # makes `teams-org` available on your PATH
   ```
   > No global link? You can always run it via `bun run extract`.

2. **Extract**
   ```bash
   teams-org extract
   ```
   The command will:
   - Check that required tools (like Bun) are installed and, if not, tell you how
     to install them.
   - Ask which folder to extract into (defaults to the current directory).
   - Create/read `tenant.json` in that folder and, on a fresh folder, ask whether
     you want to use custom client/tenant IDs (see
     [Authentication](#authentication-client-id--tenant-id)).
   - Authenticate through your browser (interactive sign-in by default; pass
     `--device-code` for the device-code flow) and download the whole org.
   - Write these files into the chosen folder:
     - `tenant.json` — the client/tenant IDs used for this folder.
     - `org_tree.json` — the readable hierarchy (no image blobs).
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
pulling a new template)? Use `update` — it reuses the existing `org_tree.json`
and `photos.json` in the folder and skips authentication and downloading:

```bash
teams-org update                 # rebuild index.html in the current folder
teams-org update --folder ./my-org
teams-org update --no-inline     # keep CDN <script> links instead of inlining
```

### Viewing the raw JSON (dev mode)

The generated `index.html` is portable and needs no server. If you instead want to
open the template in `templates/index.html` against the raw `org_tree.json` /
`photos.json` files, copy those files next to the template (or serve them from the
same folder) and serve locally to satisfy browser CORS rules:

```bash
bun run serve
```
Open the printed URL in your browser (navigate into `templates/`) to explore the
interactive chart!

