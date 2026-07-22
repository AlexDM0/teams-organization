# Using a custom Client ID & Tenant ID

> **You almost certainly don't need this.** `teams-org` works out of the box with
> Microsoft's public app registration — see the note below. Follow this guide
> **only** if your IT department has disabled that public app, or you specifically
> want a dedicated Azure AD (Microsoft Entra) application.

`teams-org` signs in with an Azure AD application identified by a **client ID** and
a **tenant ID**. By default it uses Microsoft's public **"Microsoft Graph
PowerShell"** application:

| Setting | Default value |
| --- | --- |
| `clientId` | `14d82eec-204b-4c2f-b7e8-296a70dab67e` |
| `tenantId` | `common` |

This app is multi-tenant and enabled in many organizations, so **no configuration
is required** — it usually works out of the box.

## How IDs are chosen (`tenant.json`)

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

## Registering your own app

The generated `custom_tenant_id.html` contains this same guide.

### 1. Register the application
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

### 3. Enable public client flows (crucial!)
Both the interactive browser sign-in and the device-code flow are public-client
flows, so we must explicitly allow them.
1. On the left menu of your App Registration, click **Authentication**.
2. Scroll all the way down to **Advanced settings**.
3. Look for **Allow public client flows** (or "Enable the following mobile and desktop flows") and toggle it to **Yes**.
4. Click **Save** at the top.

### 4. Grant API permissions
Lastly, we need to give the app permission to read the organization's user profiles.
1. On the left menu, click **API permissions**.
2. Click **Add a permission** → **Microsoft Graph** → **Delegated permissions**.
3. In the search box, type `User.Read.All` and check the box next to it.
4. Click **Add permissions** at the bottom.
5. **Important:** Back on the API permissions screen, click the button that says **Grant admin consent for [Your Organization]** (you may need a company administrator to click this for you if you don't have admin rights).

### 5. Run the extraction
From your extraction folder, run:
```bash
teams-org extract
```
Because `tenant.json` now differs from the defaults, `teams-org` uses your custom
IDs automatically.
