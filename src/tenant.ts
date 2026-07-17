import * as path from 'path';
import { confirm } from '@inquirer/prompts';
import { DEFAULT_CLIENT_ID, DEFAULT_TENANT_ID } from './extractor';

export interface TenantConfig {
    clientId: string;
    tenantId: string;
}

export const TENANT_FILE = 'tenant.json';
export const GUIDE_FILE = 'custom_tenant_id.html';

function isDefaults(tenantConfig: TenantConfig): boolean {
    return tenantConfig.clientId === DEFAULT_CLIENT_ID && tenantConfig.tenantId === DEFAULT_TENANT_ID;
}

async function writeTenantJson(file: string, tenantConfig: TenantConfig): Promise<void> {
    const fileContents = {
        clientId: tenantConfig.clientId,
        tenantId: tenantConfig.tenantId,
        _instructions:
            `To use your own Azure AD app registration, replace "clientId" and "tenantId" ` +
            `above with your values, then re-run 'teams-org extract' in this folder. ` +
            `See ${GUIDE_FILE} for a full step-by-step guide.`,
    };
    await Bun.write(file, JSON.stringify(fileContents, null, 2) + '\n');
}

async function readTenantJson(file: string): Promise<TenantConfig> {
    const parsedJson = JSON.parse(await Bun.file(file).text());
    const clientId =
        typeof parsedJson.clientId === 'string' && parsedJson.clientId.trim()
            ? parsedJson.clientId.trim()
            : DEFAULT_CLIENT_ID;
    const tenantId =
        typeof parsedJson.tenantId === 'string' && parsedJson.tenantId.trim()
            ? parsedJson.tenantId.trim()
            : DEFAULT_TENANT_ID;
    return { clientId, tenantId };
}

/**
 * Resolves which client/tenant IDs to use for an extraction based on a
 * `tenant.json` file in the output folder.
 *
 * Returns the config to extract with, or `null` when the user chose to customize
 * their IDs first (the CLI should then exit without extracting).
 */
export async function resolveTenantConfig(
    folder: string,
    options: { yes?: boolean; cliClientId: string; cliTenantId: string }
): Promise<TenantConfig | null> {
    const tenantPath = path.join(folder, TENANT_FILE);
    const guidePath = path.join(folder, GUIDE_FILE);
    const hasExplicitCustomIds =
        options.cliClientId !== DEFAULT_CLIENT_ID || options.cliTenantId !== DEFAULT_TENANT_ID;
    const tenantFileExists = await Bun.file(tenantPath).exists();

    // An explicit --client-id/--tenant-id (or CLIENT_ID/TENANT_ID env) always wins,
    // even on a re-run: persist it and use it so the flag is never silently ignored.
    if (hasExplicitCustomIds) {
        const explicitConfig: TenantConfig = { clientId: options.cliClientId, tenantId: options.cliTenantId };
        await writeTenantJson(tenantPath, explicitConfig);
        console.log(`🔑 Using IDs from the command line (tenant: ${explicitConfig.tenantId}); saved to ${tenantPath}.\n`);
        return explicitConfig;
    }

    // ---- Re-run: a tenant.json already exists ----
    if (tenantFileExists) {
        let tenantConfig: TenantConfig;
        try {
            tenantConfig = await readTenantJson(tenantPath);
        } catch {
            console.error(`\n❌ Could not parse ${tenantPath}. Please fix the JSON (or delete the file) and try again.`);
            return null;
        }

        if (!isDefaults(tenantConfig)) {
            console.log(`🔑 Using custom IDs from ${TENANT_FILE} (tenant: ${tenantConfig.tenantId}).\n`);
            return tenantConfig;
        }

        // Values are still the defaults.
        if (options.yes) return tenantConfig;
        const useDefaults = await confirm({
            message: `${TENANT_FILE} still contains the default Microsoft app registration. Use the defaults?`,
            default: true,
        });
        if (useDefaults) return tenantConfig;

        // User wants their own IDs — make sure the guide is present and stop.
        await writeGuide(guidePath);
        console.log(`\n📝 Customize your IDs in:  ${tenantPath}`);
        console.log(`📖 Step-by-step guide:     ${guidePath}`);
        console.log(`\nThen re-run 'teams-org extract' in this folder.`);
        return null;
    }

    // ---- First run: no tenant.json yet ----
    // Seed tenant.json with the defaults so it can be edited later.
    const defaultConfig: TenantConfig = { clientId: DEFAULT_CLIENT_ID, tenantId: DEFAULT_TENANT_ID };
    await writeTenantJson(tenantPath, defaultConfig);
    console.log(`🗂️  Created ${tenantPath} with the default Microsoft app registration.`);

    if (options.yes) return defaultConfig;

    const wantCustom = await confirm({
        message: 'Do you want to use your own client ID and tenant ID (a custom Azure AD app registration)?',
        default: false,
    });
    if (!wantCustom) return defaultConfig;

    // User wants custom IDs — write the guide and stop so they can edit tenant.json.
    await writeGuide(guidePath);
    console.log(`\n📝 Customize your IDs in:  ${tenantPath}`);
    console.log(`📖 Step-by-step guide:     ${guidePath}`);
    console.log(`\nEdit those values, then re-run 'teams-org extract' in this folder.`);
    return null;
}

/** Writes a self-contained HTML guide explaining how to obtain custom IDs. */
async function writeGuide(file: string): Promise<void> {
    await Bun.write(file, GUIDE_HTML);
}

const GUIDE_HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Using a custom Client ID & Tenant ID — teams-org</title>
<style>
  :root { color-scheme: light dark; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    line-height: 1.65; max-width: 820px; margin: 0 auto; padding: 40px 24px;
    color: #1f2937; background: #f8fafc;
  }
  h1 { font-size: 28px; margin-bottom: 4px; }
  h2 { margin-top: 36px; border-bottom: 2px solid #e2e8f0; padding-bottom: 6px; }
  code, pre { font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; }
  code { background: #eef2ff; color: #4338ca; padding: 2px 6px; border-radius: 5px; font-size: 90%; }
  pre { background: #0f172a; color: #e2e8f0; padding: 16px; border-radius: 10px; overflow-x: auto; }
  pre code { background: none; color: inherit; padding: 0; }
  .note { background: #ecfeff; border-left: 4px solid #06b6d4; padding: 12px 16px; border-radius: 6px; margin: 16px 0; }
  .warn { background: #fef3c7; border-left: 4px solid #f59e0b; padding: 12px 16px; border-radius: 6px; margin: 16px 0; }
  ol li { margin: 6px 0; }
  a { color: #4338ca; }
  @media (prefers-color-scheme: dark) {
    body { color: #e2e8f0; background: #0f172a; }
    h2 { border-color: #334155; }
    code { background: #1e293b; color: #a5b4fc; }
    .note { background: rgba(6,182,212,0.12); }
    .warn { background: rgba(245,158,11,0.12); }
    a { color: #a5b4fc; }
  }
</style>
</head>
<body>
  <h1>Using a custom Client ID &amp; Tenant ID</h1>
  <p>This folder's <code>tenant.json</code> controls which Azure AD (Microsoft
  Entra) application <code>teams-org</code> signs in with.</p>

  <div class="note">
    <strong>By default</strong> the tool uses Microsoft's public
    <em>“Microsoft Graph PowerShell”</em> application
    (<code>clientId&nbsp;=&nbsp;14d82eec-204b-4c2f-b7e8-296a70dab67e</code>,
    <code>tenantId&nbsp;=&nbsp;common</code>). It's multi-tenant and enabled in many
    organizations, so it often works out of the box — you only need your own app
    if your IT department has disabled it or you'd prefer a dedicated registration.
  </div>

  <h2>Step 1 — Register an application</h2>
  <ol>
    <li>Go to the <a href="https://entra.microsoft.com/">Microsoft Entra admin center</a>
        (or the <a href="https://portal.azure.com/">Azure Portal</a>) and sign in.</li>
    <li>Open <strong>Applications → App registrations</strong> and click <strong>New registration</strong>.</li>
    <li>Name it (e.g. <em>“Teams Org Tree Extractor”</em>) and click <strong>Register</strong>.</li>
  </ol>

  <h2>Step 2 — Copy your IDs</h2>
  <p>On the app's <strong>Overview</strong> page, copy:</p>
  <ul>
    <li><strong>Application (client) ID</strong> → this is your <code>clientId</code>.</li>
    <li><strong>Directory (tenant) ID</strong> → this is your <code>tenantId</code>.</li>
  </ul>

  <h2>Step 3 — Allow the device-code flow</h2>
  <ol>
    <li>In the app, open <strong>Authentication</strong>.</li>
    <li>Under <strong>Advanced settings</strong>, set
        <strong>Allow public client flows</strong> to <strong>Yes</strong>.</li>
    <li>Click <strong>Save</strong>.</li>
  </ol>

  <h2>Step 4 — Grant API permissions</h2>
  <ol>
    <li>Open <strong>API permissions → Add a permission → Microsoft Graph →
        Delegated permissions</strong>.</li>
    <li>Add <code>User.Read.All</code>.</li>
    <li>Click <strong>Grant admin consent for [your organization]</strong>
        (an administrator may need to do this).</li>
  </ol>

  <h2>Step 5 — Edit <code>tenant.json</code></h2>
  <p>Open <code>tenant.json</code> in this folder and replace the values:</p>
  <pre><code>{
  "clientId": "YOUR-APPLICATION-CLIENT-ID",
  "tenantId": "YOUR-DIRECTORY-TENANT-ID"
}</code></pre>
  <div class="warn">
    Keep the values as plain strings. Leaving them at the Microsoft defaults means
    the tool will keep asking whether to use the defaults.
  </div>

  <h2>Step 6 — Run the extraction</h2>
  <p>From this folder, run:</p>
  <pre><code>teams-org extract</code></pre>
  <p>Because <code>tenant.json</code> now differs from the defaults,
  <code>teams-org</code> will use your custom IDs automatically and start the
  device-code sign-in.</p>
</body>
</html>
`;
