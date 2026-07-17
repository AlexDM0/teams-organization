import {
    PublicClientApplication,
    type Configuration,
    type DeviceCodeRequest,
    type InteractiveRequest,
} from '@azure/msal-node';
import { Client, ResponseType } from '@microsoft/microsoft-graph-client';
import * as path from 'path';
import { fileURLToPath } from 'url';

// Default public app registration (Microsoft Graph PowerShell), see .env.example.
export const DEFAULT_CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
export const DEFAULT_TENANT_ID = 'common';

export interface OrgNode {
    id: string;
    parentId: string | null;
    name: string;
    position: string;
    email: string;
    role: string;
    mobilePhone: string;
    businessPhone: string;
    officeLocation: string;
    city: string;
    state: string;
    country: string;
    companyName: string;
    employeeId: string;
    children: OrgNode[];
}

/** Map of user id -> base64 data URL. Kept separate from the tree for readability. */
export type PhotoMap = Record<string, string>;

export interface OrgData {
    tree: OrgNode[];
    photos: PhotoMap;
}

export interface ExtractOptions {
    clientId: string;
    tenantId: string;
    photos: boolean;
    /** Auth flow to use. Interactive (browser) is the default; device-code is often blocked by Conditional Access. */
    deviceCode?: boolean;
}

/** Opens the given URL in the system default browser (cross-platform). */
async function openBrowser(url: string): Promise<void> {
    const cmd =
        process.platform === 'darwin'
            ? ['open', url]
            : process.platform === 'win32'
              ? ['cmd', '/c', 'start', '', url]
              : ['xdg-open', url];
    try {
        Bun.spawn(cmd, { stdout: 'ignore', stderr: 'ignore' });
    } catch {
        // Fall through — the URL is also printed so the user can open it manually.
    }
    console.log('\n========================================================');
    console.log('Opening your browser to sign in. If it does not open,');
    console.log('paste this URL manually:\n');
    console.log(url);
    console.log('========================================================\n');
}

/** Interactive browser (authorization-code) sign-in via a local loopback redirect. */
async function authenticateInteractive(pca: PublicClientApplication): Promise<string> {
    const interactiveRequest: InteractiveRequest = {
        scopes: ['User.Read.All'],
        openBrowser,
        successTemplate: 'Sign-in successful. You can close this tab and return to the terminal.',
        errorTemplate: 'Sign-in failed. Please return to the terminal and try again.',
    };

    console.log('Starting interactive browser sign-in...');
    const response = await pca.acquireTokenInteractive(interactiveRequest);
    if (response && response.accessToken) {
        console.log('Authentication successful!');
        return response.accessToken;
    }
    throw new Error('No access token returned.');
}

/** Device-code sign-in. Note: frequently blocked by Conditional Access policies. */
async function authenticateDeviceCode(pca: PublicClientApplication): Promise<string> {
    const deviceCodeRequest: DeviceCodeRequest = {
        scopes: ['User.Read.All'],
        deviceCodeCallback: (response) => {
            console.log('\n========================================================');
            console.log(response.message);
            console.log('========================================================\n');
        },
    };

    console.log('Requesting device code...');
    const response = await pca.acquireTokenByDeviceCode(deviceCodeRequest);
    if (response && response.accessToken) {
        console.log('Authentication successful!');
        return response.accessToken;
    }
    throw new Error('No access token returned.');
}

export async function authenticate(
    clientId: string,
    tenantId: string,
    deviceCode = false,
): Promise<string> {
    const msalConfig: Configuration = {
        auth: {
            clientId,
            authority: `https://login.microsoftonline.com/${tenantId}`,
        },
    };

    const pca = new PublicClientApplication(msalConfig);

    return deviceCode ? authenticateDeviceCode(pca) : authenticateInteractive(pca);
}

export async function fetchAllUsers(client: Client): Promise<any[]> {
    let users: any[] = [];
    let nextLink: string | undefined =
        '/users?$select=id,displayName,jobTitle,mail,userPrincipalName,department,officeLocation,mobilePhone,businessPhones,city,state,country,companyName,employeeId&$expand=manager($select=id)';

    console.log('Fetching users...');
    while (nextLink) {
        const response = await client.api(nextLink).get();
        users = users.concat(response.value);
        nextLink = response['@odata.nextLink'];
    }
    return users;
}

// Fetch small thumbnails, not full-resolution photos. Graph can return
// enormous originals (multiple MB each); embedding hundreds of them makes the
// portable HTML unusably large. The cards/overlay only need ~64px avatars.
const PHOTO_SIZE = '120x120';

async function fetchProfilePictureBase64(client: Client, userId: string): Promise<string | null> {
    const toDataUrl = (buf: ArrayBuffer) => `data:image/jpeg;base64,${Buffer.from(buf).toString('base64')}`;

    // Prefer the sized endpoint; fall back to the default photo if that exact
    // thumbnail size isn't available.
    try {
        const response = await client
            .api(`/users/${userId}/photos/${PHOTO_SIZE}/$value`)
            .responseType(ResponseType.ARRAYBUFFER)
            .get();
        return toDataUrl(response);
    } catch {
        // fall through to the default (largest) photo
    }

    try {
        const response = await client
            .api(`/users/${userId}/photo/$value`)
            .responseType(ResponseType.ARRAYBUFFER)
            .get();
        return toDataUrl(response);
    } catch {
        // Many users have no profile picture (404) — that's fine.
        return null;
    }
}

export async function fetchAllPhotos(client: Client, users: any[]): Promise<PhotoMap> {
    const photos: PhotoMap = {};
    console.log('Fetching profile pictures...');

    const batchSize = 10;
    const REPORT_EVERY = 100;
    let reported = 0;
    for (let i = 0; i < users.length; i += batchSize) {
        const batch = users.slice(i, i + batchSize);
        await Promise.all(
            batch.map(async (user) => {
                const photo = await fetchProfilePictureBase64(client, user.id);
                if (photo) photos[user.id] = photo;
            })
        );
        const processed = Math.min(i + batchSize, users.length);
        // Print activity every 100 users (and once at the end), regardless of org size.
        if (processed - reported >= REPORT_EVERY || processed === users.length) {
            process.stdout.write(`\rProcessed ${processed} / ${users.length} photos`);
            reported = processed;
        }
    }
    console.log('\nFinished fetching photos.');
    return photos;
}

export function buildTree(users: any[]): OrgNode[] {
    const nodeMap = new Map<string, OrgNode>();

    for (const user of users) {
        nodeMap.set(user.id, {
            id: user.id,
            parentId: user.manager?.id || null,
            name: user.displayName || 'Unknown',
            position: user.jobTitle || '',
            email: user.mail || user.userPrincipalName || '',
            role: user.department || user.officeLocation || '',
            mobilePhone: user.mobilePhone || '',
            businessPhone: (Array.isArray(user.businessPhones) && user.businessPhones[0]) || '',
            officeLocation: user.officeLocation || '',
            city: user.city || '',
            state: user.state || '',
            country: user.country || '',
            companyName: user.companyName || '',
            employeeId: user.employeeId || '',
            children: [],
        });
    }

    const roots: OrgNode[] = [];
    for (const user of users) {
        const node = nodeMap.get(user.id)!;
        if (node.parentId && nodeMap.has(node.parentId)) {
            nodeMap.get(node.parentId)!.children.push(node);
        } else {
            roots.push(node);
        }
    }
    return roots;
}

/** Runs the full extraction and returns the org tree plus a separate photo map. */
export async function extractOrgData(options: ExtractOptions): Promise<OrgData> {
    const accessToken = await authenticate(options.clientId, options.tenantId, options.deviceCode);

    const client = Client.init({
        authProvider: (done) => done(null, accessToken),
    });

    const users = await fetchAllUsers(client);
    console.log(`Fetched ${users.length} users.`);

    const photos = options.photos ? await fetchAllPhotos(client, users) : {};
    return { tree: buildTree(users), photos };
}

/** Locates the bundled index.html template regardless of CWD. */
export function templatePath(): string {
    const here = path.dirname(fileURLToPath(import.meta.url));
    return path.join(here, '..', 'templates', 'index.html');
}

const CDN_SCRIPTS = [
    'https://d3js.org/d3.v7.min.js',
    'https://cdn.jsdelivr.net/npm/d3-org-chart@3.1.0',
    'https://cdn.jsdelivr.net/npm/d3-flextree@2.1.2/build/d3-flextree.js',
    'https://cdn.jsdelivr.net/npm/fuse.js@7.0.0',
];

async function tryDownload(url: string): Promise<string | null> {
    try {
        const res = await fetch(url);
        if (!res.ok) return null;
        return await res.text();
    } catch {
        return null;
    }
}

/**
 * Builds a fully self-contained index.html: the org tree and photo map are each
 * embedded as JSON <script> tags and (when reachable) the CDN libraries are
 * inlined so the file works offline from a plain file:// path.
 */
export async function generatePortableHtml(
    template: string,
    tree: OrgNode[],
    photos: PhotoMap,
    inlineVendors = true
): Promise<string> {
    let html = template;

    if (inlineVendors) {
        console.log('Inlining vendor libraries for offline portability...');
        for (const url of CDN_SCRIPTS) {
            const code = await tryDownload(url);
            const tag = `<script src="${url}"></script>`;
            if (code) {
                const safe = code.replace(/<\/script>/gi, '<\\/script>');
                // Use a replacement *function* so `$` sequences in the library
                // source (e.g. template literals, `$&`, `$'`) are inserted
                // literally rather than interpreted by String.replace.
                html = html.replace(tag, () => `<script>\n${safe}\n</script>`);
            } else {
                console.warn(`  ! Could not inline ${url} (will keep CDN link)`);
            }
        }
    }

    // Escape the JSON so it can live safely inside a <script> tag.
    const encode = (value: unknown) => JSON.stringify(value).replace(/</g, '\\u003c');
    const dataTag = `<script type="application/json" id="org-data-embedded">${encode(tree)}</script>`;
    const photosTag = `<script type="application/json" id="org-photos-embedded">${encode(photos)}</script>`;

    // Inject the embedded data right before the main application script. A
    // replacement function keeps any `$` in the JSON (names, base64) literal.
    html = html.replace(
        '    <script>\n        (function () {',
        () => `    ${dataTag}\n    ${photosTag}\n    <script>\n        (function () {`
    );

    return html;
}
