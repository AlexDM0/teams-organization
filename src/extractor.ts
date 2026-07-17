import {
    PublicClientApplication,
    type AuthenticationResult,
    type Configuration as MsalConfiguration,
    type DeviceCodeRequest,
    type InteractiveRequest,
} from '@azure/msal-node';
import { Client, ResponseType } from '@microsoft/microsoft-graph-client';
import * as path from 'path';
import { fileURLToPath } from 'url';

// Default public app registration (Microsoft Graph PowerShell). This is
// Microsoft's multi-tenant "Microsoft Graph PowerShell" application, which is
// enabled in many organizations, so the tool works out of the box. Override it
// per-folder via tenant.json or the --client-id/--tenant-id flags.
export const DEFAULT_CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
export const DEFAULT_TENANT_ID = 'common';

/** Microsoft Graph delegated permission the tool signs in with. */
const GRAPH_SCOPES = ['User.Read.All'];

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

/**
 * Supplies a currently-valid Microsoft Graph access token. The initial sign-in
 * happens once; afterwards tokens are refreshed silently so long extractions
 * (which can outlive a token's ~1 hour lifetime) keep working.
 */
interface AccessTokenProvider {
    getAccessToken(): Promise<string>;
}

/** Opens the given URL in the system default browser (cross-platform). */
async function openBrowser(signInUrl: string): Promise<void> {
    const openCommand =
        process.platform === 'darwin'
            ? ['open', signInUrl]
            : process.platform === 'win32'
              ? ['cmd', '/c', 'start', '', signInUrl]
              : ['xdg-open', signInUrl];
    try {
        Bun.spawn(openCommand, { stdout: 'ignore', stderr: 'ignore' });
    } catch {
        // Fall through — the URL is also printed so the user can open it manually.
    }
    console.log('\n========================================================');
    console.log('Opening your browser to sign in. If it does not open,');
    console.log('paste this URL manually:\n');
    console.log(signInUrl);
    console.log('========================================================\n');
}

/** Interactive browser (authorization-code) sign-in via a local loopback redirect. */
async function authenticateWithInteractiveBrowser(
    publicClientApplication: PublicClientApplication,
): Promise<AuthenticationResult> {
    const interactiveRequest: InteractiveRequest = {
        scopes: GRAPH_SCOPES,
        openBrowser,
        successTemplate: 'Sign-in successful. You can close this tab and return to the terminal.',
        errorTemplate: 'Sign-in failed. Please return to the terminal and try again.',
    };

    console.log('Starting interactive browser sign-in...');
    const authenticationResult = await publicClientApplication.acquireTokenInteractive(interactiveRequest);
    if (authenticationResult && authenticationResult.accessToken) {
        console.log('Authentication successful!');
        return authenticationResult;
    }
    throw new Error('No access token returned.');
}

/** Device-code sign-in. Note: frequently blocked by Conditional Access policies. */
async function authenticateWithDeviceCode(
    publicClientApplication: PublicClientApplication,
): Promise<AuthenticationResult> {
    const deviceCodeRequest: DeviceCodeRequest = {
        scopes: GRAPH_SCOPES,
        deviceCodeCallback: (deviceCodeResponse) => {
            console.log('\n========================================================');
            console.log(deviceCodeResponse.message);
            console.log('========================================================\n');
        },
    };

    console.log('Requesting device code...');
    const authenticationResult = await publicClientApplication.acquireTokenByDeviceCode(deviceCodeRequest);
    if (authenticationResult && authenticationResult.accessToken) {
        console.log('Authentication successful!');
        return authenticationResult;
    }
    throw new Error('No access token returned.');
}

/**
 * Signs in once (interactive by default, device-code on request) and returns a
 * provider that hands out valid access tokens, transparently refreshing them
 * via the cached account when they expire.
 */
export async function createAccessTokenProvider(
    clientId: string,
    tenantId: string,
    useDeviceCodeFlow = false,
): Promise<AccessTokenProvider> {
    const msalConfiguration: MsalConfiguration = {
        auth: {
            clientId,
            authority: `https://login.microsoftonline.com/${tenantId}`,
        },
    };

    const publicClientApplication = new PublicClientApplication(msalConfiguration);
    const initialAuthentication = useDeviceCodeFlow
        ? await authenticateWithDeviceCode(publicClientApplication)
        : await authenticateWithInteractiveBrowser(publicClientApplication);
    const signedInAccount = initialAuthentication.account;

    return {
        async getAccessToken(): Promise<string> {
            // Prefer a silently-refreshed token so long runs survive token expiry.
            if (signedInAccount) {
                try {
                    const refreshedAuthentication = await publicClientApplication.acquireTokenSilent({
                        account: signedInAccount,
                        scopes: GRAPH_SCOPES,
                    });
                    if (refreshedAuthentication && refreshedAuthentication.accessToken) {
                        return refreshedAuthentication.accessToken;
                    }
                } catch {
                    // Silent refresh failed — fall back to the original token below.
                }
            }
            return initialAuthentication.accessToken;
        },
    };
}

export async function fetchAllUsers(graphClient: Client): Promise<any[]> {
    let users: any[] = [];
    let nextPageLink: string | undefined =
        '/users?$select=id,displayName,jobTitle,mail,userPrincipalName,department,officeLocation,mobilePhone,businessPhones,city,state,country,companyName,employeeId&$expand=manager($select=id)';

    console.log('Fetching users...');
    while (nextPageLink) {
        const usersPage = await graphClient.api(nextPageLink).get();
        users = users.concat(usersPage.value);
        nextPageLink = usersPage['@odata.nextLink'];
    }
    return users;
}

// Fetch small thumbnails, not full-resolution photos. Graph can return
// enormous originals (multiple MB each); embedding hundreds of them makes the
// portable HTML unusably large. The cards/overlay only need ~64px avatars.
const THUMBNAIL_SIZE = '120x120';

/** How many times to retry a request that Graph throttled (HTTP 429/503). */
const MAX_THROTTLE_RETRIES = 4;

/** Outcome of a single Graph photo request, distinguishing "missing" from "failed". */
type PhotoRequestOutcome =
    | { kind: 'success'; imageBytes: ArrayBuffer }
    | { kind: 'notFound' }
    | { kind: 'error' };

/** Sleeps for the given number of milliseconds. */
function delay(milliseconds: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

/** True when a Graph error represents throttling that is worth retrying. */
function isThrottlingError(error: any): boolean {
    return error?.statusCode === 429 || error?.statusCode === 503;
}

/** True when a Graph error means the requested resource simply does not exist. */
function isNotFoundError(error: any): boolean {
    return error?.statusCode === 404;
}

/**
 * Reads the number of milliseconds Graph asked us to wait before retrying,
 * honoring a `Retry-After` header when present and otherwise backing off
 * exponentially (1s, 2s, 4s, ...).
 */
function throttleWaitMilliseconds(error: any, retryAttempt: number): number {
    const retryAfterHeader = error?.headers?.get?.('Retry-After') ?? error?.headers?.['Retry-After'];
    const retryAfterSeconds = Number(retryAfterHeader);
    if (Number.isFinite(retryAfterSeconds) && retryAfterSeconds > 0) {
        return retryAfterSeconds * 1000;
    }
    return 1000 * 2 ** retryAttempt;
}

/**
 * Requests the raw bytes of a photo from a Graph endpoint. Retries on
 * throttling, reports a genuine 404 as `notFound`, and reports any other
 * failure as `error` (so callers do not mistake an outage for a missing photo).
 */
async function requestPhotoBytes(graphClient: Client, photoEndpoint: string): Promise<PhotoRequestOutcome> {
    for (let retryAttempt = 0; retryAttempt <= MAX_THROTTLE_RETRIES; retryAttempt++) {
        try {
            const imageBytes = await graphClient
                .api(photoEndpoint)
                .responseType(ResponseType.ARRAYBUFFER)
                .get();
            return { kind: 'success', imageBytes };
        } catch (error) {
            if (isNotFoundError(error)) {
                return { kind: 'notFound' };
            }
            if (isThrottlingError(error) && retryAttempt < MAX_THROTTLE_RETRIES) {
                await delay(throttleWaitMilliseconds(error, retryAttempt));
                continue;
            }
            return { kind: 'error' };
        }
    }
    return { kind: 'error' };
}

/**
 * Detects an image's MIME type from its leading "magic number" bytes so the
 * generated data URL advertises the correct format instead of assuming JPEG.
 */
function detectImageMimeType(imageBytes: ArrayBuffer): string {
    const header = new Uint8Array(imageBytes.slice(0, 12));
    const startsWith = (...signature: number[]) => signature.every((byte, index) => header[index] === byte);

    if (startsWith(0xff, 0xd8, 0xff)) return 'image/jpeg';
    if (startsWith(0x89, 0x50, 0x4e, 0x47)) return 'image/png';
    if (startsWith(0x47, 0x49, 0x46)) return 'image/gif';
    if (startsWith(0x42, 0x4d)) return 'image/bmp';
    // WEBP: "RIFF" .... "WEBP"
    if (
        startsWith(0x52, 0x49, 0x46, 0x46) &&
        header[8] === 0x57 &&
        header[9] === 0x45 &&
        header[10] === 0x42 &&
        header[11] === 0x50
    ) {
        return 'image/webp';
    }
    return 'image/jpeg';
}

/** Encodes raw image bytes as a base64 data URL with the detected MIME type. */
function encodeImageAsDataUrl(imageBytes: ArrayBuffer): string {
    const mimeType = detectImageMimeType(imageBytes);
    const base64 = Buffer.from(imageBytes).toString('base64');
    return `data:${mimeType};base64,${base64}`;
}

/**
 * Fetches a single user's profile picture as a base64 data URL. It prefers the
 * small square thumbnail and only falls back to the full-resolution photo when
 * that exact thumbnail genuinely does not exist (a 404) — never on a transient
 * error, which would risk pulling multi-megabyte originals into the HTML.
 */
async function fetchProfilePhotoDataUrl(graphClient: Client, userId: string): Promise<string | null> {
    const thumbnailOutcome = await requestPhotoBytes(graphClient, `/users/${userId}/photos/${THUMBNAIL_SIZE}/$value`);
    if (thumbnailOutcome.kind === 'success') {
        return encodeImageAsDataUrl(thumbnailOutcome.imageBytes);
    }
    if (thumbnailOutcome.kind === 'error') {
        // A transient failure — do not fall back to the (potentially huge) original.
        return null;
    }

    // The sized thumbnail does not exist; try the user's default photo instead.
    const defaultPhotoOutcome = await requestPhotoBytes(graphClient, `/users/${userId}/photo/$value`);
    if (defaultPhotoOutcome.kind === 'success') {
        return encodeImageAsDataUrl(defaultPhotoOutcome.imageBytes);
    }
    // Many users simply have no profile picture — that's fine.
    return null;
}

/** Renders a single-line terminal progress bar with a rough ETA. */
function renderProgress(completedCount: number, totalCount: number, startMilliseconds: number): void {
    const completionRatio = totalCount > 0 ? completedCount / totalCount : 1;
    const barWidth = 30;
    const filledWidth = Math.round(completionRatio * barWidth);
    const progressBar = '█'.repeat(filledWidth) + '░'.repeat(barWidth - filledWidth);
    const percentText = String(Math.floor(completionRatio * 100)).padStart(3);

    // Estimate remaining time from the average pace so far.
    const elapsedSeconds = (Date.now() - startMilliseconds) / 1000;
    const completedPerSecond = completedCount > 0 ? completedCount / elapsedSeconds : 0;
    const remainingSeconds = completedPerSecond > 0 ? Math.round((totalCount - completedCount) / completedPerSecond) : 0;
    const etaText = completedCount >= totalCount ? 'done' : `ETA ${formatDuration(remainingSeconds)}`;

    process.stdout.write(`\r  [${progressBar}] ${percentText}%  ${completedCount}/${totalCount}  ${etaText}   `);
}

/** Formats seconds as a compact m:ss (or h:mm:ss) string. */
function formatDuration(totalSeconds: number): string {
    const clampedSeconds = Math.max(0, totalSeconds);
    const hours = Math.floor(clampedSeconds / 3600);
    const minutes = Math.floor((clampedSeconds % 3600) / 60);
    const seconds = Math.floor(clampedSeconds % 60);
    if (hours > 0) return `${hours}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
    return `${minutes}:${String(seconds).padStart(2, '0')}`;
}

export async function fetchAllPhotos(graphClient: Client, users: any[]): Promise<PhotoMap> {
    const photos: PhotoMap = {};
    console.log('Fetching profile pictures...');

    const batchSize = 10;
    const reportEveryUsers = 100;
    const startMilliseconds = Date.now();
    let lastReportedCount = 0;
    renderProgress(0, users.length, startMilliseconds);
    for (let batchStartIndex = 0; batchStartIndex < users.length; batchStartIndex += batchSize) {
        const userBatch = users.slice(batchStartIndex, batchStartIndex + batchSize);
        await Promise.all(
            userBatch.map(async (user) => {
                const photoDataUrl = await fetchProfilePhotoDataUrl(graphClient, user.id);
                if (photoDataUrl) photos[user.id] = photoDataUrl;
            })
        );
        const processedCount = Math.min(batchStartIndex + batchSize, users.length);
        // Refresh the bar every 100 users (and once at the end), regardless of org size.
        if (processedCount - lastReportedCount >= reportEveryUsers || processedCount === users.length) {
            renderProgress(processedCount, users.length, startMilliseconds);
            lastReportedCount = processedCount;
        }
    }
    console.log('\nFinished fetching photos.');
    return photos;
}

export function buildTree(users: any[]): OrgNode[] {
    const nodesByUserId = new Map<string, OrgNode>();

    for (const user of users) {
        const managerId: string | null = user.manager?.id || null;
        nodesByUserId.set(user.id, {
            id: user.id,
            // A user listed as their own manager would create a self-cycle; treat them as a root instead.
            parentId: managerId && managerId !== user.id ? managerId : null,
            name: user.displayName || 'Unknown',
            position: user.jobTitle || '',
            email: user.mail || user.userPrincipalName || '',
            role: user.department || user.officeLocation || '',
            mobilePhone: user.mobilePhone || '',
            businessPhone: (Array.isArray(user.businessPhones) ? user.businessPhones.filter(Boolean).join(', ') : '') || '',
            officeLocation: user.officeLocation || '',
            city: user.city || '',
            state: user.state || '',
            country: user.country || '',
            companyName: user.companyName || '',
            employeeId: user.employeeId || '',
            children: [],
        });
    }

    const rootNodes: OrgNode[] = [];
    for (const user of users) {
        const node = nodesByUserId.get(user.id)!;
        const parentNode = node.parentId ? nodesByUserId.get(node.parentId) : undefined;
        if (parentNode && !createsManagerCycle(node, parentNode, nodesByUserId)) {
            parentNode.children.push(node);
        } else {
            // No known manager, or attaching would form a cycle — treat as a root.
            node.parentId = null;
            rootNodes.push(node);
        }
    }
    return rootNodes;
}

/**
 * Detects whether making `parentNode` the parent of `node` would form a cycle
 * (e.g. A manages B while B manages A) by walking up the manager chain and
 * checking whether we arrive back at `node`.
 */
function createsManagerCycle(node: OrgNode, parentNode: OrgNode, nodesByUserId: Map<string, OrgNode>): boolean {
    let ancestor: OrgNode | undefined = parentNode;
    const visitedIds = new Set<string>();
    while (ancestor) {
        if (ancestor.id === node.id) return true;
        if (visitedIds.has(ancestor.id)) return true;
        visitedIds.add(ancestor.id);
        ancestor = ancestor.parentId ? nodesByUserId.get(ancestor.parentId) : undefined;
    }
    return false;
}

/** Runs the full extraction and returns the org tree plus a separate photo map. */
export async function extractOrgData(options: ExtractOptions): Promise<OrgData> {
    const accessTokenProvider = await createAccessTokenProvider(
        options.clientId,
        options.tenantId,
        options.deviceCode,
    );

    const graphClient = Client.init({
        authProvider: async (done) => {
            try {
                done(null, await accessTokenProvider.getAccessToken());
            } catch (error) {
                done(error as Error, null);
            }
        },
    });

    const users = await fetchAllUsers(graphClient);
    console.log(`Fetched ${users.length} users.`);

    const photos = options.photos ? await fetchAllPhotos(graphClient, users) : {};
    return { tree: buildTree(users), photos };
}

/** Locates the bundled index.html template regardless of CWD. */
export function templatePath(): string {
    const moduleDirectory = path.dirname(fileURLToPath(import.meta.url));
    return path.join(moduleDirectory, '..', 'templates', 'index.html');
}

/**
 * A vendored front-end library loaded from a CDN. The pinned URL and Subresource
 * Integrity hash let us verify the downloaded bytes before inlining them (and let
 * the browser verify them when the CDN <script> tags are kept via --no-inline).
 */
interface VendorScript {
    url: string;
    integrity: string;
}

const VENDOR_SCRIPTS: VendorScript[] = [
    {
        url: 'https://cdn.jsdelivr.net/npm/d3@7.9.0/dist/d3.min.js',
        integrity: 'sha384-CjloA8y00+1SDAUkjs099PVfnY2KmDC2BZnws9kh8D/lX1s46w6EPhpXdqMfjK6i',
    },
    {
        url: 'https://cdn.jsdelivr.net/npm/d3-org-chart@3.1.0',
        integrity: 'sha384-jdNdDz72Yl0UnyMxO94TJN0XhwcJ5fB5IBuMJ3SWALIOOuR2XoRugVBewvDI9rH1',
    },
    {
        url: 'https://cdn.jsdelivr.net/npm/d3-flextree@2.1.2/build/d3-flextree.js',
        integrity: 'sha384-6pTgblH+kfP7e8kLkJxI96n+G6MCr28XHUtlXyr3cSjSyT/co6eOBwwCPAX8pBb5',
    },
    {
        url: 'https://cdn.jsdelivr.net/npm/fuse.js@7.0.0',
        integrity: 'sha384-PCSoOZTpbkikBEtd/+uV3WNdc676i9KUf01KOA8CnJotvlx8rRrETbDuwdjqTYvt',
    },
];

/** Placeholder comment in the template where the embedded JSON data is injected. */
const DATA_PLACEHOLDER = '<!-- ORG_DATA_PLACEHOLDER -->';

/** Result of downloading a vendor library: its source and the integrity hash of the bytes. */
interface DownloadedVendorScript {
    sourceCode: string;
    integrity: string;
}

/** Computes the Subresource Integrity hash (sha384) for a block of bytes. */
async function computeSubresourceIntegrity(bytes: ArrayBuffer): Promise<string> {
    const digest = await crypto.subtle.digest('SHA-384', bytes);
    return `sha384-${Buffer.from(digest).toString('base64')}`;
}

/**
 * Downloads a vendor library and verifies it against the pinned integrity hash.
 * Returns `null` if the download fails (the CDN <script> tag is then kept as-is,
 * so the browser still enforces integrity). Throws if the bytes are reachable
 * but do NOT match the pinned hash — we never inline unverified third-party code.
 */
async function downloadVendorScript(vendorScript: VendorScript): Promise<DownloadedVendorScript | null> {
    let bytes: ArrayBuffer;
    try {
        const response = await fetch(vendorScript.url);
        if (!response.ok) return null;
        bytes = await response.arrayBuffer();
    } catch {
        return null;
    }

    const actualIntegrity = await computeSubresourceIntegrity(bytes);
    if (actualIntegrity !== vendorScript.integrity) {
        throw new Error(
            `Integrity check failed for ${vendorScript.url}.\n` +
                `  expected: ${vendorScript.integrity}\n` +
                `  actual:   ${actualIntegrity}\n` +
                'Refusing to inline unverified third-party code. If the library was intentionally ' +
                'updated, update its pinned integrity hash in VENDOR_SCRIPTS.'
        );
    }

    return { sourceCode: new TextDecoder().decode(bytes), integrity: vendorScript.integrity };
}

/** Builds a regex matching the template's <script> tag for a given library URL. */
function vendorScriptTagPattern(url: string): RegExp {
    const escapedUrl = url.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    return new RegExp(`<script\\b[^>]*\\bsrc="${escapedUrl}"[^>]*></script>`, 'i');
}

/**
 * Replaces the first match of `pattern` in `html` using a replacement *function*
 * (so `$` sequences in the replacement are inserted literally), and asserts that
 * the replacement actually happened so a template change can never silently
 * produce a broken portable file.
 */
function replaceOrThrow(html: string, pattern: string | RegExp, replacement: string, description: string): string {
    const updatedHtml = html.replace(pattern as any, () => replacement);
    if (updatedHtml === html) {
        throw new Error(`Could not ${description} — the expected anchor was not found in templates/index.html.`);
    }
    return updatedHtml;
}

/**
 * Builds a fully self-contained index.html: the org tree and photo map are each
 * embedded as JSON <script> tags and (when reachable and integrity-verified) the
 * CDN libraries are inlined so the file works offline from a plain file:// path.
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
        for (const vendorScript of VENDOR_SCRIPTS) {
            const downloaded = await downloadVendorScript(vendorScript);
            if (!downloaded) {
                console.warn(`  ! Could not inline ${vendorScript.url} (will keep the verified CDN link)`);
                continue;
            }
            // Neutralize any literal </script> in the library source so it can't
            // terminate the surrounding <script> element early.
            const safeSourceCode = downloaded.sourceCode.replace(/<\/script>/gi, '<\\/script>');
            html = replaceOrThrow(
                html,
                vendorScriptTagPattern(vendorScript.url),
                `<script>\n${safeSourceCode}\n</script>`,
                `inline vendor library ${vendorScript.url}`
            );
        }
    }

    // Escape the JSON so it can live safely inside a <script> tag.
    const encodeForScriptTag = (value: unknown) => JSON.stringify(value).replace(/</g, '\\u003c');
    const treeDataTag = `<script type="application/json" id="org-data-embedded">${encodeForScriptTag(tree)}</script>`;
    const photosDataTag = `<script type="application/json" id="org-photos-embedded">${encodeForScriptTag(photos)}</script>`;

    // Inject the embedded data at the stable placeholder comment. A replacement
    // function keeps any `$` in the JSON (names, base64) literal.
    html = replaceOrThrow(
        html,
        DATA_PLACEHOLDER,
        `${treeDataTag}\n    ${photosDataTag}`,
        'inject embedded org data'
    );

    return html;
}
