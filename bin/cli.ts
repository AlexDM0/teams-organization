#!/usr/bin/env bun
import { Command } from 'commander';
import { input, confirm } from '@inquirer/prompts';
import * as path from 'path';
import {
    DEFAULT_CLIENT_ID,
    DEFAULT_TENANT_ID,
    extractOrgData,
    generatePortableHtml,
    templatePath,
} from '../src/extractor';
import { resolveTenantConfig } from '../src/tenant';

interface Tool {
    name: string;
    check: () => Promise<boolean>;
    install: string;
}

const REQUIRED_TOOLS: Tool[] = [
    {
        name: 'bun',
        check: async () => typeof (globalThis as any).Bun !== 'undefined' || (await which('bun')),
        install:
            'Install Bun:  curl -fsSL https://bun.sh/install | bash\n' +
            '  (or on Windows PowerShell:  powershell -c "irm bun.sh/install.ps1 | iex")\n' +
            '  More info: https://bun.sh/',
    },
];

async function which(bin: string): Promise<boolean> {
    try {
        const proc = Bun.spawn([process.platform === 'win32' ? 'where' : 'which', bin], {
            stdout: 'ignore',
            stderr: 'ignore',
        });
        return (await proc.exited) === 0;
    } catch {
        return false;
    }
}

async function checkPrerequisites(): Promise<boolean> {
    const missing: Tool[] = [];
    for (const tool of REQUIRED_TOOLS) {
        if (!(await tool.check())) missing.push(tool);
    }

    if (missing.length === 0) return true;

    console.error('\n❌ Missing required tools:\n');
    for (const tool of missing) {
        console.error(`  • ${tool.name}`);
        console.error(`    ${tool.install.split('\n').join('\n    ')}\n`);
    }
    return false;
}

async function runExtract(opts: {
    folder?: string;
    yes?: boolean;
    clientId: string;
    tenantId: string;
    photos: boolean;
    inline: boolean;
    deviceCode?: boolean;
}) {
    console.log('🔎 Checking prerequisites...');
    if (!(await checkPrerequisites())) {
        console.error('Please install the tools above and re-run `teams-org extract`.');
        process.exit(1);
    }
    console.log('✅ All required tools are installed.\n');

    // Pick output folder — default to the current working directory.
    let folder = opts.folder ?? process.cwd();
    if (!opts.folder && !opts.yes) {
        folder = await input({
            message: 'Where should the extraction be saved?',
            default: process.cwd(),
        });
    }
    folder = path.resolve(folder);
    await Bun.$`mkdir -p ${folder}`.quiet();
    console.log(`📁 Output folder: ${folder}\n`);

    // Resolve which client/tenant IDs to use via tenant.json in the folder.
    const tenant = await resolveTenantConfig(folder, {
        yes: opts.yes,
        cliClientId: opts.clientId,
        cliTenantId: opts.tenantId,
    });
    if (!tenant) {
        // User chose to customize their IDs first — nothing more to do.
        return;
    }

    // Extract data from Microsoft Graph.
    const { tree, photos } = await extractOrgData({
        clientId: tenant.clientId,
        tenantId: tenant.tenantId,
        photos: opts.photos,
        deviceCode: opts.deviceCode,
    });

    // Write the readable tree and the separate photo map.
    const treePath = path.join(folder, 'org_tree.json');
    const photosPath = path.join(folder, 'photos.json');
    await Bun.write(treePath, JSON.stringify(tree, null, 2));
    // Photos are opaque base64 blobs — write them compact (no pretty-print) to
    // roughly halve the file size and the transient string used to serialize it.
    await Bun.write(photosPath, JSON.stringify(photos));
    console.log(`💾 Saved ${treePath}`);
    console.log(`💾 Saved ${photosPath} (${Object.keys(photos).length} photos)`);

    // Generate the fully portable, self-contained index.html.
    const template = await Bun.file(templatePath()).text();
    const html = await generatePortableHtml(template, tree, photos, opts.inline);
    const htmlPath = path.join(folder, 'index.html');
    await Bun.write(htmlPath, html);
    console.log(`💾 Saved ${htmlPath}`);

    console.log(`\n✨ Done! Open the portable file directly in your browser:\n   ${htmlPath}`);
}

async function runUpdate(opts: { folder?: string; inline: boolean }) {
    const folder = path.resolve(opts.folder ?? process.cwd());
    console.log(`📁 Folder: ${folder}\n`);

    // Reuse the existing extraction data — the tree is required.
    const treePath = path.join(folder, 'org_tree.json');
    const photosPath = path.join(folder, 'photos.json');
    const treeFile = Bun.file(treePath);
    if (!(await treeFile.exists())) {
        console.error(`❌ No org_tree.json found in ${folder}.`);
        console.error('   Run `teams-org extract` here first.');
        process.exit(1);
    }
    const tree = JSON.parse(await treeFile.text());

    // photos.json is optional (e.g. extracted with --no-photos).
    const photosFile = Bun.file(photosPath);
    let photos: Record<string, string> = {};
    if (await photosFile.exists()) {
        photos = JSON.parse(await photosFile.text());
        console.log(`🖼️  Reusing ${photosPath} (${Object.keys(photos).length} photos)`);
    } else {
        console.log('🖼️  No photos.json found — building without photos.');
    }
    console.log(`🌳 Reusing ${treePath}\n`);

    // Rebake the portable HTML from the current template.
    const template = await Bun.file(templatePath()).text();
    const html = await generatePortableHtml(template, tree, photos, opts.inline);
    const htmlPath = path.join(folder, 'index.html');
    await Bun.write(htmlPath, html);
    console.log(`💾 Saved ${htmlPath}`);

    console.log(`\n✨ Done! Open the refreshed file directly in your browser:\n   ${htmlPath}`);
}

const program = new Command();

program
    .name('teams-org')
    .description('Extract a Microsoft Teams / Entra ID organization tree into a portable HTML visualization.')
    .version('1.0.0');

program
    .command('extract')
    .description('Authenticate, download the org tree, and build a portable index.html.')
    .option('-f, --folder <path>', 'Output folder (default: current directory)')
    .option('-y, --yes', 'Skip prompts and use defaults')
    .option('--client-id <id>', 'Azure AD application (client) ID (saved to tenant.json)', process.env.CLIENT_ID || DEFAULT_CLIENT_ID)
    .option('--tenant-id <id>', 'Azure AD tenant ID (saved to tenant.json)', process.env.TENANT_ID || DEFAULT_TENANT_ID)
    .option('--no-photos', 'Skip downloading profile pictures')
    .option('--no-inline', 'Keep CDN <script> links instead of inlining libraries')
    .option('--device-code', 'Use the device-code flow instead of interactive browser sign-in')
    .action(async (options) => {
        try {
            await runExtract({
                folder: options.folder,
                yes: options.yes,
                clientId: options.clientId,
                tenantId: options.tenantId,
                photos: options.photos,
                inline: options.inline,
                deviceCode: options.deviceCode,
            });
        } catch (err: any) {
            console.error('\n❌ Extraction failed:', err?.message ?? err);
            process.exit(1);
        }
    });

program
    .command('update')
    .description('Rebuild index.html from the existing org_tree.json / photos.json (no re-download).')
    .option('-f, --folder <path>', 'Folder containing the extraction (default: current directory)')
    .option('--no-inline', 'Keep CDN <script> links instead of inlining libraries')
    .action(async (options) => {
        try {
            await runUpdate({ folder: options.folder, inline: options.inline });
        } catch (err: any) {
            console.error('\n❌ Update failed:', err?.message ?? err);
            process.exit(1);
        }
    });

program.parseAsync();
