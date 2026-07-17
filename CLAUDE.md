# CLAUDE.md

Guidance for AI agents working in this repository.

## What this project is

`teams-org` is a Bun/TypeScript **CLI** that extracts a Microsoft Teams / Entra ID
organization tree via the Microsoft Graph API and produces a **fully portable,
self-contained** HTML visualization (D3 `d3-org-chart`).

The headline command is:

```bash
teams-org extract
```

It authenticates (interactive browser sign-in by default; `--device-code` for the
device-code flow), downloads every user + their manager + profile photos, and
writes three files into the chosen output folder:

- `org_tree.json` — the readable hierarchy (**no image blobs**).
- `photos.json` — a separate `{ userId: base64DataUrl }` map (kept out of the
  tree so the tree stays human-readable).
- `index.html` — a portable visualization with the tree, photos, and vendor
  libraries embedded; open it directly via `file://`, no server needed.

## Project structure

```
bin/cli.ts            CLI entry point (Commander). The package `bin` → teams-org.
src/extractor.ts      Graph auth + fetching, tree building, portable-HTML generation.
src/tenant.ts         tenant.json flow (default vs custom IDs) + custom_tenant_id.html guide.
templates/index.html  The visualization template (D3 org chart + settings + overlay).
setup.sh              Installer: checks Bun, runs `bun install`, `bun link`.
package.json          bin/scripts; deps: @azure/msal-node, @microsoft/microsoft-graph-client,
                      commander, @inquirer/prompts.
.env.example          Default CLIENT_ID/TENANT_ID (public MS Graph PowerShell app).
```

## Commands

```bash
./setup.sh                 # install deps + link `teams-org` onto PATH (macOS/Linux/Windows)
bun run extract            # = bun run bin/cli.ts extract
teams-org extract          # after linking
teams-org extract --help   # all flags
bunx tsc --noEmit          # type-check (the primary validation — there is no test suite yet)
```

CLI flags: `--folder`, `--yes`, `--client-id`, `--tenant-id`, `--no-photos`,
`--no-inline`, `--device-code`.

## Architecture & data flow

1. `bin/cli.ts` (`runExtract`): checks prerequisites (Bun), picks the output
   folder (defaults to CWD, prompts via `@inquirer/prompts` unless `--folder`/
   `--yes`), resolves client/tenant IDs via `resolveTenantConfig` (see below),
   then calls into `src/extractor.ts`.
2. `resolveTenantConfig()` (`src/tenant.ts`) manages a per-folder `tenant.json`:
   first run seeds it with the defaults and asks whether to use custom IDs (if so
   it writes `custom_tenant_id.html` and returns `null` so the CLI exits); re-runs
   use the file's IDs if they differ from the defaults, otherwise ask to confirm
   defaults. `--yes` and explicit `--client-id`/`--tenant-id` bypass the prompts.
3. `extractOrgData()` → `authenticate()` (interactive by default, MSAL
   device-code with `--device-code`) → `fetchAllUsers()` (Graph `/users` with
   `$expand=manager`) → `fetchAllPhotos()` → `buildTree()`.
   `fetchAllPhotos()` batches (10 at a time) and requests **120×120 thumbnails**
   (`/users/{id}/photos/120x120/$value`, falling back to full-res
   `/photo/$value`); 404s are ignored. Thumbnails matter: full-res photos bloat
   the embedded `index.html` to hundreds of MB and break `file://` rendering.
4. `generatePortableHtml()` reads `templates/index.html`, optionally inlines the
   CDN `<script>` libs (downloads them; falls back to the CDN link on failure),
   and injects the tree + photos as `<script type="application/json">` tags.
   **Every `html.replace()` here must use a replacement _function_** (`() =>
   ...`), never a replacement string — inlined library source and JSON contain
   `$` sequences (`$&`, `` $` ``, `$'`) that `String.replace` would otherwise
   interpret, silently corrupting/truncating the output (this broke d3 once).
5. `templatePath()` resolves `templates/index.html` relative to the module
   (`import.meta.url`), so it works regardless of CWD — keep it in sync if files
   move.

### The template (`templates/index.html`)

- Dual-mode data loading: uses embedded `#org-data-embedded` / `#org-photos-embedded`
  JSON if present (portable build), otherwise `fetch()`es `org_tree.json` /
  `photos.json` (dev mode). Keep both paths working when editing the loader.
- Photos are merged back onto nodes by id at render time (`photos[node.id]`).
- Cogwheel (⚙️) settings panel lets the user pick **up to 3** fields to show per
  card (persisted in `localStorage`); clicking a card opens a contact overlay
  with all fields. Field definitions live in the `FIELD_DEFS` array — add new
  profile fields there **and** in `src/extractor.ts` (`$select`, `OrgNode`,
  `buildTree`).

## Conventions

- **No `.env` required**: `.env.example`'s public app registration is baked in as
  the default (`DEFAULT_CLIENT_ID` / `DEFAULT_TENANT_ID`). Env vars / flags
  override.
- **Keep the tree readable**: never inline base64 photos into `org_tree.json` —
  they belong in `photos.json` / the embedded photo map.
- When embedding data in HTML, escape it for `<script>` context (`<` → `\u003c`)
  as `generatePortableHtml` does, and inject with a replacement **function**, not
  a string (see data-flow step 4 — protects against `$`-pattern corruption).
- Adding a profile field is a two-place change: `src/extractor.ts` (query +
  `OrgNode` + `buildTree`) and `templates/index.html` (`FIELD_DEFS`).
- After changing `bin`/scripts/paths, re-run `bun link` so the global `teams-org`
  shim points at the right file.

## Bun usage (this repo targets Bun, not Node)

- `bun <file>`, `bun install`, `bun run <script>`, `bunx <pkg>`.
- Bun auto-loads `.env` — do not add `dotenv`.
- Prefer `Bun.file` / `Bun.write` over `node:fs`, and `Bun.$` over `execa`.
- Type-check with `bunx tsc --noEmit`. There is no test suite yet; if you add one,
  use `bun test`.

For Bun API details, read `node_modules/bun-types/docs/**.mdx`.
