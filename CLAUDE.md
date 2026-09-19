# Draftspect — project guidance for Claude Code

**Draftspect** — Microsoft Office (Word + Excel) add-ins that wrap a local Claude Code daemon and expose it inside Office via Office.js task panes. Cross-platform (macOS + Windows).

The **product name is "Draftspect"** (independent name, per Anthropic's Agent-SDK branding rule that "Claude Code" must not be the product/marketing name). The sanctioned attribution lockup is **"Powered by Claude"**. Naming convention:

- Product / app name: `Draftspect` — surfaced as `Draftspect for Word`, `Draftspect for Excel`, `Draftspect for Office` (tray), always with the "Powered by Claude" lockup where a tagline fits.
- npm package: `draftspect-office-addins`
- GitHub repo: name/description carries the **descriptive** phrase "Add-Ins for Word and Excel, Powered by Claude Code" — that's nominative use (it factually wraps Claude Code), distinct from naming the _product_ "Claude Code".

**Keep accurate engine references.** Text that tells the user about the _actual_ Claude Code they must install / sign into / whose OAuth is used (README setup, auth-error banner, the Agent SDK running their Claude Code) stays "Claude Code" — that's correct, honest, nominative use; renaming it would be misleading. Only the **product's own name** is Draftspect.

## Architecture (3 paragraphs)

The **daemon** (`daemon/`) is the brain. It wraps the Claude Agent SDK, hosts a WebSocket bridge for the taskpane, and registers an in-process MCP server (`office-tools.mjs`) that exposes Word + Excel tools. The daemon inherits the user's MCP servers from `~/.claude.json` and forwards them into the SDK session (the SDK doesn't read that file by default — see `feedback_sdk_does_not_load_claude_json`).

The **Electron shell** (`app/main.mjs`) is the tray app. It spawns the daemon as a child process over an IPC channel (used for native folder pickers and add-in install/uninstall), restarts it on crash, and surfaces status in a menu-bar / system-tray icon. The shell also handles **auto-sideload** (`app/sideload.mjs`) — on first run it copies manifests into Word's and Excel's wef/ folder on macOS, or registers each manifest under the `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` registry key on Windows (no admin, no network share — a local-path Trusted Catalog is silently ignored by Office). No manual XML drops or Trust Center configuration required.

The **taskpane** (`taskpane/`) is what shows up inside Word/Excel. One shared `taskpane.js` handles both hosts; the active host is detected via `Office.context.host` and the tool dispatcher routes `office_*` calls to Word handlers and `excel_*` calls to Excel handlers. The user picks a **workspace folder** (the agent's cwd, where its `CLAUDE.md` lives) and optionally adds **context files** — paths the agent reads on demand via standard `Read`/`Glob`/`Grep`.

## Key constraints

- **Auth.** Each user clones the repo and runs it on their own machine with their own Claude Code OAuth (or `ANTHROPIC_API_KEY`). That per-user model is sanctioned — and explicitly so from **2026-06-15**, when Anthropic's per-user monthly programmatic credit covers third-party Agent SDK apps run on each user's own subscription (no partner approval needed for the per-user model; Draftspect uses the Agent SDK's own auth, not extracted tokens). What is still **not** allowed: pooling one subscription across many users / a hosted multi-tenant service on someone else's plan. See `feedback_subscription_auth_only` and `reference_anthropic_april_2026_policy`; the README "How usage is charged" row is the user-facing version.
- **Branding.** The product is named **Draftspect** with a **"Powered by Claude"** lockup — an independent name, per Anthropic's Agent-SDK rule that "Claude Code" must not be a product/marketing name. Accurate references to the user's real Claude Code (install, sign-in, OAuth, the SDK driving it) stay as "Claude Code" — nominative/descriptive use is fine and necessary for accuracy. The GitHub repo's name/description may carry the descriptive "…Powered by Claude Code" (it factually wraps Claude Code). See `feedback_product_branding_compliance`.
- **Filesystem-write safety.** A `canUseTool` guard refuses `Write`/`Edit`/`MultiEdit` against `.docx`/`.docm`/`.xlsx`/`.xlsm` paths — the active doc is held by Office with unsaved changes, and a filesystem write would corrupt it. The agent must use `office_*` / `excel_*` tools instead.
- **Feature branches, not main.** Per the user's standing preference, every feature goes on a branch and lands via PR — no direct commits to main.
- **MCP forwarding.** The daemon must forward `~/.claude.json`'s `mcpServers` into the SDK session. The SDK doesn't read that file. If a server is unreachable at daemon startup, the SDK silently drops it for the session's lifetime — restart to retry. See `feedback_sdk_silently_drops_failed_mcp`.

## Useful commands

```bash
npm start              # Launch the Electron tray app (daemon + UI)
npm run dev            # Daemon only, in the terminal (for debugging)
node --check daemon/index.mjs taskpane/shared/taskpane.js app/main.mjs
```

## File layout

- `daemon/index.mjs` — Agent SDK loop, per-(host,document) sessions, permission guard
- `daemon/bridge.mjs` — WebSocket server + per-pane tool-call protocol
- `daemon/office-tools.mjs` — Word + Excel tool defs (zod schemas)
- `daemon/workspace.mjs` — workspace = the document's own folder
- `daemon/context.mjs` — per-workspace CONTEXT-FILES block in CLAUDE.md
- `daemon/sessions.mjs` — per-document/workbook session-id persistence
- `daemon/transcript.mjs` — replay reconstruction from the SDK `.jsonl`
- `daemon/system-paths.mjs` — OS-managed `$HOME` children deny-list
- `daemon/diag.mjs` — opt-in `[diag]` logger (`CC_OFFICE_DEBUG=1`)
- `daemon/system-prompt.md` + `system-prompt-word.md` / `system-prompt-excel.md` — shared + per-host system-prompt append
- `app/main.mjs` — Electron menu-bar shell + daemon lifecycle
- `app/sideload.mjs` — macOS/Windows add-in install/uninstall
- `taskpane/shared/taskpane.js` — host-aware taskpane (Word + Excel branches)
- `taskpane/shared/tools-word.js`, `tools-excel.js` — Office.js tool implementations
- `taskpane/shared/paths.js` — pure path helpers (URL/Windows/POSIX)
- `taskpane/shared/styles.css` — taskpane styles
- `taskpane/word/index.html`, `taskpane/excel/index.html` — per-host entry points
- `manifests/word.xml`, `manifests/excel.xml` — Office Add-in manifests
- `examples/{word,excel}-demo/` — ready-to-run demo workspaces

## Where to read more

- `README.md` — user-facing install/usage/troubleshooting docs.
- Project memories under `~/.claude/projects/<this-project>/memory/` (not in the repo) — non-obvious constraints, Office.js gotchas, SDK behavior notes.
