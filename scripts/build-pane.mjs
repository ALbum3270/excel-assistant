// Bundles the task pane (taskpane/app/pane.tsx) with esbuild and its stylesheet
// with the Tailwind v4 CLI into taskpane/app/dist, which taskpane/excel/index.html
// loads. Usage: node scripts/build-pane.mjs [--dev]
import { execFileSync } from "node:child_process";
import { copyFileSync, mkdirSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const app = join(root, "taskpane", "app");
const out = join(app, "dist");
const dev = process.argv.includes("--dev");
mkdirSync(out, { recursive: true });

const plugin = join(app, "stubs", "streamdown-plugin.js");
await build({
  entryPoints: [join(app, "pane.tsx")],
  outfile: join(out, "pane.js"),
  bundle: true,
  format: "esm",
  jsx: "automatic",
  // Excel on Windows hosts the pane in WebView2 (Chromium).
  target: "chrome110",
  minify: !dev,
  sourcemap: true,
  legalComments: "linked",
  define: { "process.env.NODE_ENV": dev ? '"development"' : '"production"' },
  alias: {
    "@repo/shadcn-ui": join(app, "vendor", "shadcn-ui"),
    "@pane/i18n": join(app, "core", "i18n.js"),
    shiki: join(app, "stubs", "shiki.js"),
    "@streamdown/code": plugin,
    "@streamdown/math": plugin,
    "@streamdown/mermaid": plugin,
  },
  logLevel: "warning",
});

// The stylesheet's @font-face rules point at ./fonts next to it.
mkdirSync(join(out, "fonts"), { recursive: true });
for (const [pkg, file] of [
  ["geist", "geist-latin-wght-normal.woff2"],
  ["geist-mono", "geist-mono-latin-wght-normal.woff2"],
]) {
  copyFileSync(
    join(root, "node_modules", "@fontsource-variable", pkg, "files", file),
    join(out, "fonts", file),
  );
}

execFileSync(
  process.execPath,
  [
    join(root, "node_modules", "@tailwindcss", "cli", "dist", "index.mjs"),
    "-i",
    join(app, "app.css"),
    "-o",
    join(out, "pane.css"),
    ...(dev ? [] : ["--minify"]),
  ],
  { cwd: app, stdio: ["ignore", "ignore", "inherit"] },
);
console.log(`built ${out}`);
