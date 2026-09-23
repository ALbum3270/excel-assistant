// Copies the chosen Vercel AI Elements components, plus the shadcn/ui files they
// import, verbatim into taskpane/app/vendor. The pane's own code only composes them.
// Usage: node scripts/vendor-ai-elements.mjs [path-to-ai-elements-checkout]
import { execFileSync } from "node:child_process";
import { copyFileSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { PATCHED_FILES, patchAiElement } from "./ai-elements-patches.mjs";

const PINNED = "6a9d5b1822ffb10bba4bd97175f01edd7d8651cd";
const COMPONENTS = [
  "confirmation",
  "context",
  "conversation",
  "message",
  "prompt-input",
  "queue",
  "shimmer",
];
// shadcn/ui parts the pane uses directly that no component above imports.
const EXTRA_UI = ["badge", "label", "switch"];

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const source = resolve(process.argv[2] ?? join(root, "..", "_sdks", "ai-elements"));
const target = join(root, "taskpane", "app", "vendor");

const head = execFileSync("git", ["-C", source, "rev-parse", "HEAD"], { encoding: "utf8" }).trim();
if (head !== PINNED) throw new Error(`ai-elements checkout is at ${head}, expected ${PINNED}`);
const dirty = execFileSync("git", ["-C", source, "status", "--porcelain", "packages"], {
  encoding: "utf8",
});
if (dirty.trim()) throw new Error("ai-elements checkout has local changes under packages/");

const IMPORT = /from "(\.\/[^"]+|@repo\/shadcn-ui\/[^"]+)"/g;
const queue = [
  ...COMPONENTS.map((name) => `elements/src/${name}.tsx`),
  ...EXTRA_UI.map((name) => `shadcn-ui/components/ui/${name}.tsx`),
];
const seen = new Set();

function resolveImport(from, spec) {
  if (spec.startsWith("@repo/shadcn-ui/")) {
    const rest = spec.slice("@repo/shadcn-ui/".length);
    for (const ext of [".tsx", ".ts"]) {
      if (existsSync(join(source, "packages", "shadcn-ui", rest + ext)))
        return `shadcn-ui/${rest}${ext}`;
    }
  } else {
    const base = join(dirname(from), spec).replaceAll("\\", "/");
    for (const ext of [".tsx", ".ts"]) {
      if (existsSync(join(source, "packages", base + ext))) return base + ext;
    }
  }
  throw new Error(`cannot resolve ${spec} from ${from}`);
}

while (queue.length) {
  const file = queue.shift();
  if (seen.has(file)) continue;
  seen.add(file);
  const text = readFileSync(join(source, "packages", file), "utf8");
  for (const [, spec] of text.matchAll(IMPORT)) queue.push(resolveImport(file, spec));
}

rmSync(target, { recursive: true, force: true });
for (const file of [...seen].sort()) {
  const out = join(target, file.replace(/^elements\/src\//, "ai-elements/"));
  mkdirSync(dirname(out), { recursive: true });
  copyFileSync(join(source, "packages", file), out);
}
const patched = [];
for (const file of PATCHED_FILES) {
  const path = join(target, "ai-elements", file);
  writeFileSync(path, patchAiElement(file, readFileSync(path, "utf8")));
  patched.push(`ai-elements/${file}`);
}
copyFileSync(join(source, "LICENSE"), join(target, "LICENSE"));

// Theme tokens (@theme inline, :root, .dark: colors, radii, layered shadows and
// easing; plus the base layers) from Vercel's chatbot template, which is built
// on the same shadcn/ui components. Its Tailwind plugins and imports are left out.
const THEME_PINNED = "c2f8235e1f3ea903ad8b7f61447c4f74164b5c58";
const themeSource = resolve(process.argv[3] ?? join(root, "..", "_sdks", "vercel-chatbot"));
const themeHead = execFileSync("git", ["-C", themeSource, "rev-parse", "HEAD"], {
  encoding: "utf8",
}).trim();
if (themeHead !== THEME_PINNED) {
  throw new Error(`vercel chatbot checkout is at ${themeHead}, expected ${THEME_PINNED}`);
}
const globalsCss = readFileSync(join(themeSource, "app", "globals.css"), "utf8").replaceAll(
  String.fromCharCode(13),
  "",
);
const themeStart = globalsCss.indexOf("@custom-variant dark");
const themeEnd = globalsCss.indexOf("@layer base {\n  body {\n    position: relative;");
if (themeStart < 0 || themeEnd < themeStart)
  throw new Error("theme markers not found in globals.css");
writeFileSync(
  join(target, "theme.css"),
  `/* From vercel/ai-chatbot app/globals.css at ${THEME_PINNED} (Apache-2.0) */\n` +
    globalsCss
      .slice(themeStart, themeEnd)
      .split("\n")
      .filter((line) => !line.startsWith("@plugin"))
      .join("\n")
      .trimEnd() +
    "\n",
);
writeFileSync(
  join(target, "SOURCE.md"),
  `Copied verbatim from https://github.com/vercel/ai-elements at ${PINNED}\n` +
    `by scripts/vendor-ai-elements.mjs. Apache-2.0 (see LICENSE); the shadcn-ui\n` +
    `files are shadcn/ui components (MIT) as kept in that repository.\n` +
    `Patched after copying (scripts/ai-elements-patches.mjs): ${patched.join(", ")}.\n\n` +
    [...seen]
      .sort()
      .map((file) => `- packages/${file}`)
      .join("\n") +
    "\n",
);
console.log(`vendored ${seen.size} files into ${target}`);
