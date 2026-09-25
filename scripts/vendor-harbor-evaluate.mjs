// Copies Harbor's hardened SpreadsheetBench Verified grader, verbatim, into
// evals/vendor. It fixes crashes and a vacuous pass in the benchmark's own
// evaluation.py (column-only ranges such as A:G, commas inside quoted sheet
// names, BD2:308-style ranges) while keeping its cell comparison rules, so
// local scores are graded the way Harbor's published parity runs are.
// Usage: node scripts/vendor-harbor-evaluate.mjs [path-to-harbor-checkout]
import { execFileSync } from "node:child_process";
import { copyFileSync, mkdirSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const PINNED = "6cb9ff3167596c456e0b24622d473b59fc9ab6c7";
const EVALUATE =
  "adapters/spreadsheetbench-verified/src/spreadsheetbench_verified/task-template/tests/evaluate.py";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const source = resolve(process.argv[2] ?? join(root, "..", "_sdks", "harbor"));
const git = (...args) => execFileSync("git", ["-C", source, ...args], { encoding: "utf8" }).trim();

const head = git("rev-parse", "HEAD");
if (head !== PINNED) throw new Error(`harbor checkout is at ${head}, expected ${PINNED}`);
if (git("status", "--porcelain", "--", EVALUATE, "LICENSE")) {
  throw new Error("harbor checkout has local changes to the grader or its license");
}

const target = join(root, "evals", "vendor");
mkdirSync(target, { recursive: true });
copyFileSync(join(source, EVALUATE), join(target, "harbor_evaluate.py"));
copyFileSync(join(source, "LICENSE"), join(target, "harbor.LICENSE"));
writeFileSync(
  join(target, "SOURCE.md"),
  `\`harbor_evaluate.py\` is copied verbatim from
https://github.com/harbor-framework/harbor at ${PINNED},
\`${EVALUATE}\`, by \`scripts/vendor-harbor-evaluate.mjs\`.
Apache-2.0; see \`harbor.LICENSE\`. It is imported as a module; its \`main()\` is not used.
`,
);
console.log(`vendored harbor evaluate.py @ ${PINNED.slice(0, 7)} into ${target}`);
