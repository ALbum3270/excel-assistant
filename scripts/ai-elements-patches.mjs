// Edits applied to the vendored AI Elements files after they are copied, so the
// copies stay reproducible from the pinned commit. Each edit must match exactly
// once; an upstream change that moves the text fails the vendor step loudly.
import { replaceOnce } from "./office-agents-patches.mjs";

// The few fixed English strings these components render go through the pane's
// translation function (`@pane/i18n`, aliased by scripts/build-pane.mjs). The
// English text is its own key, so only other languages list it.
const IMPORT_T = [
  `"use client";\n`,
  `"use client";\n\nimport { t as translate } from "@pane/i18n";\n`,
];
const text = (word) => [`>${word}<`, `>{translate("${word}")}<`];

// Without a `modelId` there is no price table, so upstream formats `undefined ??
// 0` and every row reads "$0.00" next to a real token count. The pane never
// passes a model id — the one cost figure it shows is the SDK's own, in the
// footer — so print the tokens alone instead of an invented zero.
const noZeroCost = (name) => [
  `  const ${name}CostText = new Intl.NumberFormat("en-US", {
    currency: "USD",
    style: "currency",
  }).format(${name}Cost ?? 0);`,
  `  const ${name}CostText =
    ${name}Cost === undefined
      ? undefined
      : new Intl.NumberFormat("en-US", {
          currency: "USD",
          style: "currency",
        }).format(${name}Cost);`,
];

const EDITS = {
  "context.tsx": [
    IMPORT_T,
    text("Total cost"),
    text("Input"),
    text("Output"),
    text("Reasoning"),
    text("Cache"),
    noZeroCost("input"),
    noZeroCost("output"),
    noZeroCost("reasoning"),
    noZeroCost("cache"),
  ],
  "prompt-input.tsx": [
    IMPORT_T,
    // Some Windows IMEs confirm a candidate with an Enter whose keydown reports
    // keyCode 229 but not isComposing; that Enter must not send the half-typed text.
    [
      "if (isComposing || e.nativeEvent.isComposing) {",
      "if (isComposing || e.nativeEvent.isComposing || e.nativeEvent.keyCode === 229) {",
    ],
    [
      `aria-label={isGenerating ? "Stop" : "Submit"}`,
      `aria-label={translate(isGenerating ? "Stop" : "Submit")}`,
    ],
  ],
};

export function patchAiElement(file, source) {
  const edits = EDITS[file];
  if (!edits) return source;
  const lf = source.replaceAll(String.fromCharCode(13), "");
  return edits.reduce(
    (current, [before, after], index) => replaceOnce(current, `${file}#${index}`, before, after),
    lf,
  );
}

export const PATCHED_FILES = Object.keys(EDITS);
