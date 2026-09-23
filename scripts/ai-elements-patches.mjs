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

const EDITS = {
  "context.tsx": [
    IMPORT_T,
    text("Total cost"),
    text("Input"),
    text("Output"),
    text("Reasoning"),
    text("Cache"),
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
