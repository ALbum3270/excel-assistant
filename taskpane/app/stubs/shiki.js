// Stand-in for shiki in the pane bundle: code blocks fall back to the plain tokens
// AI Elements shows while highlighting loads (code-block.tsx catches this rejection).
export const createHighlighter = () =>
  Promise.reject(new Error("syntax highlighting is not bundled"));
