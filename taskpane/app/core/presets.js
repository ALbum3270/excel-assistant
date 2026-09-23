// Saved prompts. Built-in ones are stored by id only and shown (and sent) in the
// pane's current language; editing one turns it into the user's own text.
import { t } from "./i18n.js";

export const BUILTIN_PRESETS = [
  { builtin: "summarize", pinned: true, auto_send: true },
  { builtin: "explainSelection", pinned: true, auto_send: true },
  { builtin: "totalsRow", pinned: false, auto_send: true },
  { builtin: "checkData", pinned: false, auto_send: true },
  { builtin: "findValue", pinned: false, auto_send: false },
  { builtin: "contextFiles", pinned: false, auto_send: false },
];

export function defaultPresets(newId) {
  return BUILTIN_PRESETS.map((preset) => ({ id: newId("p"), ...preset }));
}

export function presetText(preset) {
  if (!preset.builtin) {
    return {
      title: preset.title ?? "",
      prompt: preset.prompt ?? "",
      category: preset.category ?? "",
    };
  }
  return {
    title: t(`preset.${preset.builtin}.title`),
    prompt: t(`preset.${preset.builtin}.prompt`),
    category: t(`preset.${preset.builtin}.category`),
  };
}

// Earlier versions saved the default presets as plain text. A saved copy that
// still matches one of those texts exactly was never edited, so it becomes the
// built-in again and follows the language; anything the user changed is kept.
const LEGACY_TEXT = {
  summarize: [
    "List the worksheets, then read the active sheet's used range and give me a tight summary — what the data is, columns, row count, anything notable. Don't change anything.",
    "列出所有工作表，然后读取当前工作表的已用区域，简要总结：数据是什么、有哪些列、多少行、有什么值得注意的地方。不要修改任何内容。",
  ],
  explainSelection: [
    "Read my current selection and explain what it contains — the columns, the values, and any pattern or total worth noting. Don't change anything.",
    "读取我当前的选区，说明其中的内容：各列含义、数值，以及值得注意的规律或合计。不要修改任何内容。",
  ],
  totalsRow: [
    "Add a labelled Total row beneath the data on the active sheet, using SUM formulas (not pre-computed numbers) for each numeric column. Read the range first; don't overwrite existing formulas.",
    "在当前工作表数据下方添加一行带标签的合计行，每个数值列用 SUM 公式（不要写死计算结果）。先读取区域，不要覆盖已有公式。",
  ],
  checkData: [
    "Scan the active sheet for data problems — blank cells in a filled column, inconsistent formatting/casing, likely typos, duplicates. List what you find in chat with cell addresses; don't change anything yet.",
    "检查当前工作表的数据问题：应填列中的空白单元格、格式或大小写不一致、疑似错别字、重复行。在对话里列出发现的问题和单元格地址，先不要修改。",
  ],
  findValue: ["Find every cell containing: ", "找出所有包含以下内容的单元格："],
  contextFiles: [
    "Use the context files I've added to this workspace to answer: ",
    "根据我添加到这个工作区的参考文件回答：",
  ],
};

const LEGACY_TITLES = {
  summarize: ["Summarize this sheet", "总结这张表"],
  explainSelection: ["Explain the selected range", "解释选中的区域"],
  totalsRow: ["Add a totals row", "添加合计行"],
  checkData: ["Check the data for problems", "检查数据问题"],
  findValue: ["Find a value", "查找某个值"],
  contextFiles: ["Answer using my context files", "根据参考文件回答"],
};

export function migratePresets(presets) {
  return presets.map((preset) => {
    if (preset.builtin) return preset;
    const builtin = Object.keys(LEGACY_TEXT).find(
      (id) => LEGACY_TEXT[id].includes(preset.prompt) && LEGACY_TITLES[id].includes(preset.title),
    );
    if (!builtin) return preset;
    const { title: _title, prompt: _prompt, category: _category, ...rest } = preset;
    return { ...rest, builtin };
  });
}
