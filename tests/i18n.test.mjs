import { test } from "node:test";
import assert from "node:assert/strict";
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { i18n, resolveLanguage, setLanguage, t } from "../taskpane/app/core/i18n.js";
import en from "../taskpane/app/core/locales/en.js";
import zh from "../taskpane/app/core/locales/zh.js";
import {
  decisionLabel,
  describeGroup,
  mutationSummary,
  recoveryOperationLabel,
  statusForTool,
  toolTitle,
} from "../taskpane/app/core/labels.js";
import { BUILTIN_PRESETS, migratePresets, presetText } from "../taskpane/app/core/presets.js";

const root = new URL("../taskpane/app/", import.meta.url);

function sources(dir) {
  return readdirSync(new URL(dir, root), { withFileTypes: true }).flatMap((entry) =>
    entry.isDirectory()
      ? sources(`${dir}${entry.name}/`)
      : /\.(js|tsx)$/.test(entry.name)
        ? [readFileSync(new URL(`${dir}${entry.name}`, root), "utf8")]
        : [],
  );
}

// Keys written literally in the pane's code: t("…"), translate("…"), stored
// { key: "…" } entries and the controller's event/notice/error/setAgent calls.
function literalKeys() {
  const code = [...sources("core/"), ...sources("view/"), ...sources("vendor/ai-elements/")].join(
    "\n",
  );
  const patterns = [
    /\bt\("([^"]+)"/g,
    /translate\("([^"]+)"/g,
    /translate\((\w+) \? "([^"]+)" : "([^"]+)"\)/g,
    /\bt\((\w+) \? "([^"]+)" : "([^"]+)"/g,
    /(?:key|messageKey): "([^"]+)"/g,
    /\b(?:event|notice|error)\("([a-z]+\.[A-Za-z.]+)"/g,
    /setAgent\("\w+", "([^"]+)"/g,
    /"((?:conn|agent|divider|notice|event|error|backups)\.[A-Za-z.]+)"/g,
  ];
  const keys = new Set();
  for (const pattern of patterns) {
    for (const match of code.matchAll(pattern)) {
      for (const key of match.slice(1))
        if ((key && !/^\w+$/.test(key)) || /^[A-Z]/.test(key ?? "")) keys.add(key);
    }
  }
  return keys;
}

function has(table, key) {
  return key in table || `${key}_one` in table || `${key}_other` in table;
}

test("every key used in the pane exists in Chinese and English", () => {
  const vendorEnglish = new Set(Object.keys(zh).filter((key) => /^[A-Z]/.test(key)));
  const missing = [];
  for (const key of literalKeys()) {
    if (!has(zh, key)) missing.push(`zh: ${key}`);
    // The vendored components' English text is its own key.
    if (!has(en, key) && !vendorEnglish.has(key)) missing.push(`en: ${key}`);
  }
  assert.deepEqual(missing, []);
});

test("the two tables have the same keys", () => {
  const base = (key) => key.replace(/_(one|other)$/, "");
  const enKeys = new Set(Object.keys(en).map(base));
  const zhKeys = new Set(
    Object.keys(zh)
      .filter((key) => !/^[A-Z]/.test(key))
      .map(base),
  );
  assert.deepEqual(
    [...zhKeys].filter((key) => !enKeys.has(key)),
    [],
  );
  assert.deepEqual(
    [...enKeys].filter((key) => !zhKeys.has(key)),
    [],
  );
});

// i18next returns the key itself when nothing is found.
function resolved(value) {
  return typeof value === "string" && !/^[a-z]+\.[\w.]+$/.test(value);
}

test("keys built at run time resolve in both languages", () => {
  const tools = [
    "excel_set_cell_range",
    "mcp__office__excel_copy_to",
    "excel_get_range_as_csv",
    "excel_bash",
    "Read",
    "WebSearch",
  ];
  for (const language of ["zh", "en"]) {
    setLanguage(language);
    for (const tool of tools) {
      assert.ok(resolved(statusForTool(tool)), `${language} status ${tool}`);
      assert.ok(resolved(toolTitle(tool)), `${language} title ${tool}`);
    }
    assert.ok(resolved(statusForTool("mcp__thepexcel__power_query")));
    for (const operation of ["excel_set_cell_range", "excel_modify_sheet_structure", "restore"]) {
      assert.ok(resolved(recoveryOperationLabel(operation)), `${language} op ${operation}`);
    }
    for (const decision of [
      "approve",
      "approve_turn",
      "reject",
      "timeout",
      "cancelled",
      "disabled",
      null,
    ]) {
      assert.ok(resolved(decisionLabel(decision)), `${language} decision ${decision}`);
    }
    for (const tool of ["excel_get_cell_ranges", "excel_copy_to", "excel_select_range"]) {
      assert.ok(resolved(describeGroup(tool, 3)), `${language} group ${tool}`);
    }
    for (const preset of BUILTIN_PRESETS) {
      const text = presetText(preset);
      for (const part of ["title", "prompt", "category"]) {
        assert.ok(resolved(text[part]), `${language} preset ${preset.builtin}.${part}`);
      }
    }
    for (const key of [
      "model.hint.haiku",
      "provider.tier.opus",
      "approvalMode.ask.hint",
      "receipt.commit.not_committed",
      "receipt.verify.read_back",
    ]) {
      assert.ok(resolved(t(key)), `${language} ${key}`);
    }
  }
});

test("switching language changes labels, counts and plurals", () => {
  setLanguage("zh");
  assert.equal(
    mutationSummary("excel_set_cell_range", { cells: [[{ value: 1 }, { formula: "=A1*2" }]] }),
    "1×2 个单元格 · 1 个公式 · 1 个值",
  );
  assert.equal(t("backups.cells", { count: 3 }), "3 个单元格");
  assert.equal(t("Input"), "输入");
  setLanguage("en");
  assert.equal(
    mutationSummary("excel_set_cell_range", { cells: [[{ value: 1 }, { formula: "=A1*2" }]] }),
    "1×2 cells · 1 formula · 1 value",
  );
  assert.equal(t("backups.cells", { count: 3 }), "3 cells");
  assert.equal(describeGroup("excel_get_cell_ranges", 1), "1 read");
  assert.equal(t("Input"), "Input");
  assert.equal(i18n.language, "en");
});

test("auto language follows Office's display language", () => {
  assert.equal(resolveLanguage("auto", "zh-CN"), "zh");
  assert.equal(resolveLanguage("auto", "zh-TW"), "zh");
  assert.equal(resolveLanguage("auto", "en-US"), "en");
  assert.equal(resolveLanguage("auto", "de-DE"), "en");
  assert.equal(resolveLanguage("auto", null), "en");
  assert.equal(resolveLanguage("en", "zh-CN"), "en");
});

test("untouched presets saved by older versions follow the language again", () => {
  const migrated = migratePresets([
    {
      id: "a",
      title: "Summarize this sheet",
      prompt: en["preset.summarize.prompt"],
      pinned: true,
      auto_send: true,
    },
    // Renamed by the user: kept as their own text.
    {
      id: "b",
      title: "My summary",
      prompt: en["preset.summarize.prompt"],
      pinned: false,
      auto_send: true,
    },
    {
      id: "c",
      title: "解释选中的区域",
      prompt: zh["preset.explainSelection.prompt"],
      pinned: true,
      auto_send: true,
    },
    { id: "d", title: "Mine", prompt: "Do my thing", pinned: false, auto_send: false },
  ]);
  assert.deepEqual(migrated[0], { id: "a", pinned: true, auto_send: true, builtin: "summarize" });
  assert.equal(migrated[1].builtin, undefined);
  assert.equal(migrated[2].builtin, "explainSelection");
  assert.equal(migrated[3].prompt, "Do my thing");
  setLanguage("zh");
  assert.equal(presetText(migrated[0]).title, "总结这张表");
  setLanguage("en");
  assert.equal(presetText(migrated[0]).title, "Summarize this sheet");
});
