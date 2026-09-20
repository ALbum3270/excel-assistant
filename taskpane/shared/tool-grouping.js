// Collapse consecutive calls to the same tool into one expandable block.
//
// Ported from pi-for-excel's src/ui/tool-grouping.ts (MIT), which runs on plain
// DOM: a MutationObserver re-groups after each render, runs of three or more
// start collapsed, and a WeakMap keeps whatever the user expanded or collapsed
// by hand across re-groupings. Adapted to this panel's markup: cards are
// `.msg.tool` with `data-tool-name` and `data-tool-state`, and only completed
// cards group, so a running call always stays visible.

const COLLAPSE_THRESHOLD = 3;

const GROUP_LABELS = {
  excel_set_cell_range: "edits",
  excel_fill_formula: "formula fills",
  excel_get_cell_ranges: "reads",
  excel_get_range_as_csv: "reads",
  excel_search_data: "searches",
  excel_set_format: "formatting changes",
  excel_clear_cell_range: "clears",
  excel_copy_to: "copies",
  excel_modify_sheet_structure: "row/column changes",
  excel_modify_workbook_structure: "sheet changes",
  excel_verify_task: "result checks",
  excel_bash: "calculations",
};

function describeGroup(toolName, count) {
  return `${count} ${GROUP_LABELS[toolName] ?? `${toolName} calls`}`;
}

function buildHeader(toolName, count, collapsed) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "tool-group-header";
  button.setAttribute("aria-expanded", collapsed ? "false" : "true");
  button.textContent = describeGroup(toolName, count);
  button.addEventListener("click", () => {
    const wrapper = button.parentElement;
    if (!wrapper) return;
    const isCollapsed = wrapper.classList.toggle("tool-group-collapsed");
    button.setAttribute("aria-expanded", isCollapsed ? "false" : "true");
    wrapper.dataset.userToggled = "1";
  });
  return button;
}

function consecutiveSiblings(a, b) {
  for (let node = a.nextSibling; node; node = node.nextSibling) {
    if (node === b) return true;
    if (node.nodeType === Node.ELEMENT_NODE) return false;
  }
  return false;
}

export function initToolGrouping(root) {
  let frame = 0;
  const userToggleState = new WeakMap();

  function unwrapAll() {
    for (const wrapper of root.querySelectorAll(".tool-group")) {
      const parent = wrapper.parentNode;
      if (!parent) continue;
      const leader = wrapper.querySelector(".msg.tool");
      if (leader && wrapper.dataset.userToggled) {
        userToggleState.set(leader, wrapper.classList.contains("tool-group-collapsed"));
      }
      wrapper.querySelector(".tool-group-header")?.remove();
      while (wrapper.firstChild) parent.insertBefore(wrapper.firstChild, wrapper);
      parent.removeChild(wrapper);
    }
  }

  function applyGrouping() {
    observer.disconnect();
    unwrapAll();

    const runs = [];
    let current = [];
    let currentName = "";
    for (const card of root.querySelectorAll(".msg.tool")) {
      card.classList.remove("tool-group-member");
      const name = card.dataset.toolName ?? "";
      const grouped = name && card.dataset.toolState === "success";
      const continues =
        grouped && current.length > 0 && currentName === name && consecutiveSiblings(current.at(-1), card);
      if (continues) {
        current.push(card);
        continue;
      }
      if (current.length >= 2) runs.push({ cards: current, name: currentName });
      current = grouped ? [card] : [];
      currentName = grouped ? name : "";
    }
    if (current.length >= 2) runs.push({ cards: current, name: currentName });

    for (const run of runs) {
      const leader = run.cards[0];
      const collapsed = userToggleState.get(leader) ?? run.cards.length >= COLLAPSE_THRESHOLD;
      const wrapper = document.createElement("div");
      wrapper.className = `tool-group${collapsed ? " tool-group-collapsed" : ""}`;
      if (userToggleState.has(leader)) wrapper.dataset.userToggled = "1";
      wrapper.append(buildHeader(run.name, run.cards.length, collapsed));
      leader.parentNode?.insertBefore(wrapper, leader);
      for (const card of run.cards) wrapper.append(card);
      for (const card of run.cards.slice(1)) card.classList.add("tool-group-member");
    }

    observer.observe(root, { childList: true, subtree: true });
  }

  const observer = new MutationObserver(() => {
    if (frame) return;
    frame = requestAnimationFrame(() => {
      frame = 0;
      applyGrouping();
    });
  });
  observer.observe(root, { childList: true, subtree: true });
  applyGrouping();

  return () => {
    observer.disconnect();
    if (frame) cancelAnimationFrame(frame);
    unwrapAll();
  };
}
