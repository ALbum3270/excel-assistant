/* global Office */
// Task pane entry: wire the controller to the daemon bridge, Excel and the
// recovery store, then render the view.
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { I18nextProvider } from "react-i18next";
import { createWorkbookCoordinator } from "../shared/vendor/pi-context.js";
import { prepareMutationRecovery, takeMutationDiff } from "../shared/recovery.js";
import { createBridge } from "./core/bridge.js";
import { createController } from "./core/controller.js";
import { i18n } from "./core/i18n.js";
import {
  changeTracker,
  executeTool,
  getWorkbookMetadata,
  navigateToRange,
  readSelection,
  watchSelection,
} from "./core/excel-tools.js";
import { createOfficeRunner, describeOfficeToolError } from "./core/office-runner.js";
import { App } from "./view/app";
import { PaneContext } from "./view/common";

// Excel's own theme; without one (older builds), the OS preference. The pane
// follows it until the user picks light or dark in the header.
function hostPrefersDark() {
  let dark = globalThis.matchMedia?.("(prefers-color-scheme: dark)").matches ?? false;
  try {
    const hex = /^#?([0-9a-f]{6})$/i.exec(
      Office.context?.officeTheme?.bodyBackgroundColor ?? "",
    )?.[1];
    if (hex) {
      const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
      dark = (0.299 * r + 0.587 * g + 0.114 * b) / 255 < 0.5;
    }
  } catch {
    /* keep the OS preference */
  }
  return dark;
}

function render(content: React.ReactNode) {
  createRoot(document.getElementById("root")!).render(<StrictMode>{content}</StrictMode>);
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    render(
      <p className="p-4 text-sm">
        This add-in only runs in Excel ({String(info.host)}). 这个加载项只用于 Excel。
      </p>,
    );
    return;
  }
  let activeDocUrl: string | null = null;
  try {
    activeDocUrl = Office.context.document.url || null;
  } catch {
    /* unsaved workbook */
  }

  const coordinator = createWorkbookCoordinator();
  const pane = createController({
    makeBridge: (handlers) => createBridge(handlers),
    makeRunner: ({ send, onSettled }) =>
      createOfficeRunner({
        send,
        onSettled,
        executeTool,
        prepareMutationRecovery,
        takeMutationDiff,
        describeError: (error, args) =>
          describeOfficeToolError(error, args, { getWorkbookMetadata }),
        runWorkbookWrite: (opId, toolName, execute) =>
          coordinator.runWrite(
            {
              workbookId: activeDocUrl || "workbook:unknown",
              sessionId: "taskpane",
              opId,
              toolName,
            },
            execute,
          ),
      }),
    excel: { readSelection: () => readSelection(), watchSelection, navigateToRange },
    changeTracker,
  });

  let hostLanguage: string | null = null;
  try {
    hostLanguage = Office.context.displayLanguage || null;
  } catch {
    /* older host: English */
  }
  // Start first so the first frame is already in the right language and theme.
  pane.start({ activeDocUrl, hostDark: hostPrefersDark(), hostLanguage });
  render(
    <I18nextProvider i18n={i18n}>
      <PaneContext.Provider value={pane}>
        <App />
      </PaneContext.Provider>
    </I18nextProvider>,
  );
});
