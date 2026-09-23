import { useMemo, useState } from "react";
import { RefreshCwIcon, Redo2Icon, Undo2Icon } from "lucide-react";
import { Button } from "../vendor/shadcn-ui/components/ui/button";
import { Input } from "../vendor/shadcn-ui/components/ui/input";
import { recoveryOperationLabel } from "../core/labels.js";
import { collapseRestoreChains } from "../core/restore-chains.js";
import { ConfirmButton, formatTime, usePane, usePaneState, useT } from "./common";

export function BackupsView() {
  const pane = usePane();
  const t = useT();
  const backups = usePaneState((state) => state.backups);
  const [query, setQuery] = useState("");

  // One row per change: the restore checkpoints a click leaves behind are
  // folded into the change they undo, so the row count never grows by clicking.
  const rows = useMemo(() => collapseRestoreChains(backups.snapshots), [backups.snapshots]);

  const visible = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return rows;
    return rows.filter((row: any) =>
      [recoveryOperationLabel(row.operation), ...(row.addresses || []), ...(row.kinds || [])]
        .join(" ")
        .toLowerCase()
        .includes(q),
    );
  }, [rows, query]);

  return (
    <div className="flex min-h-0 flex-1 flex-col gap-3 overflow-y-auto p-3">
      <div>
        <h2 className="font-medium text-sm">{t("backups.title")}</h2>
        <p className="text-muted-foreground text-xs">{t("backups.hint")}</p>
      </div>
      <div className="flex items-center gap-1">
        <Input
          className="h-8 text-xs"
          onChange={(e) => setQuery(e.target.value)}
          placeholder={t("backups.search")}
          value={query}
        />
        <Button
          disabled={backups.busy}
          onClick={() => pane.loadBackups()}
          size="icon-sm"
          title={t("backups.refresh")}
          variant="ghost"
        >
          <RefreshCwIcon className="size-4" />
        </Button>
        <ConfirmButton
          confirmLabel={t("backups.clearConfirm")}
          disabled={backups.busy || backups.snapshots.length === 0}
          onConfirm={() => pane.clearBackups()}
        >
          {t("backups.clear")}
        </ConfirmButton>
      </div>
      {backups.status && (
        <p className={`text-xs ${backups.error ? "text-destructive" : "text-muted-foreground"}`}>
          {backups.status}
        </p>
      )}
      <div className="space-y-2">
        {backups.loaded && visible.length === 0 && (
          <p className="text-muted-foreground text-xs">
            {rows.length ? t("backups.noMatch") : t("backups.empty")}
          </p>
        )}
        {visible.map((row: any) => {
          const Icon = row.undone ? Redo2Icon : Undo2Icon;
          return (
            <div className="space-y-1 rounded-md border p-2" key={row.id}>
              <div className="flex items-center justify-between gap-2">
                <span className="font-medium text-sm">{recoveryOperationLabel(row.operation)}</span>
                {row.undone ? (
                  <span className="rounded bg-muted px-1.5 py-0.5 text-muted-foreground text-xs">
                    {t("backups.undoneTag")}
                  </span>
                ) : (
                  <span className="text-muted-foreground text-xs">
                    {t("backups.cells", { count: Number(row.changedCount || 0) })}
                  </span>
                )}
              </div>
              <div className="break-all font-mono text-xs">
                {(row.addresses || []).join(", ") || t("backups.structure")}
              </div>
              <div className="text-muted-foreground text-xs">
                {formatTime(row.createdAt)}
                {row.restoreCount > 0 &&
                  ` · ${t("backups.restoredTimes", { count: row.restoreCount })}`}
              </div>
              <div className="flex gap-1">
                <Button
                  disabled={backups.busy}
                  onClick={() => pane.restoreBackup(row.tipId, { redo: row.undone })}
                  size="sm"
                  title={row.undone ? t("card.redoTooltip") : t("card.undoTooltip")}
                  variant="outline"
                >
                  <Icon className="size-3.5" />
                  {row.undone ? t("card.redo") : t("card.undo")}
                </Button>
                <ConfirmButton
                  confirmLabel={t("common.deleteConfirm")}
                  disabled={backups.busy}
                  onConfirm={() => pane.deleteBackup(row.chain)}
                >
                  {t("common.delete")}
                </ConfirmButton>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}
