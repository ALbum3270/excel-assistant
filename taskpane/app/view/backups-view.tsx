import { useMemo, useState } from "react";
import { RefreshCwIcon } from "lucide-react";
import { Button } from "../vendor/shadcn-ui/components/ui/button";
import { Input } from "../vendor/shadcn-ui/components/ui/input";
import { recoveryOperationLabel } from "../core/labels.js";
import { ConfirmButton, formatTime, usePane, usePaneState, useT } from "./common";

export function BackupsView() {
  const pane = usePane();
  const t = useT();
  const backups = usePaneState((state) => state.backups);
  const [query, setQuery] = useState("");

  const visible = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return backups.snapshots;
    return backups.snapshots.filter((snapshot: any) =>
      [
        recoveryOperationLabel(snapshot.operation),
        ...(snapshot.addresses || []),
        ...(snapshot.kinds || []),
      ]
        .join(" ")
        .toLowerCase()
        .includes(q),
    );
  }, [backups.snapshots, query]);

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
            {backups.snapshots.length ? t("backups.noMatch") : t("backups.empty")}
          </p>
        )}
        {visible.map((snapshot: any) => (
          <div className="space-y-1 rounded-md border p-2" key={snapshot.id}>
            <div className="flex items-center justify-between gap-2">
              <span className="font-medium text-sm">
                {recoveryOperationLabel(snapshot.operation)}
              </span>
              <span className="text-muted-foreground text-xs">
                {t("backups.cells", { count: Number(snapshot.changedCount || 0) })}
              </span>
            </div>
            <div className="break-all font-mono text-xs">
              {(snapshot.addresses || []).join(", ") || t("backups.structure")}
            </div>
            <div className="text-muted-foreground text-xs">{formatTime(snapshot.createdAt)}</div>
            <div className="flex gap-1">
              <Button
                disabled={backups.busy}
                onClick={() => pane.restoreBackup(snapshot.id)}
                size="sm"
                variant="outline"
              >
                {t("backups.restore")}
              </Button>
              <ConfirmButton
                confirmLabel={t("common.deleteConfirm")}
                disabled={backups.busy}
                onConfirm={() => pane.deleteBackup(snapshot.id)}
              >
                {t("common.delete")}
              </ConfirmButton>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}
