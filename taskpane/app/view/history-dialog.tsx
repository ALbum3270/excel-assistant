import { useRef, useState } from "react";
import { Badge } from "../vendor/shadcn-ui/components/ui/badge";
import { Button } from "../vendor/shadcn-ui/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from "../vendor/shadcn-ui/components/ui/dialog";
import { ConfirmButton, formatTime, usePane, usePaneState, useT } from "./common";

function download(archive: unknown, title: string) {
  const safeTitle = (title || "conversation").replace(/[\\/:*?"<>|]+/g, "-").slice(0, 80);
  // Compact, so the file is exactly the size the daemon budgeted for and can
  // always be imported back; indentation would add bytes it did not count.
  const blob = new Blob([JSON.stringify(archive)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `${safeTitle || "conversation"}.excel-assistant.json`;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function SessionRow({
  session,
  active,
  onError,
}: {
  session: any;
  active: boolean;
  onError: (m: string) => void;
}) {
  const pane = usePane();
  const t = useT();
  const [busy, setBusy] = useState(false);
  const run = async (action: () => Promise<unknown>) => {
    setBusy(true);
    try {
      await action();
    } catch (error: any) {
      onError(error?.message ?? String(error));
    } finally {
      setBusy(false);
    }
  };
  const meta = session.archived
    ? t("history.archived")
    : session.resume_compatible
      ? ""
      : t("history.viewOnly");
  return (
    <div
      className={`space-y-1 rounded-md border p-2 ${active ? "border-primary/50 bg-primary/5" : ""}`}
    >
      <div className="flex items-center gap-2">
        <div className="min-w-0 flex-1 truncate font-medium text-sm" title={session.title}>
          {session.title || t("history.untitled")}
        </div>
        {active && <Badge variant="secondary">{t("history.current")}</Badge>}
      </div>
      <div className="text-muted-foreground text-xs">
        {formatTime(session.last_used)}
        {meta && ` · ${meta}`}
      </div>
      <div className="flex gap-1">
        <Button
          disabled={active || busy}
          onClick={() => run(() => pane.activateSession(session.session_id))}
          size="sm"
          variant="outline"
        >
          {session.resume_compatible ? t("history.continue") : t("history.view")}
        </Button>
        <Button
          disabled={busy}
          onClick={() =>
            run(async () => download(await pane.exportSession(session.session_id), session.title))
          }
          size="sm"
          variant="ghost"
        >
          {t("history.export")}
        </Button>
        <ConfirmButton
          confirmLabel={t("common.deleteConfirm")}
          disabled={busy}
          onConfirm={() => run(() => pane.deleteSession(session.session_id))}
        >
          {t("common.delete")}
        </ConfirmButton>
      </div>
    </div>
  );
}

export function HistoryDialog() {
  const pane = usePane();
  const t = useT();
  const history = usePaneState((state) => state.history);
  const fileInput = useRef<HTMLInputElement>(null);
  const [error, setError] = useState<string | null>(null);
  const [importing, setImporting] = useState(false);

  return (
    <Dialog
      onOpenChange={(open) => (open ? pane.openHistory() : pane.closeHistory())}
      open={history.open}
    >
      <DialogContent className="max-h-[85vh] max-w-[calc(100%-1.5rem)] overflow-y-auto p-4">
        <DialogHeader>
          <DialogTitle>{t("history.title")}</DialogTitle>
          <DialogDescription>{t("history.hint")}</DialogDescription>
        </DialogHeader>
        <div className="flex items-center gap-2">
          <Button
            disabled={importing}
            onClick={() => fileInput.current?.click()}
            size="sm"
            variant="outline"
          >
            {t("history.import")}
          </Button>
          <input
            accept="application/json,.json"
            hidden
            onChange={async (event) => {
              const file = event.target.files?.[0];
              if (!file) return;
              setImporting(true);
              setError(null);
              try {
                await pane.importSession(file);
              } catch (failure: any) {
                setError(failure?.message ?? String(failure));
              } finally {
                setImporting(false);
                event.target.value = "";
              }
            }}
            ref={fileInput}
            type="file"
          />
        </div>
        {(error || history.error) && (
          <p className="text-destructive text-xs">{error || history.error}</p>
        )}
        <div className="space-y-2">
          {history.loading && history.sessions.length === 0 && (
            <p className="text-muted-foreground text-xs">{t("common.loading")}</p>
          )}
          {!history.loading && history.sessions.length === 0 && !history.error && (
            <p className="text-muted-foreground text-xs">{t("history.empty")}</p>
          )}
          {history.sessions.map((session: any) => (
            <SessionRow
              active={session.session_id === history.activeId}
              key={session.session_id}
              onError={setError}
              session={session}
            />
          ))}
        </div>
      </DialogContent>
    </Dialog>
  );
}
