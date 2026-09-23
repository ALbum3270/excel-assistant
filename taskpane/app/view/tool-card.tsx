// One card per tool call. The layout follows pi-for-excel's tool cards
// (src/ui/tool-renderers.ts): a one-line header with the action, the range and
// how many cells changed, and a Before/After table as the body. Undo and redo
// sit on the card of the change itself, as Cline and Cursor do for edits.
import { ChevronDownIcon, LoaderIcon, Redo2Icon, Undo2Icon, type LucideIcon } from "lucide-react";
import {
  Collapsible,
  CollapsibleContent,
  CollapsibleTrigger,
} from "../vendor/shadcn-ui/components/ui/collapsible";
import { Button } from "../vendor/shadcn-ui/components/ui/button";
import { cn } from "../vendor/shadcn-ui/lib/utils";
import { csvPreview } from "../core/csv-preview.js";
import { describeArgs, mutationSummary, toolTitle } from "../core/labels.js";
import { CellLink, CellLinks, Rows, usePane, usePaneState, useT } from "./common";
import { ToolIcon } from "./brand";

function target(args: any, result?: any) {
  // A restore names its ranges only in the result.
  if (result?.addresses?.length) return result.addresses.join(", ");
  if (!args || typeof args !== "object") return "";
  return args.range ?? args.address ?? args.destinationRange ?? args.table ?? args.sheetName ?? "";
}

function undoPoint(item: any): string | null {
  if (item.undo) return item.undo.snapshotId;
  return item.result?.recovery?.snapshotIds?.[0] ?? item.error?.recovery?.snapshotIds?.[0] ?? null;
}

// ---- Header -----------------------------------------------------------------

function StatusChip({ item }: { item: any }) {
  const t = useT();
  const chip = (label: string, className: string, Icon?: LucideIcon) => (
    <span
      className={cn(
        "inline-flex shrink-0 items-center gap-1 rounded-full px-2 py-0.5 text-xs",
        className,
      )}
    >
      {Icon && <Icon className="size-3 animate-spin" />}
      {label}
    </span>
  );
  if (item.state === "running")
    return chip(t("card.running"), "bg-muted text-muted-foreground", LoaderIcon);
  if (item.state === "error") return chip(t("card.failed"), "bg-destructive/10 text-destructive");
  if (item.state === "unknown")
    return chip(t("card.needsCheck"), "bg-amber-500/15 text-amber-700 dark:text-amber-400");
  if (item.undo?.undone) return chip(t("card.undone"), "bg-muted text-muted-foreground");
  return null;
}

function UndoButton({ item }: { item: any }) {
  const pane = usePane();
  const t = useT();
  const busyTurn = usePaneState((state) => state.turnInFlight || state.submitPending);
  const recoveryBusy = usePaneState((state) => state.backups.busy);
  if (!undoPoint(item) || item.state === "running") return null;
  const undone = Boolean(item.undo?.undone);
  const Icon = item.undo?.busy ? LoaderIcon : undone ? Redo2Icon : Undo2Icon;
  return (
    <Button
      className="h-7 shrink-0 gap-1 px-2 text-xs"
      disabled={item.undo?.busy || busyTurn || recoveryBusy}
      onClick={() => pane.toggleUndo(item.id)}
      size="sm"
      title={
        busyTurn
          ? t("error.restoreWhileBusy")
          : undone
            ? t("card.redoTooltip")
            : t("card.undoTooltip")
      }
      variant="ghost"
    >
      <Icon className={cn("size-3.5", item.undo?.busy && "animate-spin")} />
      {undone ? t("card.redo") : t("card.undo")}
    </Button>
  );
}

function Header({ item }: { item: any }) {
  const t = useT();
  const where = target(item.args, item.result);
  const changed = item.diff?.changed;
  // Only the count of changed cells earns header space; the range must not be
  // squeezed out by a summary that mostly repeats the title.
  // Once undone, the "Undone" tag says what matters; the count would crowd out the range.
  const meta =
    item.state === "success" && item.write && changed && !item.undo?.undone
      ? t("card.changes", { count: changed })
      : null;
  return (
    <div className="flex items-center gap-1 pr-1">
      <CollapsibleTrigger className="group flex min-w-0 flex-1 items-center gap-2.5 py-2 pr-2 pl-2.5 text-left">
        <ToolIcon name={item.local} />
        <span className="min-w-0 truncate text-[13px]">
          <span className="font-medium">{toolTitle(item.local)}</span>
          {where && (
            <span className="ml-1.5 font-mono text-[11.5px] text-muted-foreground">{where}</span>
          )}
        </span>
        {meta && <span className="shrink-0 text-muted-foreground text-xs">· {meta}</span>}
        <StatusChip item={item} />
        <ChevronDownIcon className="ml-auto size-4 shrink-0 text-muted-foreground transition-transform group-data-[state=open]:rotate-180" />
      </CollapsibleTrigger>
      <UndoButton item={item} />
    </div>
  );
}

// ---- Body -------------------------------------------------------------------

function DiffValue({ value, formula }: { value?: string; formula?: string | null }) {
  const t = useT();
  return (
    <div>
      <div className={cn(!value && "text-muted-foreground")}>{value || t("card.empty")}</div>
      {formula && <div className="font-mono text-[11px] text-muted-foreground">ƒ {formula}</div>}
    </div>
  );
}

// The Before/After table of pi-for-excel's renderWorkbookCellDiff.
function ChangeTable({ diff }: { diff: any }) {
  const t = useT();
  return (
    <div className="space-y-1">
      <table className="w-full border-collapse text-xs">
        <thead>
          <tr className="text-left text-muted-foreground">
            <th className="w-20 py-1 pr-2 font-normal">{t("card.cell")}</th>
            <th className="py-1 pr-2 font-normal">{t("card.before")}</th>
            <th className="py-1 font-normal">{t("card.after")}</th>
          </tr>
        </thead>
        <tbody>
          {diff.changes.map((change: any) => (
            <tr className="border-t align-top" key={change.cell}>
              <td className="py-1 pr-2">
                <CellLink address={change.cell} label={change.cell.split("!").pop()} />
              </td>
              <td className="py-1 pr-2">
                {"beforeValue" in change ? (
                  <DiffValue formula={change.beforeFormula} value={change.beforeValue} />
                ) : (
                  <DiffValue value={change.before === "(blank)" ? "" : change.before} />
                )}
              </td>
              <td className="py-1">
                {"afterValue" in change ? (
                  <DiffValue formula={change.afterFormula} value={change.afterValue} />
                ) : (
                  <DiffValue value={change.after === "(blank)" ? "" : change.after} />
                )}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      {diff.changed > diff.changes.length && (
        <p className="text-muted-foreground text-xs">
          {t("card.shownFirst", { count: diff.changes.length, total: diff.changed })}
        </p>
      )}
    </div>
  );
}

function CsvTable({ item }: { item: any }) {
  const t = useT();
  const preview = csvPreview(item.result.csv, item.args?.range);
  const notes = [];
  if (preview.hiddenRows > 0) notes.push(t("csv.moreRows", { count: preview.hiddenRows }));
  if (preview.hiddenColumns > 0) notes.push(t("csv.moreColumns", { count: preview.hiddenColumns }));
  if (item.result.hasMore) notes.push(t("csv.continues", { range: item.result.nextRange }));
  return (
    <div className="space-y-1">
      <div className="overflow-x-auto">
        <table className="w-full border-collapse font-mono text-[11px]">
          <thead>
            <tr>
              <th className="border bg-muted px-1" />
              {preview.columns.map((column) => (
                <th className="border bg-muted px-1 font-normal text-muted-foreground" key={column}>
                  {column}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {preview.rows.map((row) => (
              <tr key={row.number}>
                <td className="border bg-muted px-1 text-right text-muted-foreground">
                  {row.number}
                </td>
                {row.cells.map((cell, index) => (
                  <td className="max-w-32 truncate border px-1" key={index} title={cell}>
                    {cell}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {notes.length > 0 && (
        <p className="text-muted-foreground text-xs">
          {t("csv.notShown", { notes: notes.join(t("csv.separator")) })}
        </p>
      )}
    </div>
  );
}

const COMMIT_STATUSES = new Set(["committed", "not_committed", "unknown"]);
const VERIFICATIONS = new Set(["read_back", "commit_acknowledged"]);

// Arguments and receipt fields are for checking what happened, not for reading
// every time, so they stay folded away.
function TechnicalDetails({ item }: { item: any }) {
  const t = useT();
  const receipt = item.result ?? {};
  const rows: [string, React.ReactNode][] = [];
  const commit = receipt.commitStatus;
  if (commit) {
    rows.push([
      t("receipt.status"),
      COMMIT_STATUSES.has(commit) ? t(`receipt.commit.${commit}`) : commit,
    ]);
  }
  const verification = receipt.verification?.status;
  if (verification) {
    rows.push([
      t("receipt.verification"),
      VERIFICATIONS.has(verification) ? t(`receipt.verify.${verification}`) : verification,
    ]);
  }
  if (Number.isInteger(receipt.workbookRevision)) {
    rows.push([t("receipt.revision"), String(receipt.workbookRevision)]);
  }
  return (
    <details className="group/details text-xs">
      <summary className="cursor-pointer select-none text-muted-foreground hover:text-foreground">
        {t("card.details")}
      </summary>
      <div className="mt-2 space-y-2">
        <pre className="whitespace-pre-wrap break-words rounded-md bg-muted/50 p-2 font-mono text-[11px]">
          {describeArgs(item.local, item.args)}
        </pre>
        {rows.length > 0 && (
          <div className="rounded-md bg-muted/50">
            <Rows rows={rows} />
          </div>
        )}
      </div>
    </details>
  );
}

function Unlock({ item }: { item: any }) {
  const pane = usePane();
  const t = useT();
  const unlock = item.unlock;
  if (unlock?.revision !== undefined) {
    return <p className="text-xs">{t("receipt.unlocked", { revision: unlock.revision })}</p>;
  }
  return (
    <div className="space-y-1">
      <Button
        disabled={unlock?.busy}
        onClick={() => pane.acknowledgeWorkbook(item.id)}
        size="sm"
        variant="outline"
      >
        {t("receipt.unlock")}
      </Button>
      {unlock?.error && <p className="text-destructive text-xs">{unlock.error}</p>}
    </div>
  );
}

function Body({ item }: { item: any }) {
  const t = useT();
  const receipt = item.result ?? {};
  const failed = item.state === "error" || item.state === "unknown";
  const formulaErrors = receipt.formulaErrorCount || receipt.formulaErrors?.length;
  return (
    <div className="space-y-3 border-border/60 border-t px-3 py-3">
      {item.undo?.error && <p className="text-destructive text-xs">{item.undo.error}</p>}
      {failed && (
        <div className="space-y-2">
          <p className="whitespace-pre-wrap rounded-md bg-destructive/10 p-2 text-destructive text-xs">
            {item.error?.messageKey
              ? t(item.error.messageKey)
              : (item.error?.message ?? t("card.failed"))}
          </p>
          {(item.state === "unknown" || item.error?.code === "WORKBOOK_UNCERTAIN") && (
            <Unlock item={item} />
          )}
        </div>
      )}
      {formulaErrors > 0 && (
        <p className="rounded-md bg-amber-500/15 p-2 text-amber-700 text-xs dark:text-amber-400">
          {t("card.formulaErrors", { count: formulaErrors })}
        </p>
      )}
      {item.state === "success" && item.write && item.diff?.changes?.length > 0 && (
        <ChangeTable diff={item.diff} />
      )}
      {item.state === "success" &&
        item.write &&
        !item.diff &&
        receipt.affectedTargets?.length > 0 && (
          // Replayed from an earlier session: the cell-by-cell diff is not kept.
          <Rows
            rows={[
              [t("receipt.change"), mutationSummary(item.local, item.args ?? {}) ?? ""],
              [t("receipt.range"), <CellLinks addresses={receipt.affectedTargets} />],
            ]}
          />
        )}
      {item.state === "success" && item.write && receipt.recovery?.status === "not_available" && (
        <p className="text-muted-foreground text-xs">
          {receipt.recovery.reason
            ? t("card.noUndoReason", { reason: receipt.recovery.reason })
            : t("card.noUndo")}
        </p>
      )}
      {item.state === "success" &&
        item.local === "excel_get_range_as_csv" &&
        typeof receipt.csv === "string" &&
        receipt.csv && <CsvTable item={item} />}
      <TechnicalDetails item={item} />
    </div>
  );
}

export function ToolCard({ item }: { item: any }) {
  // Tools the SDK ran itself (files, web) are only announced; no result comes back here.
  if (item.state === "announced") {
    const where = target(item.args);
    return (
      <div
        className="flex items-center gap-2.5 rounded-xl border border-border/60 bg-card/50 py-1.5 pr-3 pl-2 text-muted-foreground text-xs"
        title={item.name}
      >
        <ToolIcon className="size-5" name={item.local} />
        <span className="truncate">
          {toolTitle(item.local)}
          {where && ` · ${where}`}
        </span>
      </div>
    );
  }

  return (
    <Collapsible
      className={cn(
        "rounded-xl border border-border/70 bg-card shadow-[var(--shadow-card)] transition-colors",
        item.state === "error" && "border-destructive/30",
        item.state === "unknown" && "border-amber-500/40",
      )}
      defaultOpen={item.state === "error" || item.state === "unknown"}
    >
      <Header item={item} />
      <CollapsibleContent>
        <Body item={item} />
      </CollapsibleContent>
    </Collapsible>
  );
}
