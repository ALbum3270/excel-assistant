import {
  createContext,
  useContext,
  useRef,
  useState,
  useSyncExternalStore,
  type ReactNode,
} from "react";
import { cn } from "../vendor/shadcn-ui/lib/utils";
import { Button } from "../vendor/shadcn-ui/components/ui/button";
import type { createController } from "../core/controller.js";
import { i18n } from "../core/i18n.js";
import { useTranslation } from "react-i18next";

// Re-renders the calling component when the language changes.
export function useT() {
  return useTranslation().t;
}

const localeTag = () => (i18n.language === "zh" ? "zh-CN" : "en-US");

export type Pane = ReturnType<typeof createController>;
export const PaneContext = createContext<Pane | null>(null);

export function usePane(): Pane {
  const pane = useContext(PaneContext);
  if (!pane) throw new Error("PaneContext is missing");
  return pane;
}

// Select a slice of pane state; the slice must be a stored value (not a new
// object per call) so React can tell when it changed.
export function usePaneState<T>(select: (state: any) => T): T {
  const pane = usePane();
  return useSyncExternalStore(pane.store.subscribe, () => select(pane.store.getState()));
}

// Addresses the pane already knows are addresses become links that select the
// range in Excel (pi-for-excel's cell-link.ts). Prose is never scanned for them.
const CELL_NAV_DEBOUNCE_MS = 300;
let lastNavigation = 0;

export function CellLink({ address, label }: { address: string; label?: string }) {
  const pane = usePane();
  const t = useT();
  const [failure, setFailure] = useState<string | null>(null);
  return (
    <button
      className={cn(
        "rounded bg-brand/10 px-1 font-mono text-[0.85em] text-brand hover:bg-brand/20",
        failure && "bg-destructive/10 text-destructive line-through",
      )}
      onClick={(event) => {
        event.preventDefault();
        event.stopPropagation();
        const now = Date.now();
        if (now - lastNavigation < CELL_NAV_DEBOUNCE_MS) return;
        lastNavigation = now;
        pane.navigate(address).then(
          () => setFailure(null),
          (error: any) => setFailure(error?.message ?? String(error)),
        );
      }}
      title={failure ?? t("cell.select", { address })}
      type="button"
    >
      {label ?? address}
    </button>
  );
}

export function CellLinks({ addresses }: { addresses: string[] }) {
  return (
    <span className="inline-flex flex-wrap gap-1">
      {addresses.map((address) => (
        <CellLink address={address} key={address} />
      ))}
    </span>
  );
}

export function Rows({ rows }: { rows: [string, ReactNode][] }) {
  return (
    <dl className="grid grid-cols-[auto_1fr] gap-x-3 gap-y-1 p-3 text-xs">
      {rows.map(([label, value]) => (
        <div className="contents" key={label}>
          <dt className="text-muted-foreground">{label}</dt>
          <dd className="min-w-0 break-words">{value}</dd>
        </div>
      ))}
    </dl>
  );
}

// Destructive actions ask by changing their own label; Office task panes block
// window.confirm on some builds.
export function ConfirmButton({
  children,
  confirmLabel,
  onConfirm,
  disabled,
  className,
  size = "sm",
}: {
  children: ReactNode;
  confirmLabel: string;
  onConfirm: () => void;
  disabled?: boolean;
  className?: string;
  size?: "sm" | "default";
}) {
  const [armed, setArmed] = useState(false);
  const timer = useRef<ReturnType<typeof setTimeout>>();
  return (
    <Button
      className={className}
      disabled={disabled}
      onClick={() => {
        if (!armed) {
          setArmed(true);
          clearTimeout(timer.current);
          timer.current = setTimeout(() => setArmed(false), 4000);
          return;
        }
        setArmed(false);
        onConfirm();
      }}
      size={size}
      variant={armed ? "destructive" : "ghost"}
    >
      {armed ? confirmLabel : children}
    </Button>
  );
}

export function formatTokens(value: number) {
  return new Intl.NumberFormat(localeTag(), {
    notation: "compact",
    maximumFractionDigits: 1,
  }).format(value);
}

export function formatTime(value: string | number) {
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? "" : date.toLocaleString(localeTag(), { hour12: false });
}

export function Section({
  title,
  description,
  children,
  action,
}: {
  title: string;
  description?: ReactNode;
  children: ReactNode;
  action?: ReactNode;
}) {
  return (
    <section className="space-y-3 rounded-xl border border-border/70 bg-card p-3.5 shadow-[var(--shadow-card)]">
      <div className="flex items-start justify-between gap-2">
        <div className="space-y-1">
          <h3 className="font-semibold text-[13px] tracking-tight">{title}</h3>
          {description && (
            <p className="text-muted-foreground text-xs leading-relaxed">{description}</p>
          )}
        </div>
        {action}
      </div>
      {children}
    </section>
  );
}
