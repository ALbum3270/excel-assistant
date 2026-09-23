import { useEffect } from "react";
import {
  AlertTriangleIcon,
  ArrowLeftIcon,
  ArchiveRestoreIcon,
  HistoryIcon,
  MoonIcon,
  SquarePenIcon,
  Settings2Icon,
  SunIcon,
  XIcon,
} from "lucide-react";
import { Alert, AlertDescription, AlertTitle } from "../vendor/shadcn-ui/components/ui/alert";
import { Button } from "../vendor/shadcn-ui/components/ui/button";
import { cn } from "../vendor/shadcn-ui/lib/utils";
import { BackupsView } from "./backups-view";
import { LogoMark } from "./brand";
import { ChatView } from "./chat";
import { usePane, usePaneState, useT } from "./common";
import { HistoryDialog } from "./history-dialog";
import { SettingsView } from "./settings-view";

function workbookName(url: string | null) {
  if (!url) return null;
  const last = url.split(/[\\/]/).filter(Boolean).pop() ?? "";
  try {
    return decodeURIComponent(last.split("?")[0]);
  } catch {
    return last;
  }
}

// Light or dark: follows Excel until the user picks one here, then stays.
function useDarkMode() {
  const theme = usePaneState((state) => state.settings.theme);
  const hostDark = usePaneState((state) => state.hostDark);
  const dark = theme === "auto" ? hostDark : theme === "dark";
  useEffect(() => {
    document.documentElement.classList.toggle("dark", dark);
  }, [dark]);
  return dark;
}

// Header buttons: small, quiet until hovered, filled when their page is open.
function HeaderButton({
  active,
  title,
  onClick,
  children,
}: {
  active?: boolean;
  title: string;
  onClick: () => void;
  children: React.ReactNode;
}) {
  return (
    <Button
      className={cn(
        "size-7 rounded-lg text-muted-foreground hover:text-foreground",
        active && "bg-accent text-foreground",
      )}
      onClick={onClick}
      size="icon-sm"
      title={title}
      variant="ghost"
    >
      {children}
    </Button>
  );
}

function ThemeToggle() {
  const pane = usePane();
  const t = useT();
  const dark = useDarkMode();
  return (
    <HeaderButton
      onClick={() => pane.setTheme(dark ? "light" : "dark")}
      title={dark ? t("theme.toLight") : t("theme.toDark")}
    >
      {dark ? <SunIcon className="size-4" /> : <MoonIcon className="size-4" />}
    </HeaderButton>
  );
}

function Header() {
  const pane = usePane();
  const t = useT();
  const connection = usePaneState((state) => state.connection);
  const workspace = usePaneState((state) => state.workspace);
  const view = usePaneState((state) => state.view);
  const folder = workspace.cwd?.split(/[\\/]/).filter(Boolean).pop();
  const toggle = (target: string) => pane.setView(view === target ? "chat" : target);
  const dot = { ok: "bg-brand", err: "bg-destructive", idle: "bg-amber-400 animate-pulse" }[
    connection.state as string
  ];

  return (
    <header className="flex items-center gap-2.5 border-border/60 border-b bg-background/85 px-3 py-2.5 backdrop-blur">
      <LogoMark className="size-7 shrink-0 drop-shadow-sm" />
      <button
        className="min-w-0 flex-1 text-left"
        onClick={() => pane.setView("settings")}
        title={workspace.cwd ?? ""}
        type="button"
      >
        <div className="truncate font-semibold text-[13px] tracking-tight">
          {workbookName(workspace.activeDocUrl) ?? t("app.name")}
        </div>
        <div className="flex items-center gap-1.5 truncate text-[11px] text-muted-foreground">
          <span
            className={cn("inline-block size-1.5 shrink-0 rounded-full", dot)}
            title={t(connection.key)}
          />
          {connection.state === "ok" ? (folder ?? t("workspace.none")) : t(connection.key)}
          {workspace.mismatch && <AlertTriangleIcon className="size-3 shrink-0 text-amber-500" />}
        </div>
      </button>
      <div className="flex items-center gap-0.5">
        <ThemeToggle />
        <HeaderButton onClick={() => pane.openHistory()} title={t("history.title")}>
          <HistoryIcon className="size-4" />
        </HeaderButton>
        <HeaderButton
          active={view === "backups"}
          onClick={() => toggle("backups")}
          title={t("backups.title")}
        >
          <ArchiveRestoreIcon className="size-4" />
        </HeaderButton>
        <HeaderButton
          active={view === "settings"}
          onClick={() => toggle("settings")}
          title={t("settings.title")}
        >
          <Settings2Icon className="size-4" />
        </HeaderButton>
        <HeaderButton
          onClick={() => {
            pane.setView("chat");
            pane.newChat();
          }}
          title={t("app.newChat")}
        >
          <SquarePenIcon className="size-4" />
        </HeaderButton>
      </div>
    </header>
  );
}

// Settings and restore points are pages over the chat; this bar leads back.
function BackBar({ title }: { title: string }) {
  const pane = usePane();
  const t = useT();
  return (
    <div className="flex items-center gap-1 px-2 pt-2">
      <Button
        className="h-7 gap-1 rounded-lg px-2 text-muted-foreground text-xs hover:text-foreground"
        onClick={() => pane.setView("chat")}
        size="sm"
        variant="ghost"
      >
        <ArrowLeftIcon className="size-3.5" />
        {t("app.backToChat")}
      </Button>
      <span className="ml-auto pr-2 font-semibold text-sm tracking-tight">{title}</span>
    </div>
  );
}

// Shown when the daemon cannot authenticate with the model provider. Recovery
// happens outside the pane, so it stays until dismissed.
function AuthBanner() {
  const pane = usePane();
  const t = useT();
  const authError = usePaneState((state) => state.authError);
  if (!authError) return null;
  return (
    <Alert className="m-3 mb-0 w-auto" variant="destructive">
      <AlertTriangleIcon />
      <AlertTitle className="flex items-center justify-between">
        {t("auth.title")}
        <button
          className="rounded p-0.5 hover:bg-destructive/10"
          onClick={() => pane.dismissAuthError()}
          title={t("common.close")}
          type="button"
        >
          <XIcon className="size-3.5" />
        </button>
      </AlertTitle>
      <AlertDescription className="space-y-1 text-xs">
        <p>{t("auth.body")}</p>
        <ul className="list-disc pl-4">
          <li>{t("auth.env")}</li>
          <li>{t("auth.claude")}</li>
        </ul>
        <p>{t("auth.restart")}</p>
        <details>
          <summary className="cursor-pointer">{t("auth.raw")}</summary>
          <pre className="whitespace-pre-wrap break-all">{authError}</pre>
        </details>
      </AlertDescription>
    </Alert>
  );
}

export function App() {
  const pane = usePane();
  // Subscribing here also re-renders the whole pane, vendored components
  // included, when the language changes.
  const t = useT();
  const view = usePaneState((state) => state.view);

  // Escape stops the running turn, as the working indicator says. It only acts
  // while the agent works and no dialog is open, so Escape keeps its usual
  // meaning (closing a dialog or menu) the rest of the time.
  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      if (event.key !== "Escape") return;
      if (document.querySelector('[role="dialog"], [role="menu"]')) return;
      const state = pane.store.getState();
      if (state.agent.state === "working") {
        event.preventDefault();
        pane.stop();
      } else if (state.view !== "chat") {
        // On a settings or restore-points page, Escape goes back to the chat.
        event.preventDefault();
        pane.setView("chat");
      }
    };
    document.addEventListener("keydown", onKey);
    return () => document.removeEventListener("keydown", onKey);
  }, [pane]);

  return (
    <div className="flex h-full flex-col bg-background">
      <Header />
      <AuthBanner />
      {view === "chat" && <ChatView />}
      {view !== "chat" && (
        <BackBar title={t(view === "backups" ? "backups.title" : "settings.title")} />
      )}
      {view === "backups" && <BackupsView />}
      {view === "settings" && <SettingsView />}
      <HistoryDialog />
    </div>
  );
}
