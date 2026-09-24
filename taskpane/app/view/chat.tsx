import { useEffect, useMemo, useState } from "react";
import {
  AlertTriangleIcon,
  ChevronDownIcon,
  ClipboardListIcon,
  InfoIcon,
  LibraryIcon,
  ScanSearchIcon,
  SearchIcon,
  ShieldCheckIcon,
  SigmaIcon,
  SparklesIcon,
  TableIcon,
  Undo2Icon,
  XIcon,
  type LucideIcon,
} from "lucide-react";
import { motion } from "motion/react";
import {
  Confirmation,
  ConfirmationAccepted,
  ConfirmationAction,
  ConfirmationActions,
  ConfirmationRejected,
  ConfirmationRequest,
  ConfirmationTitle,
} from "../vendor/ai-elements/confirmation";
import {
  Conversation,
  ConversationContent,
  ConversationScrollButton,
} from "../vendor/ai-elements/conversation";
import { Message, MessageContent, MessageResponse } from "../vendor/ai-elements/message";
import {
  Queue,
  QueueItem,
  QueueItemAction,
  QueueItemActions,
  QueueItemContent,
  QueueItemIndicator,
  QueueList,
  QueueSection,
  QueueSectionContent,
  QueueSectionLabel,
  QueueSectionTrigger,
} from "../vendor/ai-elements/queue";
import { Shimmer } from "../vendor/ai-elements/shimmer";
import {
  Collapsible,
  CollapsibleContent,
  CollapsibleTrigger,
} from "../vendor/shadcn-ui/components/ui/collapsible";
import { Alert, AlertDescription } from "../vendor/shadcn-ui/components/ui/alert";
import { groupTimeline } from "../core/tool-groups.js";
import {
  decisionLabel,
  describeArgs,
  describeGroup,
  localToolName,
  mutationSummary,
  statusForTool,
} from "../core/labels.js";
import { presetText } from "../core/presets.js";
import { usePane, usePaneState, useT } from "./common";
import { ToolCard } from "./tool-card";
import { LogoMark, ToolIcon } from "./brand";
import { Composer } from "./composer";

const ALLOWED = new Set(["approve", "approve_turn", "disabled"]);

function approvalTarget(input: any) {
  if (!input || typeof input !== "object") return "";
  return (
    input.range ?? input.address ?? input.destinationRange ?? input.table ?? input.sheetName ?? ""
  );
}

function ApprovalCard({ item }: { item: any }) {
  const pane = usePane();
  const t = useT();
  const state =
    item.decision === null
      ? "approval-requested"
      : ALLOWED.has(item.decision)
        ? "approval-responded"
        : "output-denied";
  const label = decisionLabel(item.decision);
  return (
    <Confirmation
      className="rounded-xl border-amber-500/40 bg-amber-500/[0.06] shadow-[var(--shadow-card)]"
      approval={
        {
          id: item.requestId,
          approved: item.decision === null ? undefined : ALLOWED.has(item.decision),
        } as any
      }
      state={state as any}
    >
      <ConfirmationTitle>
        <ConfirmationRequest>
          <span className="flex items-center gap-2 font-semibold text-[13px] text-foreground">
            <ToolIcon name={localToolName(item.tool)} />
            {t("approval.question", { action: statusForTool(item.tool) })}
          </span>
          {/* What will change, in words; the raw arguments are one click away. */}
          <span className="mt-1 block text-muted-foreground text-xs">
            {[
              approvalTarget(item.input),
              mutationSummary(localToolName(item.tool), item.input ?? {}),
            ]
              .filter(Boolean)
              .join(" · ")}
          </span>
          <details className="mt-2 text-xs">
            <summary className="cursor-pointer select-none text-muted-foreground">
              {t("card.details")}
            </summary>
            <pre className="mt-1 max-h-48 overflow-auto whitespace-pre-wrap break-words rounded-md bg-muted/50 p-2 font-mono text-[11px]">
              {describeArgs(item.tool, item.input)}
            </pre>
          </details>
          {item.error && <p className="mt-1 text-destructive text-xs">{item.error}</p>}
        </ConfirmationRequest>
        <ConfirmationAccepted>{label}</ConfirmationAccepted>
        <ConfirmationRejected>{label}</ConfirmationRejected>
      </ConfirmationTitle>
      <ConfirmationActions>
        <ConfirmationAction
          disabled={item.busy}
          onClick={() => pane.respondApproval(item.requestId, "reject")}
          variant="outline"
        >
          {t("approval.reject")}
        </ConfirmationAction>
        <ConfirmationAction
          disabled={item.busy}
          onClick={() => pane.respondApproval(item.requestId, "approve_turn")}
          variant="outline"
        >
          {t("approval.approveTurn")}
        </ConfirmationAction>
        <ConfirmationAction
          disabled={item.busy}
          onClick={() => pane.respondApproval(item.requestId, "approve")}
        >
          {t("approval.approve")}
        </ConfirmationAction>
      </ConfirmationActions>
    </Confirmation>
  );
}

function TimelineItem({ item, streaming }: { item: any; streaming: boolean }) {
  const t = useT();
  // Entries the pane wrote carry a key; daemon and model text is shown as sent.
  const text = item.key ? t(item.key, item.params) : item.text;
  switch (item.kind) {
    case "user":
      return (
        <Message from="user">
          <MessageContent className="whitespace-pre-wrap group-[.is-user]:rounded-2xl group-[.is-user]:rounded-br-md group-[.is-user]:px-3.5 group-[.is-user]:py-2 leading-relaxed">
            {item.text}
          </MessageContent>
        </Message>
      );
    case "assistant":
      return (
        <Message from="assistant">
          <MessageContent className="leading-relaxed">
            <MessageResponse isAnimating={streaming}>{item.text}</MessageResponse>
          </MessageContent>
        </Message>
      );
    case "tool":
      return <ToolCard item={item} />;
    case "approval":
      return <ApprovalCard item={item} />;
    case "error":
      return (
        <Alert className="rounded-xl border-destructive/30 bg-destructive/5" variant="destructive">
          <AlertTriangleIcon />
          <AlertDescription className="whitespace-pre-wrap">{text}</AlertDescription>
        </Alert>
      );
    case "notice":
      return (
        <div className="flex gap-2 rounded-xl border border-border/60 bg-muted/40 px-3 py-2 text-muted-foreground text-xs leading-relaxed">
          <InfoIcon className="mt-0.5 size-3.5 shrink-0" />
          <span className="whitespace-pre-wrap">{text}</span>
        </div>
      );
    case "event":
      return <div className="text-center text-[11px] text-muted-foreground/80">{text}</div>;
    case "divider":
      return (
        <div className="flex items-center gap-3 py-1 text-[11px] text-muted-foreground/80">
          <span className="h-px flex-1 bg-border/70" />
          {text}
          <span className="h-px flex-1 bg-border/70" />
        </div>
      );
    default:
      return null;
  }
}

function ToolGroup({ group }: { group: any }) {
  useT(); // the group label follows the language
  return (
    <Collapsible
      className="rounded-xl border border-border/60 bg-card/40"
      defaultOpen={!group.collapsed}
    >
      <CollapsibleTrigger className="group flex w-full items-center gap-2 px-3 py-2 text-muted-foreground text-xs hover:text-foreground">
        <ToolIcon className="size-5" name={group.toolName} />
        {describeGroup(group.toolName, group.items.length)}
        <ChevronDownIcon className="ml-auto size-4 transition-transform group-data-[state=open]:rotate-180" />
      </CollapsibleTrigger>
      <CollapsibleContent className="space-y-2 px-2 pb-2">
        {group.items.map((item: any) => (
          <ToolCard item={item} key={item.id} />
        ))}
      </CollapsibleContent>
    </Collapsible>
  );
}

// The greeting and staggered suggestion cards follow Vercel's chatbot template
// (components/chat/greeting.tsx, suggested-actions.tsx): a short fade-up with
// its spring easing, then two columns of cards once the pane is wide enough.
const EASE = [0.22, 1, 0.36, 1] as const;

const PRESET_ICONS: Record<string, LucideIcon> = {
  summarize: ClipboardListIcon,
  explainSelection: ScanSearchIcon,
  totalsRow: SigmaIcon,
  checkData: ShieldCheckIcon,
  findValue: SearchIcon,
  contextFiles: LibraryIcon,
};

function greetingKey() {
  const hour = new Date().getHours();
  return hour < 5
    ? "greeting.night"
    : hour < 12
      ? "greeting.morning"
      : hour < 18
        ? "greeting.afternoon"
        : "greeting.evening";
}

function EmptyState() {
  const pane = usePane();
  const t = useT();
  const presets = usePaneState((state) => state.presets);
  const shown = useMemo(
    () =>
      [...presets].sort((a: any, b: any) => Number(!!b.pinned) - Number(!!a.pinned)).slice(0, 4),
    [presets],
  );
  const capabilities: [LucideIcon, string][] = [
    [TableIcon, t("empty.capRead")],
    [SigmaIcon, t("empty.capWrite")],
    [Undo2Icon, t("empty.capUndo")],
  ];
  return (
    <div className="@container flex min-h-full flex-col justify-center gap-6 px-1 py-8">
      <motion.div
        animate={{ opacity: 1, y: 0 }}
        className="flex flex-col items-center gap-3 text-center"
        initial={{ opacity: 0, y: 10 }}
        transition={{ duration: 0.5, ease: EASE }}
      >
        <div className="relative">
          <div className="absolute inset-0 scale-150 rounded-full bg-brand/25 blur-xl" />
          <LogoMark className="relative size-12 drop-shadow-md" />
        </div>
        <div className="space-y-1.5">
          <h2 className="font-semibold text-xl tracking-tight">{t(greetingKey())}</h2>
          <p className="text-muted-foreground text-sm">{t("empty.title")}</p>
        </div>
        <div className="flex flex-wrap justify-center gap-1.5">
          {capabilities.map(([Icon, label]) => (
            <span
              className="inline-flex items-center gap-1 rounded-full border border-border/60 bg-card/60 px-2.5 py-1 text-[11px] text-muted-foreground"
              key={label}
            >
              <Icon className="size-3 text-brand" />
              {label}
            </span>
          ))}
        </div>
      </motion.div>
      <div className="grid w-full gap-2 @[22rem]:grid-cols-2">
        {shown.map((preset: any, index: number) => {
          // Title only: the full prompt is what a click sends, not what the card shows.
          const { title } = presetText(preset);
          const Icon = PRESET_ICONS[preset.builtin] ?? SparklesIcon;
          return (
            <motion.button
              animate={{ opacity: 1, y: 0 }}
              className="group flex flex-col gap-1.5 rounded-xl border border-border/60 bg-card/50 p-3 text-left transition-all duration-200 hover:-translate-y-0.5 hover:bg-card hover:shadow-[var(--shadow-card)]"
              initial={{ opacity: 0, y: 16 }}
              key={preset.id}
              onClick={() => pane.usePreset(preset)}
              transition={{ delay: 0.15 + 0.06 * index, duration: 0.4, ease: EASE }}
              type="button"
            >
              <span className="flex items-center gap-2 font-medium text-[13px]">
                <span className="inline-flex size-6 items-center justify-center rounded-md bg-muted text-muted-foreground transition-colors group-hover:bg-brand/12 group-hover:text-brand">
                  <Icon className="size-3.5" />
                </span>
                {title}
              </span>
            </motion.button>
          );
        })}
      </div>
      <p className="text-center text-[11px] text-muted-foreground/80">{t("empty.hint")}</p>
    </div>
  );
}

function elapsed(since: number) {
  const seconds = Math.max(0, Math.round((Date.now() - since) / 1000));
  return seconds < 60 ? `${seconds}s` : `${Math.floor(seconds / 60)}m ${seconds % 60}s`;
}

// While the agent works, say how long it has been at it and how to stop it
// (pi-for-excel's working indicator leads with "escape to interrupt").
function WorkingIndicator() {
  const t = useT();
  const agent = usePaneState((state) => state.agent);
  const [, tick] = useState(0);
  useEffect(() => {
    if (agent.state !== "working") return;
    const timer = setInterval(() => tick((n) => n + 1), 1000);
    return () => clearInterval(timer);
  }, [agent.state]);
  if (agent.state !== "working") return null;
  return (
    <div className="flex items-center gap-2.5 py-1 text-sm">
      <span className="relative flex size-2">
        <span className="absolute inline-flex size-full animate-ping rounded-full bg-brand opacity-60" />
        <span className="relative inline-flex size-2 rounded-full bg-brand" />
      </span>
      <Shimmer as="span" duration={2}>
        {t(agent.key, agent.params)}
      </Shimmer>
      <span className="text-[11px] text-muted-foreground tabular-nums">
        {elapsed(agent.since)} · {t("working.escToStop")}
      </span>
    </div>
  );
}

function FollowUps() {
  const pane = usePane();
  const t = useT();
  const queue = usePaneState((state) => state.queue);
  if (queue.length === 0) return null;
  return (
    <Queue className="mx-3 mb-2">
      <QueueSection>
        <QueueSectionTrigger>
          <QueueSectionLabel count={queue.length} label={t("queue.label")} />
        </QueueSectionTrigger>
        <QueueSectionContent>
          <QueueList>
            {queue.map((turn: any) => (
              <QueueItem key={turn.id}>
                <div className="flex items-center gap-2">
                  <QueueItemIndicator />
                  <QueueItemContent title={turn.text}>{turn.text}</QueueItemContent>
                  <QueueItemActions>
                    <QueueItemAction
                      aria-label={t("queue.remove")}
                      onClick={() => pane.removeQueued(turn.id)}
                      title={t("queue.remove")}
                    >
                      <XIcon className="size-3" />
                    </QueueItemAction>
                  </QueueItemActions>
                </div>
              </QueueItem>
            ))}
          </QueueList>
        </QueueSectionContent>
      </QueueSection>
    </Queue>
  );
}

export function ChatView() {
  const items = usePaneState((state) => state.items);
  const showDiagnostics = usePaneState((state) => state.settings.showDiagnostics);
  const turnInFlight = usePaneState((state) => state.turnInFlight);

  // Reads and session events stay behind the diagnostics toggle; a change to the
  // workbook, an error, a notice or a decision always shows.
  const timeline = useMemo(() => {
    const visible = showDiagnostics
      ? items
      : items.filter(
          (item: any) => item.kind !== "event" && !(item.kind === "tool" && !item.write),
        );
    return groupTimeline(visible);
  }, [items, showDiagnostics]);
  const lastId = items.at(-1)?.id;

  return (
    <div className="flex min-h-0 flex-1 flex-col">
      {items.length === 0 ? (
        // The welcome page scrolls from its top; only a conversation sticks to
        // its newest message.
        <div className="min-h-0 flex-1 overflow-y-auto px-3">
          <EmptyState />
        </div>
      ) : (
        <Conversation className="min-h-0 flex-1">
          <ConversationContent className="gap-3 p-3">
            {timeline.map((entry: any) =>
              entry.kind === "group" ? (
                <ToolGroup group={entry} key={entry.id} />
              ) : (
                <TimelineItem
                  item={entry.item}
                  key={entry.item.id}
                  streaming={turnInFlight && entry.item.id === lastId}
                />
              ),
            )}
            <WorkingIndicator />
          </ConversationContent>
          <ConversationScrollButton />
        </Conversation>
      )}
      <FollowUps />
      <Composer />
    </div>
  );
}
