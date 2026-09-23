import { useEffect } from "react";
import {
  ArrowUpIcon,
  GaugeIcon,
  BookmarkIcon,
  CheckIcon,
  ChevronDownIcon,
  HandIcon,
  LoaderIcon,
  SettingsIcon,
  TableIcon,
  XIcon,
  ZapIcon,
  type LucideIcon,
} from "lucide-react";
import {
  Context,
  ContextCacheUsage,
  ContextContent,
  ContextContentBody,
  ContextContentFooter,
  ContextContentHeader,
  ContextInputUsage,
  ContextOutputUsage,
  ContextTrigger,
} from "../vendor/ai-elements/context";
import {
  PromptInput,
  PromptInputBody,
  PromptInputFooter,
  PromptInputHeader,
  PromptInputProvider,
  PromptInputSubmit,
  PromptInputTextarea,
  PromptInputTools,
  usePromptInputController,
} from "../vendor/ai-elements/prompt-input";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "../vendor/shadcn-ui/components/ui/dropdown-menu";
import { presetText } from "../core/presets.js";
import { formatTokens, usePane, usePaneState, useT } from "./common";

const PILL =
  "inline-flex h-7 shrink-0 items-center gap-1 whitespace-nowrap rounded-lg px-2 text-[12px] text-muted-foreground transition-colors hover:bg-accent hover:text-foreground data-[state=open]:bg-accent data-[state=open]:text-foreground";

// The input card follows Vercel's chatbot template (multimodal-input.tsx):
// a large radius, the composer shadow, a stronger shadow instead of a focus
// ring, and a dark round send button. Styled through the vendored InputGroup.
const COMPOSER =
  "[&>div]:rounded-2xl [&>div]:border-border/60 [&>div]:bg-card [&>div]:shadow-[var(--shadow-composer)] [&>div]:transition-shadow [&>div]:duration-300 focus-within:[&>div]:shadow-[var(--shadow-composer-focus)] [&>div]:has-[[data-slot=input-group-control]:focus-visible]:border-border [&>div]:has-[[data-slot=input-group-control]:focus-visible]:ring-0 dark:[&>div]:bg-card";

const TIERS = ["haiku", "sonnet", "opus"];

// The daemon's approvals are one on/off switch per pane.
const APPROVAL_MODES: [boolean, string, LucideIcon][] = [
  [true, "ask", HandIcon],
  [false, "auto", ZapIcon],
];

function SelectionChip() {
  const pane = usePane();
  const t = useT();
  const selection = usePaneState((state) => state.selection);
  if (!selection.attach || !selection.last?.text) return null;
  const text = selection.last.text;
  return (
    <PromptInputHeader className="px-2 pt-2">
      <span
        className="inline-flex h-6 max-w-full items-center gap-1.5 rounded-lg border border-border/60 bg-muted/50 pr-1 pl-2 text-[11.5px]"
        title={t("composer.selectionAttached", { text })}
      >
        <TableIcon className="size-3.5 shrink-0 text-brand" />
        <span className="truncate">{text.length > 40 ? `${text.slice(0, 39)}…` : text}</span>
        <button
          className="rounded p-0.5 text-muted-foreground hover:bg-accent hover:text-foreground"
          onClick={() => pane.detachSelection()}
          title={t("composer.detachSelection")}
          type="button"
        >
          <XIcon className="size-3" />
        </button>
      </span>
    </PromptInputHeader>
  );
}

function PresetMenu() {
  const pane = usePane();
  const t = useT();
  const presets = usePaneState((state) => state.presets);
  return (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <button className={PILL} title={t("presets.title")} type="button">
          <BookmarkIcon className="size-3.5" />
        </button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="start" className="w-64">
        <DropdownMenuLabel className="text-muted-foreground text-xs">
          {t("presets.title")}
        </DropdownMenuLabel>
        {presets.map((preset: any) => {
          const { title, prompt } = presetText(preset);
          return (
            <DropdownMenuItem
              className="flex-col items-start gap-0"
              key={preset.id}
              onSelect={() => pane.usePreset(preset)}
            >
              <span>{title}</span>
              <span className="line-clamp-1 text-muted-foreground text-xs">{prompt}</span>
            </DropdownMenuItem>
          );
        })}
        <DropdownMenuSeparator />
        <DropdownMenuItem onSelect={() => pane.setView("settings")}>
          <SettingsIcon className="size-4" />
          {t("presets.manage")}
        </DropdownMenuItem>
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

function ModelMenu() {
  const pane = usePane();
  const t = useT();
  const model = usePaneState((state) => state.settings.model);
  const tierModels = usePaneState((state) => state.tierModels);
  const pending = usePaneState((state) => state.pendingModel);
  const name = (tier: string) => tierModels[tier] ?? tier;
  return (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <button
          className={PILL}
          title={pending ? t("model.switching") : t("model.title")}
          type="button"
        >
          {pending && <LoaderIcon className="size-3 animate-spin" />}
          {name(model)}
          <ChevronDownIcon className="size-3 opacity-60" />
        </button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="start" className="w-56">
        <DropdownMenuLabel className="text-muted-foreground text-xs">
          {t("model.title")}
        </DropdownMenuLabel>
        {TIERS.map((tier) => (
          <DropdownMenuItem
            className="flex items-start gap-2"
            key={tier}
            onSelect={() => pane.setModel(tier)}
          >
            <div className="flex-1">
              <div>{name(tier)}</div>
              <div className="text-muted-foreground text-xs">{t(`model.hint.${tier}`)}</div>
            </div>
            {tier === model && <CheckIcon className="mt-0.5 size-4" />}
          </DropdownMenuItem>
        ))}
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

// Token counts and cost as the SDK reports them for this agent run. These are
// totals across turns, not how full the context window is, so no percentage.
function UsagePill() {
  const t = useT();
  const usage = usePaneState((state) => state.usage);
  const input = usage.input + usage.cacheRead + usage.cacheCreate;
  const total = input + usage.output;
  return (
    <Context
      maxTokens={usage.contextWindow || 1}
      usage={
        {
          inputTokens: input,
          outputTokens: usage.output,
          cachedInputTokens: usage.cacheRead,
          totalTokens: total,
        } as any
      }
      usedTokens={0}
    >
      <ContextTrigger>
        <button className={PILL} type="button">
          {/* Tokens only; the cost is in the hover card, so the bar stays one line. */}
          <GaugeIcon className="size-3.5" />
          {usage.turns ? formatTokens(total) : t("usage.title")}
        </button>
      </ContextTrigger>
      <ContextContent align="start" className="w-64">
        <ContextContentHeader>
          <div className="space-y-1 text-xs">
            <div className="font-medium">{t("usage.turns", { count: usage.turns })}</div>
            <div className="text-muted-foreground">
              {usage.contextWindow
                ? t("usage.window", { window: usage.contextWindow.toLocaleString() })
                : t("usage.cumulative")}
            </div>
          </div>
        </ContextContentHeader>
        <ContextContentBody className="space-y-1">
          <ContextInputUsage />
          <ContextOutputUsage />
          <ContextCacheUsage />
        </ContextContentBody>
        <ContextContentFooter>
          {usage.costGuessed ? (
            // Showing a figure the SDK priced off the wrong table would read as
            // a bill; say why there isn't one instead.
            <span className="text-muted-foreground">{t("usage.costNoTable")}</span>
          ) : (
            <>
              <span className="text-muted-foreground">{t("usage.cost")}</span>
              <span>{usage.priced ? `$${usage.cost.toFixed(4)}` : t("usage.notReported")}</span>
            </>
          )}
        </ContextContentFooter>
      </ContextContent>
    </Context>
  );
}

function ApprovalMenu() {
  const pane = usePane();
  const t = useT();
  const approveWrites = usePaneState((state) => state.settings.approveWrites);
  const [, mode, Icon] = APPROVAL_MODES.find(([ask]) => ask === approveWrites)!;
  return (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <button className={PILL} type="button">
          <Icon className="size-3.5" />
          {t(`approvalMode.${mode}`)}
        </button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="end" className="w-64">
        <DropdownMenuLabel className="text-muted-foreground text-xs">
          {t("approvalMode.title")}
        </DropdownMenuLabel>
        {APPROVAL_MODES.map(([ask, id, ModeIcon]) => (
          <DropdownMenuItem
            className="flex items-start gap-2.5"
            key={id}
            onSelect={() => pane.setApproveWrites(ask)}
          >
            <ModeIcon className="mt-0.5 size-4" />
            <div className="flex-1">
              <div>{t(`approvalMode.${id}`)}</div>
              <div className="text-muted-foreground text-xs">{t(`approvalMode.${id}.hint`)}</div>
            </div>
            {ask === approveWrites && <CheckIcon className="mt-0.5 size-4" />}
          </DropdownMenuItem>
        ))}
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

// A preset that does not send by itself fills the input for editing.
function DraftSync() {
  const pane = usePane();
  const draft = usePaneState((state) => state.draft);
  const controller = usePromptInputController();
  useEffect(() => {
    if (!draft) return;
    controller.textInput.setInput(draft.text);
    pane.consumeDraft();
    requestAnimationFrame(() => {
      const textarea = document.querySelector<HTMLTextAreaElement>("textarea[name=message]");
      textarea?.focus();
      textarea?.setSelectionRange(textarea.value.length, textarea.value.length);
    });
  }, [draft, controller, pane]);
  return null;
}

export function Composer() {
  const pane = usePane();
  const t = useT();
  const turnInFlight = usePaneState((state) => state.turnInFlight);
  const submitPending = usePaneState((state) => state.submitPending);
  const connected = usePaneState((state) => state.connection.state === "ok");
  const working = usePaneState((state) => state.agent.state === "working");

  return (
    <div className="px-3 pt-1 pb-3">
      <PromptInputProvider>
        <DraftSync />
        <PromptInput
          className={COMPOSER}
          onSubmit={async ({ text }) => {
            // Throwing keeps the text in the box when the message could not go out.
            if (!(await pane.sendUserTurn(text))) throw new Error("not sent");
          }}
        >
          <SelectionChip />
          <PromptInputBody>
            <PromptInputTextarea
              className="min-h-14 px-3.5 pt-3 text-[13.5px] leading-relaxed placeholder:text-muted-foreground/70"
              disabled={submitPending}
              placeholder={
                turnInFlight ? t("composer.placeholderQueued") : t("composer.placeholder")
              }
            />
          </PromptInputBody>
          <PromptInputFooter className="px-2 pb-2">
            <PromptInputTools className="min-w-0 gap-0.5 overflow-hidden">
              <PresetMenu />
              <ModelMenu />
              <UsagePill />
            </PromptInputTools>
            <div className="flex items-center gap-1">
              <ApprovalMenu />
              <PromptInputSubmit
                className="size-8 rounded-xl transition-transform active:scale-95"
                disabled={!connected && !working}
                onStop={() => pane.stop()}
                status={working ? "streaming" : submitPending ? "submitted" : "ready"}
              >
                {working || submitPending ? undefined : <ArrowUpIcon className="size-4" />}
              </PromptInputSubmit>
            </div>
          </PromptInputFooter>
        </PromptInput>
      </PromptInputProvider>
    </div>
  );
}
