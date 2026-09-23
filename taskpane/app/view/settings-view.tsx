import { useEffect, useState } from "react";
import { PencilIcon, PinIcon, PinOffIcon, PlusIcon, XIcon } from "lucide-react";
import { Button } from "../vendor/shadcn-ui/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "../vendor/shadcn-ui/components/ui/dialog";
import { Input } from "../vendor/shadcn-ui/components/ui/input";
import { Label } from "../vendor/shadcn-ui/components/ui/label";
import { Switch } from "../vendor/shadcn-ui/components/ui/switch";
import { Textarea } from "../vendor/shadcn-ui/components/ui/textarea";
import { cn } from "../vendor/shadcn-ui/lib/utils";
import { presetText } from "../core/presets.js";
import { Section, usePane, usePaneState, useT } from "./common";

function WorkspaceSection() {
  const pane = usePane();
  const t = useT();
  const workspace = usePaneState((state) => state.workspace);
  return (
    <Section description={t("workspace.hint")} title={t("workspace.title")}>
      <div
        className="break-all rounded-md bg-muted/60 px-2 py-1.5 font-mono text-xs"
        title={workspace.cwd ?? ""}
      >
        {workspace.cwd ?? t("workspace.noneParen")}
      </div>
      {workspace.mismatch && <p className="text-amber-600 text-xs">{t("workspace.mismatch")}</p>}
      <Button onClick={() => pane.changeWorkspace()} size="sm" variant="outline">
        {t("workspace.change")}
      </Button>
      {workspace.error && <p className="text-destructive text-xs">{workspace.error}</p>}
    </Section>
  );
}

function ContextDialog() {
  const pane = usePane();
  const t = useT();
  const dialog = usePaneState((state) => state.context.dialog);
  const [path, setPath] = useState("");
  const [description, setDescription] = useState("");
  const [saving, setSaving] = useState(false);
  useEffect(() => {
    setPath(dialog?.path ?? "");
    setDescription("");
  }, [dialog?.path]);
  const browse = async (includeFiles: boolean) => {
    const picked = await pane.browseContextPath(includeFiles, path.trim()).catch(() => null);
    if (picked) setPath(picked);
  };
  const save = async () => {
    setSaving(true);
    try {
      await pane.addContextEntry(path.trim(), description.trim());
    } finally {
      setSaving(false);
    }
  };
  return (
    <Dialog onOpenChange={(open) => !open && pane.closeContextDialog()} open={Boolean(dialog)}>
      <DialogContent className="max-w-[calc(100%-1.5rem)] p-4">
        <DialogHeader>
          <DialogTitle>
            {dialog?.kind === "file" ? t("context.addFileTitle") : t("context.addFolderTitle")}
          </DialogTitle>
        </DialogHeader>
        <div className="space-y-2">
          <Label htmlFor="context-path">{t("context.path")}</Label>
          <Input
            id="context-path"
            onChange={(e) => setPath(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && save()}
            value={path}
          />
          <div className="flex gap-1">
            <Button onClick={() => browse(true)} size="sm" variant="ghost">
              {t("context.chooseFile")}
            </Button>
            <Button onClick={() => browse(false)} size="sm" variant="ghost">
              {t("context.chooseFolder")}
            </Button>
          </div>
          <Label htmlFor="context-description">{t("context.description")}</Label>
          <Input
            autoFocus
            id="context-description"
            onChange={(e) => setDescription(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && save()}
            placeholder={t("context.descriptionPlaceholder")}
            value={description}
          />
          {dialog?.error && <p className="text-destructive text-xs">{dialog.error}</p>}
        </div>
        <DialogFooter>
          <Button onClick={() => pane.closeContextDialog()} variant="outline">
            {t("common.cancel")}
          </Button>
          <Button disabled={saving} onClick={save}>
            {t("common.add")}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}

function ContextSection() {
  const pane = usePane();
  const t = useT();
  const context = usePaneState((state) => state.context);
  return (
    <Section description={t("context.hint")} title={t("context.title")}>
      <div className="space-y-1">
        {context.loading && !context.entries && (
          <p className="text-muted-foreground text-xs">{t("common.loading")}</p>
        )}
        {context.entries?.length === 0 && (
          <p className="text-muted-foreground text-xs">{t("context.empty")}</p>
        )}
        {context.entries?.map((entry: any, index: number) => (
          <div
            className="flex items-start gap-2 rounded-md border px-2 py-1.5"
            key={`${entry.path}:${index}`}
            title={entry.path}
          >
            <div className="min-w-0 flex-1">
              <div className="break-all font-mono text-xs">
                <span className="mr-1 rounded bg-muted px-1 text-muted-foreground">
                  {entry.kind === "file"
                    ? t("context.kindFile")
                    : entry.kind === "folder"
                      ? t("context.kindFolder")
                      : "?"}
                </span>
                {entry.path}
              </div>
              {entry.description && (
                <div className="text-muted-foreground text-xs">{entry.description}</div>
              )}
            </div>
            <Button
              onClick={() => pane.removeContextEntry(index)}
              size="icon-sm"
              title={t("common.remove")}
              variant="ghost"
            >
              <XIcon className="size-3.5" />
            </Button>
          </div>
        ))}
      </div>
      <div className="flex gap-1">
        <Button onClick={() => pane.beginAddContext(false)} size="sm" variant="outline">
          <PlusIcon className="size-3.5" />
          {t("context.kindFolder")}
        </Button>
        <Button onClick={() => pane.beginAddContext(true)} size="sm" variant="outline">
          <PlusIcon className="size-3.5" />
          {t("context.kindFile")}
        </Button>
      </div>
      {context.error && <p className="text-destructive text-xs">{context.error}</p>}
      <ContextDialog />
    </Section>
  );
}

const TIERS = ["haiku", "sonnet", "opus"];

function ProviderSection() {
  const pane = usePane();
  const t = useT();
  const provider = usePaneState((state) => state.provider);
  const [baseUrl, setBaseUrl] = useState("");
  const [credential, setCredential] = useState("");
  const [clearCredential, setClearCredential] = useState(false);
  const [models, setModels] = useState<Record<string, string>>({});
  const [status, setStatus] = useState<{ text: string; error: boolean } | null>(null);
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    if (!provider) return;
    setBaseUrl(provider.baseUrl);
    setModels(provider.models);
    setCredential("");
  }, [provider]);

  const save = async () => {
    setSaving(true);
    setStatus(null);
    try {
      const text = await pane.saveProvider({
        base_url: baseUrl,
        credential,
        clear_credential: clearCredential,
        models: Object.fromEntries(TIERS.map((tier) => [tier, models[tier] ?? ""])),
      });
      setStatus({ text, error: false });
    } catch (error: any) {
      setStatus({ text: error?.message ?? String(error), error: true });
    } finally {
      setSaving(false);
    }
  };

  return (
    <Section
      description={t("provider.name", { name: provider?.name ?? t("common.loading") })}
      title={t("provider.title")}
    >
      <div className="space-y-2">
        <Label htmlFor="provider-url">{t("provider.baseUrl")}</Label>
        <Input
          id="provider-url"
          onChange={(e) => setBaseUrl(e.target.value)}
          placeholder={t("provider.baseUrlPlaceholder")}
          value={baseUrl}
        />
        <Label htmlFor="provider-credential">{t("provider.credential")}</Label>
        <Input
          autoComplete="new-password"
          id="provider-credential"
          onChange={(e) => setCredential(e.target.value)}
          placeholder={
            provider?.credentialConfigured
              ? t("provider.credentialSaved")
              : t("provider.credentialPlaceholder")
          }
          type="password"
          value={credential}
        />
        <div className="flex items-center gap-2">
          <Switch
            checked={clearCredential}
            id="provider-clear"
            onCheckedChange={setClearCredential}
          />
          <Label className="font-normal text-xs" htmlFor="provider-clear">
            {t("provider.clearCredential")}
          </Label>
        </div>
        {TIERS.map((tier) => (
          <div className="space-y-1" key={tier}>
            <Label htmlFor={`provider-model-${tier}`}>{t(`provider.tier.${tier}`)}</Label>
            <Input
              id={`provider-model-${tier}`}
              onChange={(e) => setModels((current) => ({ ...current, [tier]: e.target.value }))}
              placeholder={tier}
              value={models[tier] ?? ""}
            />
          </div>
        ))}
        <Button disabled={saving} onClick={save} size="sm">
          {t("provider.save")}
        </Button>
        {status && (
          <p className={`text-xs ${status.error ? "text-destructive" : "text-muted-foreground"}`}>
            {status.text}
          </p>
        )}
      </div>
    </Section>
  );
}

// A row of mutually exclusive choices.
function Choice({
  value,
  options,
  onChange,
}: {
  value: string;
  options: [string, string][];
  onChange: (value: string) => void;
}) {
  return (
    <div className="inline-flex rounded-md border p-0.5">
      {options.map(([id, label]) => (
        <button
          className={cn(
            "rounded px-2.5 py-1 text-xs",
            id === value ? "bg-primary text-primary-foreground" : "hover:bg-accent",
          )}
          key={id}
          onClick={() => onChange(id)}
          type="button"
        >
          {label}
        </button>
      ))}
    </div>
  );
}

function PreferencesSection() {
  const pane = usePane();
  const t = useT();
  const settings = usePaneState((state) => state.settings);
  const switches: [string, string, boolean, (value: boolean) => void][] = [
    [t("prefs.approve"), t("prefs.approveHint"), settings.approveWrites, pane.setApproveWrites],
    [
      t("prefs.diagnostics"),
      t("prefs.diagnosticsHint"),
      settings.showDiagnostics,
      pane.setShowDiagnostics,
    ],
  ];
  return (
    <Section title={t("prefs.title")}>
      <div className="flex items-center justify-between gap-3">
        <div className="text-sm">{t("prefs.language")}</div>
        <Choice
          onChange={pane.setLanguage}
          // Language names are written in their own language.
          options={[
            ["auto", t("prefs.followOffice")],
            ["zh", "中文"],
            ["en", "English"],
          ]}
          value={settings.language}
        />
      </div>
      <div className="flex items-center justify-between gap-3">
        <div className="text-sm">{t("prefs.theme")}</div>
        <Choice
          onChange={pane.setTheme}
          options={[
            ["auto", t("prefs.followExcel")],
            ["light", t("prefs.light")],
            ["dark", t("prefs.dark")],
          ]}
          value={settings.theme}
        />
      </div>
      {switches.map(([title, hint, checked, onChange]) => (
        <div className="flex items-start justify-between gap-3" key={title}>
          <div>
            <div className="text-sm">{title}</div>
            <div className="text-muted-foreground text-xs">{hint}</div>
          </div>
          <Switch checked={checked} onCheckedChange={onChange} />
        </div>
      ))}
    </Section>
  );
}

function PresetDialog({
  preset,
  onClose,
}: {
  preset: any | null | undefined;
  onClose: () => void;
}) {
  const pane = usePane();
  const t = useT();
  const [form, setForm] = useState({
    title: "",
    prompt: "",
    category: "",
    auto_send: false,
    pinned: false,
  });
  useEffect(() => {
    if (preset === undefined) return;
    const text = preset ? presetText(preset) : { title: "", prompt: "", category: "" };
    setForm({
      ...text,
      auto_send: Boolean(preset?.auto_send),
      pinned: Boolean(preset?.pinned),
    });
  }, [preset]);
  const valid = form.title.trim() && form.prompt.trim();
  const field = (key: keyof typeof form) => (value: any) =>
    setForm((current) => ({ ...current, [key]: value }));
  return (
    <Dialog onOpenChange={(open) => !open && onClose()} open={preset !== undefined}>
      <DialogContent className="max-w-[calc(100%-1.5rem)] p-4">
        <DialogHeader>
          <DialogTitle>{preset ? t("presets.editTitle") : t("presets.newTitle")}</DialogTitle>
        </DialogHeader>
        <div className="space-y-2">
          <Label htmlFor="preset-title">{t("presets.fieldTitle")}</Label>
          <Input
            id="preset-title"
            onChange={(e) => field("title")(e.target.value)}
            value={form.title}
          />
          <Label htmlFor="preset-prompt">{t("presets.fieldPrompt")}</Label>
          <Textarea
            id="preset-prompt"
            onChange={(e) => field("prompt")(e.target.value)}
            rows={5}
            value={form.prompt}
          />
          <Label htmlFor="preset-category">{t("presets.fieldCategory")}</Label>
          <Input
            id="preset-category"
            onChange={(e) => field("category")(e.target.value)}
            value={form.category}
          />
          <div className="flex items-center gap-2">
            <Switch
              checked={form.auto_send}
              id="preset-auto"
              onCheckedChange={field("auto_send")}
            />
            <Label className="font-normal text-xs" htmlFor="preset-auto">
              {t("presets.autoSend")}
            </Label>
          </div>
          <div className="flex items-center gap-2">
            <Switch checked={form.pinned} id="preset-pinned" onCheckedChange={field("pinned")} />
            <Label className="font-normal text-xs" htmlFor="preset-pinned">
              {t("presets.pinned")}
            </Label>
          </div>
        </div>
        <DialogFooter>
          <Button onClick={onClose} variant="outline">
            {t("common.cancel")}
          </Button>
          <Button
            disabled={!valid}
            onClick={() => {
              pane.upsertPreset(
                { ...form, title: form.title.trim(), category: form.category.trim() },
                preset?.id ?? null,
              );
              onClose();
            }}
          >
            {t("common.save")}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}

function PresetsSection() {
  const pane = usePane();
  const t = useT();
  const presets = usePaneState((state) => state.presets);
  // undefined: closed; null: new; object: editing.
  const [editing, setEditing] = useState<any | null | undefined>(undefined);
  return (
    <Section
      action={
        <Button onClick={() => setEditing(null)} size="sm" variant="outline">
          <PlusIcon className="size-3.5" />
          {t("presets.new")}
        </Button>
      }
      description={t("presets.hint")}
      title={t("presets.title")}
    >
      <div className="space-y-1">
        {presets.length === 0 && (
          <p className="text-muted-foreground text-xs">{t("presets.empty")}</p>
        )}
        {presets.map((preset: any) => {
          const { title, prompt, category } = presetText(preset);
          return (
            <div
              className="flex items-center gap-1 rounded-md border px-2 py-1"
              key={preset.id}
              title={prompt}
            >
              <button
                className="min-w-0 flex-1 truncate text-left text-sm"
                onClick={() => pane.usePreset(preset)}
                type="button"
              >
                {title}
                {category && <span className="ml-2 text-muted-foreground text-xs">{category}</span>}
              </button>
              <Button
                onClick={() => pane.togglePin(preset.id)}
                size="icon-sm"
                title={preset.pinned ? t("presets.unpin") : t("presets.pin")}
                variant="ghost"
              >
                {preset.pinned ? (
                  <PinIcon className="size-3.5" />
                ) : (
                  <PinOffIcon className="size-3.5 opacity-50" />
                )}
              </Button>
              <Button
                onClick={() => setEditing(preset)}
                size="icon-sm"
                title={t("common.edit")}
                variant="ghost"
              >
                <PencilIcon className="size-3.5" />
              </Button>
              <Button
                onClick={() => pane.deletePreset(preset.id)}
                size="icon-sm"
                title={t("common.delete")}
                variant="ghost"
              >
                <XIcon className="size-3.5" />
              </Button>
            </div>
          );
        })}
      </div>
      <PresetDialog onClose={() => setEditing(undefined)} preset={editing} />
    </Section>
  );
}

export function SettingsView() {
  return (
    <div className="flex min-h-0 flex-1 flex-col gap-3 overflow-y-auto p-3">
      <PreferencesSection />
      <WorkspaceSection />
      <ContextSection />
      <PresetsSection />
      <ProviderSection />
    </div>
  );
}
