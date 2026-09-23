// The product mark and the icon each kind of tool call shows on its card.
import {
  ArrowUpDownIcon,
  CalculatorIcon,
  FileTextIcon,
  FilterIcon,
  GlobeIcon,
  HistoryIcon,
  MousePointerClickIcon,
  PaintbrushIcon,
  PencilLineIcon,
  Rows3Icon,
  SearchIcon,
  SheetIcon,
  SigmaIcon,
  SquareTerminalIcon,
  Table2Icon,
  TableIcon,
  WrenchIcon,
  EraserIcon,
  CopyIcon,
  ChartColumnIcon,
  type LucideIcon,
} from "lucide-react";
import { useId } from "react";
import { cn } from "../vendor/shadcn-ui/lib/utils";

// A green tile with a 2×2 grid and a spark in the free cell: a sheet that works
// by itself. Drawn on a 24-unit grid so it stays crisp at 20px.
export function LogoMark({ className }: { className?: string }) {
  const id = useId();
  return (
    <svg aria-hidden="true" className={cn("size-5", className)} viewBox="0 0 24 24">
      <defs>
        <linearGradient id={`${id}-bg`} x1="0" x2="1" y1="0" y2="1">
          <stop offset="0" stopColor="oklch(0.7 0.16 150)" />
          <stop offset="1" stopColor="oklch(0.52 0.13 175)" />
        </linearGradient>
      </defs>
      <rect fill={`url(#${id}-bg)`} height="24" rx="6.5" width="24" />
      <g fill="white">
        <rect height="5.5" opacity="0.95" rx="1.4" width="5.5" x="5.5" y="5.5" />
        <rect height="5.5" opacity="0.6" rx="1.4" width="5.5" x="13" y="5.5" />
        <rect height="5.5" opacity="0.6" rx="1.4" width="5.5" x="5.5" y="13" />
        <path d="M15.75 12.4l.85 2.15 2.15.85-2.15.85-.85 2.15-.85-2.15-2.15-.85 2.15-.85z" />
      </g>
    </svg>
  );
}

type Tone = "read" | "write" | "neutral";

const TOOL_ICONS: Record<string, [LucideIcon, Tone]> = {
  excel_get_workbook_metadata: [SheetIcon, "read"],
  excel_get_selected_range: [MousePointerClickIcon, "read"],
  excel_get_cell_ranges: [TableIcon, "read"],
  excel_get_range_as_csv: [TableIcon, "read"],
  excel_search_data: [SearchIcon, "read"],
  excel_get_all_objects: [ChartColumnIcon, "read"],
  excel_explain_formula: [SigmaIcon, "read"],
  excel_trace_dependencies: [SigmaIcon, "read"],
  excel_set_cell_range: [PencilLineIcon, "write"],
  excel_clear_cell_range: [EraserIcon, "write"],
  excel_copy_to: [CopyIcon, "write"],
  excel_modify_sheet_structure: [Rows3Icon, "write"],
  excel_modify_workbook_structure: [SheetIcon, "write"],
  excel_resize_range: [Rows3Icon, "write"],
  excel_modify_object: [ChartColumnIcon, "write"],
  excel_select_range: [MousePointerClickIcon, "neutral"],
  excel_set_format: [PaintbrushIcon, "write"],
  excel_sort_range: [ArrowUpDownIcon, "write"],
  excel_autofilter: [FilterIcon, "write"],
  excel_create_table: [Table2Icon, "write"],
  excel_add_table_rows: [Table2Icon, "write"],
  excel_workbook_history: [HistoryIcon, "write"],
  excel_bash: [CalculatorIcon, "neutral"],
  Bash: [SquareTerminalIcon, "neutral"],
  Read: [FileTextIcon, "neutral"],
  Write: [FileTextIcon, "neutral"],
  Edit: [FileTextIcon, "neutral"],
  MultiEdit: [FileTextIcon, "neutral"],
  Glob: [SearchIcon, "neutral"],
  Grep: [SearchIcon, "neutral"],
  WebFetch: [GlobeIcon, "neutral"],
  WebSearch: [GlobeIcon, "neutral"],
};

const TONES: Record<Tone, string> = {
  read: "bg-sky-500/10 text-sky-600 dark:text-sky-400",
  write: "bg-brand/12 text-brand",
  neutral: "bg-muted text-muted-foreground",
};

export function ToolIcon({ name, className }: { name: string; className?: string }) {
  const [Icon, tone] = TOOL_ICONS[name] ?? [WrenchIcon, "neutral" as Tone];
  return (
    <span
      className={cn(
        "inline-flex size-6 shrink-0 items-center justify-center rounded-md",
        TONES[tone],
        className,
      )}
    >
      <Icon className="size-3.5" />
    </span>
  );
}
