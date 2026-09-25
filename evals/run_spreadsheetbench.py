"""Run SpreadsheetBench tasks through Excel Assistant in the live Excel instance.

For each task: copy the initial workbook into its own folder (fresh agent
session), tag it so the task pane auto-opens, open it in Excel, inject the
prompt through the daemon's /eval/run endpoint, save, grade with the official
SpreadsheetBench comparison, and count edits outside the authorized range.

Each run directory is pinned to one configuration (manifest.json); resuming
with a different commit, model mapping, prompt or task set is refused.

  uv run --project evals python evals/run_spreadsheetbench.py \
      --dataset <.../spreadsheetbench_verified_400> --run dev40-qwen-plus --sample 40
"""

import argparse
import hashlib
import json
import random
import shutil
import statistics
import subprocess
import sys
import time
import urllib.request
import openpyxl
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path

from embed_taskpane import embed_taskpane, read_manifest
from workbook_diff import unauthorized_edits, parse_answer_position, compare_answer_workbooks

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DAEMON = "http://127.0.0.1:47834"
TOKEN_FILE = Path.home() / ".claude" / "office-addins" / "bridge-token"


@contextmanager
def managed_workbook(app, workbook_path):
    """Open one evaluation workbook and always restore Excel-wide UI flags."""
    alerts, visible = app.DisplayAlerts, app.Visible
    workbook = None
    try:
        app.DisplayAlerts = False
        app.Visible = True
        workbook = app.Workbooks.Open(str(workbook_path), UpdateLinks=0)
        yield workbook
    finally:
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        finally:
            app.DisplayAlerts, app.Visible = alerts, visible

# Field descriptions follow SpreadsheetBench's PROMPT_NO_DF_RCT_FORMAT, adapted
# from "write Python to an output file" to editing the open workbook in place.
# SpreadsheetBench (https://github.com/RUCKBReasoning/SpreadsheetBench) is
# CC BY-SA 4.0; this adapted prompt text is shared under the same license.
PROMPT = """You need to solve the given spreadsheet manipulation question, which contains three types of information:
- instruction: The question about spreadsheet manipulation.
- instruction_type: There are two values (Cell-Level Manipulation, Sheet-Level Manipulation) used to indicate whether the answer to this question applies only to specific cells or to the entire worksheet.
- answer_position: The position need to be modified or filled. For Cell-Level Manipulation questions, this field is filled with the cell position; for Sheet-Level Manipulation, it is the maximum range of cells you need to modify. You only need to modify or fill in values within the cell range specified by answer_position.

The workbook is already open in Excel. Make the changes directly in it; do not create other files. You are authorized to overwrite existing cells within answer_position when the instruction requires it, but do not change cells outside that scope. Do not ask clarifying questions — complete the task.

### instruction
{instruction}

### instruction_type
{instruction_type}

### answer_position
{answer_position}
"""

# Daemon results that mean the harness, not the model, failed.
INFRA_STATUSES = {"no_pane", "busy", "bad_request"}
XL_MAXIMIZED = -4137


def qualified_answer_position(task: dict, sheet_names: list[str] | None = None) -> str:
    position = task["answer_position"]
    sheet = task.get("answer_sheet")
    ranges = parse_answer_position(position)
    if sheet and sheet_names and sheet not in sheet_names and any(name is None for name, _ in ranges):
        # Some records list all involved sheets here. Unqualified addresses
        # follow the official grader and refer to the workbook's first sheet.
        if all(name.strip() in sheet_names for name in sheet.split(",")):
            sheet = sheet_names[0]
        else:
            raise ValueError(f"answer_sheet does not exist in workbook: {sheet}")
    return ",".join(
        "'" + (name or sheet).replace("'", "''") + "'!" + cells if name or sheet else cells
        for name, cells in ranges
    )


def daemon_request(path: str, token: str, payload: dict | None = None, timeout: float = 30) -> dict:
    request = urllib.request.Request(
        DAEMON + path,
        data=None if payload is None else json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json", "x-bridge-token": token},
        method="GET" if payload is None else "POST",
    )
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    with opener.open(request, timeout=timeout) as response:
        return json.load(response)


def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def git(*args: str, check: bool = True) -> str:
    return subprocess.run(
        ["git", "-C", str(PROJECT_ROOT), *args], capture_output=True, text=True, encoding="utf-8", check=check
    ).stdout


def uncommitted_patch() -> str:
    """Capture project implementation changes without copying local workbooks or notes."""
    paths = ("app", "daemon", "taskpane", "scripts", "tests", "evals", "manifests",
             "package.json", "package-lock.json", "agent.config.example.json", "CLAUDE.md")
    patch = git("diff", "HEAD", "--", *paths)
    for path in git("ls-files", "--others", "--exclude-standard", "--", *paths).splitlines():
        if path.startswith("evals/runs/"):
            continue
        # --no-index exits 1 when the files differ, which is always the case here.
        patch += git("diff", "--no-index", "--", "/dev/null", path, check=False)
    return patch


def select_tasks(tasks: list[dict], args) -> list[dict]:
    if args.ids:
        wanted = [str(i) for i in args.ids]
        by_id = {str(t["id"]): t for t in tasks}
        return [by_id[i] for i in wanted]
    if args.sample:
        # Stratified by instruction type so a small dev set keeps the benchmark's mix.
        rng = random.Random(args.seed)
        groups: dict[str, list[dict]] = {}
        for task in tasks:
            groups.setdefault(task["instruction_type"], []).append(task)
        chosen = []
        for kind in sorted(groups):
            share = round(args.sample * len(groups[kind]) / len(tasks))
            chosen += rng.sample(groups[kind], min(share, len(groups[kind])))
        return sorted(chosen, key=lambda t: str(t["id"]))
    return tasks[args.offset : args.offset + args.limit]


def build_manifest(args, selected: list[dict], info: dict) -> dict:
    dirty_diff = uncommitted_patch()
    return {
        "config": {
            "commit": git("rev-parse", "HEAD").strip(),
            "uncommittedDiffSha256": hashlib.sha256(dirty_diff.encode("utf-8")).hexdigest() if dirty_diff else None,
            "vendorSha256": sha256_file(PROJECT_ROOT / "taskpane/shared/vendor/office-agents-excel-api.js"),
            "datasetSha256": sha256_file(args.dataset / "dataset.json"),
            "taskIds": [str(t["id"]) for t in selected],
            "tier": args.model,
            "timeoutSeconds": args.timeout,
            "promptTemplateSha256": hashlib.sha256(PROMPT.encode("utf-8")).hexdigest(),
            "agent": info,
        },
        "createdAt": datetime.now(timezone.utc).isoformat(),
        "dataset": str(args.dataset),
        "selection": {"ids": args.ids, "sample": args.sample, "seed": args.seed, "offset": args.offset, "limit": args.limit},
    }


def excel_app():
    import win32com.client

    try:
        app = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        app = win32com.client.Dispatch("Excel.Application")
    return app


EXCEL_READY_TIMEOUT_S = 600


class ExcelUnavailable(RuntimeError):
    """Excel stopped answering COM calls; later tasks would fail the same way."""


def ready_excel(run_dir: Path):
    # A task can leave Excel recalculating for a long time (e.g. whole-column
    # array formulas). Wait for it, and close eval workbooks a failed task left
    # open, instead of failing every remaining task in seconds.
    deadline = time.monotonic() + EXCEL_READY_TIMEOUT_S
    last_error = None
    while time.monotonic() < deadline:
        try:
            app = excel_app()
            if app.Ready:
                for workbook in list(app.Workbooks):
                    if Path(workbook.FullName).resolve().is_relative_to(run_dir.resolve()):
                        workbook.Close(SaveChanges=False)
                return app
        except Exception as exc:
            last_error = exc
        time.sleep(5)
    raise ExcelUnavailable(f"Excel did not respond within {EXCEL_READY_TIMEOUT_S}s: {last_error}")


def tool_error_summary(transcript: Path | None) -> dict:
    if not transcript or not transcript.exists():
        return {"count": 0, "categories": {}, "first": None}
    errors = []
    for line in transcript.read_text(encoding="utf-8").splitlines():
        try:
            entry = json.loads(line)
        except json.JSONDecodeError:
            continue
        content = (entry.get("message") or {}).get("content")
        if entry.get("type") == "user" and isinstance(content, list):
            for block in content:
                if block.get("type") != "tool_result" or not block.get("is_error"):
                    continue
                value = block.get("content")
                if isinstance(value, list):
                    value = " ".join(str(part.get("text", "")) for part in value if isinstance(part, dict))
                errors.append(str(value or "Unknown tool error"))
    categories = {}
    for error in errors:
        lowered = error.lower()
        if any(term in lowered for term in ("could not be parsed as json", "not valid json", "invalid json")):
            category = "model_malformed_json"
        elif "spill outside range" in lowered or "flat list" in lowered:
            category = "cell_shape"
        elif "input validation error" in lowered:
            category = "argument_validation"
        elif "would overwrite" in lowered:
            category = "overwrite_guard"
        elif "worksheet with id" in lowered and "not found" in lowered:
            category = "stale_sheet_id"
        elif "no such tool available" in lowered:
            category = "tool_name"
        elif any(term in lowered for term in ("modulenotfounderror", "powershell: command not found", "no such file or directory")):
            category = "sandbox_environment"
        elif "timed out" in lowered or "timeout" in lowered:
            category = "timeout"
        else:
            category = "other"
        categories[category] = categories.get(category, 0) + 1
    return {"count": len(errors), "categories": categories, "first": errors[0][:500] if errors else None}


def classify_outcome(infra_status: str, passed: bool, agent_status: str | None) -> str:
    """Describe the observed outcome, not an inferred root cause."""
    if infra_status != "ok":
        return "infrastructure"
    if passed:
        return "passed"
    return "answer_mismatch" if agent_status == "completed" else "agent_incomplete"


def harness_error_record(partial: dict, exc: Exception) -> dict:
    """Retain evidence collected before an outer COM/harness exception."""
    return {
        **partial,
        "passed": False,
        "infra_status": "harness_error",
        "failure_class": "infrastructure",
        "unauthorized_cells": None,
        "tool_calls": partial.get("tool_calls", 0),
        "tool_errors": partial.get("tool_errors", 0),
        "agent_duration_s": partial.get("agent_duration_s", 0),
        "agent_status": partial.get("agent_status"),
        "error": f"{type(exc).__name__}: {exc}",
    }


def run_task(task: dict, dataset: Path, run_dir: Path, args, token: str, addin: tuple[str, str], compare,
             record: dict | None = None) -> dict:
    task_id = str(task["id"])
    record = record if record is not None else {"id": task_id, "instruction_type": task["instruction_type"]}
    phases: dict[str, float] = record.setdefault("phases_s", {})

    def phase(name: str, started: float) -> None:
        phases[name] = round(time.perf_counter() - started, 2)

    started = time.perf_counter()
    record["stage"] = "prepare"
    source_dir = dataset / task["spreadsheet_path"]
    # Most folders use N_<id>_init/golden.xlsx; a few use initial.xlsx/golden.xlsx.
    init_file = next(p for p in source_dir.glob("*.xlsx") if p.stem.endswith(("_init", "initial")))
    golden_file = next(p for p in source_dir.glob("*.xlsx") if p.stem.endswith("golden"))
    task_dir = run_dir / task_id
    shutil.rmtree(task_dir, ignore_errors=True)
    task_dir.mkdir(parents=True)
    workbook_path = task_dir / f"{task_id}.xlsx"
    shutil.copyfile(init_file, workbook_path)
    embed_taskpane(workbook_path, *addin)
    source_workbook = openpyxl.load_workbook(init_file, read_only=True)
    try:
        answer_position = qualified_answer_position(task, source_workbook.sheetnames)
    finally:
        source_workbook.close()
    prompt = PROMPT.format(
        instruction=task["instruction"], instruction_type=task["instruction_type"], answer_position=answer_position
    )
    phase("prepare", started)

    record["stage"] = "excel_ready"
    app = ready_excel(run_dir)
    started = time.perf_counter()
    record["stage"] = "open"
    with managed_workbook(app, workbook_path) as workbook:
        # Many benchmark files were saved minimized; Excel then never loads the task pane.
        for window in workbook.Windows:
            window.WindowState = XL_MAXIMIZED
        phase("open", started)
        started = time.perf_counter()
        record["stage"] = "agent"
        agent = daemon_request(
            "/eval/run",
            token,
            {"doc": str(workbook_path), "prompt": prompt, "model": args.model, "timeoutMs": args.timeout * 1000},
            timeout=args.timeout + 180,
        )
        phase("agent", started)
        record.update({
            "agent_status": agent.get("status"),
            "tool_calls": len(agent.get("tools", [])),
            "agent_duration_s": round((agent.get("monotonicMs") or 0) / 1000, 1),
            "agent_wall_s": round((agent.get("durationMs") or 0) / 1000, 1),
            "task_check": agent.get("taskCheck"),
            "usage": agent.get("usage"),
        })
        transcript = Path(agent["transcriptPath"]) if agent.get("transcriptPath") else None
        if transcript and transcript.exists():
            shutil.copyfile(transcript, task_dir / "transcript.jsonl")
            errors = tool_error_summary(task_dir / "transcript.jsonl")
            record["tool_errors"] = errors["count"]
            record["tool_error_categories"] = errors["categories"]
            record["first_tool_error"] = errors["first"]
        started = time.perf_counter()
        record["stage"] = "save"
        try:
            workbook.Save()
            record["saved"] = True
        except Exception as exc:
            record["saved"] = False
            record["save_error"] = f"{type(exc).__name__}: {exc}"
        phase("save", started)
        started = time.perf_counter()
        record["stage"] = "close"
    phase("close", started)

    started = time.perf_counter()
    record["stage"] = "grade"
    try:
        passed, _ = compare(str(golden_file), str(workbook_path), task["instruction_type"], answer_position)
    except Exception as exc:
        passed = False
        record["grade_error"] = f"{type(exc).__name__}: {exc}"
    phase("grade", started)

    started = time.perf_counter()
    record["stage"] = "preservation"
    try:
        preservation = unauthorized_edits(init_file, workbook_path, answer_position, args.compare_values)
        # When the reference solution itself edits outside answer_position (e.g. the
        # instruction also asks to sort a source column), the check can't judge the task.
        gold = unauthorized_edits(init_file, golden_file, answer_position, args.compare_values, limit=0)
        preservation["gold_unauthorized_cells"] = gold["unauthorized_cells"]
        preservation["gold_sheets_removed"] = gold["sheets_removed"]
        preservation["gold_sheets_added"] = gold["sheets_added"]
    except Exception as exc:
        preservation = {"error": f"{type(exc).__name__}: {exc}"}
    phase("preservation", started)

    agent_status = agent.get("status")
    if agent_status in INFRA_STATUSES:
        infra = agent_status
    elif agent.get("stalled"):
        infra = "stalled"  # machine sleep or starvation, not model latency
    elif not record.get("saved"):
        infra = "save_error"
    elif "grade_error" in record:
        infra = "grade_error"
    else:
        infra = "ok"

    error_summary = tool_error_summary(task_dir / "transcript.jsonl")
    record.update(
        {
            "passed": bool(passed),
            "agent_status": agent_status,
            "infra_status": infra,
            "unauthorized_cells": preservation.get("unauthorized_cells"),
            "gold_unauthorized_cells": preservation.get("gold_unauthorized_cells"),
            "sheets_removed": preservation.get("sheets_removed", []),
            "sheets_added": preservation.get("sheets_added", []),
            "gold_sheets_removed": preservation.get("gold_sheets_removed", []),
            "gold_sheets_added": preservation.get("gold_sheets_added", []),
            "model": agent.get("model"),
            "session_id": agent.get("sessionId"),
            "tool_calls": len(agent.get("tools", [])),
            "tool_errors": error_summary["count"],
            "tool_error_categories": error_summary["categories"],
            "first_tool_error": error_summary["first"],
            "agent_duration_s": round((agent.get("monotonicMs") or 0) / 1000, 1),
            "agent_wall_s": round((agent.get("durationMs") or 0) / 1000, 1),
            "max_gap_s": round(max(agent.get("maxWallGapMs") or 0, agent.get("maxMonotonicGapMs") or 0) / 1000, 1),
            "usage": agent.get("usage"),
            "num_turns": agent.get("numTurns"),
            "error": agent.get("error"),
        }
    )
    record["failure_class"] = classify_outcome(infra, passed, agent_status)
    record["stage"] = "done"

    (task_dir / "attempt.json").write_text(
        json.dumps({**record, "preservation": preservation, "agent": agent, "prompt": prompt}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return record


def percentile(values: list[float], q: float) -> float | None:
    if not values:
        return None
    ordered = sorted(values)
    return round(ordered[min(len(ordered) - 1, int(q * len(ordered)))], 1)


def summarize(run: str, results: list[dict]) -> dict:
    n = len(results)
    ok = [r for r in results if r["infra_status"] == "ok"]
    durations = [r["agent_duration_s"] for r in ok]
    usage_keys = ("input_tokens", "cache_read_input_tokens", "cache_creation_input_tokens", "output_tokens")
    tokens = {k: sum((r.get("usage") or {}).get(k) or 0 for r in ok) for k in usage_keys}
    def damage_count(record: dict, prefix: str = "") -> int:
        return (
            int(record.get(f"{prefix}unauthorized_cells") or 0)
            + len(record.get(f"{prefix}sheets_removed") or [])
            + len(record.get(f"{prefix}sheets_added") or [])
        )

    checkable = [r for r in ok if r.get("unauthorized_cells") is not None]
    preserved = [r for r in checkable if damage_count(r, "gold_") == 0]
    tool_error_categories: dict[str, int] = {}
    for result in ok:
        for category, count in (result.get("tool_error_categories") or {}).items():
            tool_error_categories[category] = tool_error_categories.get(category, 0) + count
    return {
        "run": run,
        "attempted": n,
        # Infrastructure failures stay in the denominator of the headline rate.
        "end_to_end_pass_rate": round(sum(r["passed"] for r in results) / max(n, 1), 3),
        "infra_completion_rate": round(len(ok) / max(n, 1), 3),
        "pass_rate_given_infra_ok": round(sum(r["passed"] for r in ok) / max(len(ok), 1), 3),
        "by_type": {
            kind: {"attempted": len(g), "passed": sum(r["passed"] for r in g)}
            for kind in sorted({r["instruction_type"] for r in results})
            for g in [[r for r in results if r["instruction_type"] == kind]]
        },
        "infra_statuses": {s: sum(r["infra_status"] == s for r in results) for s in sorted({r["infra_status"] for r in results})},
        "agent_statuses": {s: sum(str(r["agent_status"]) == s for r in results) for s in sorted({str(r["agent_status"]) for r in results})},
        "failure_classes": {
            status: sum((r.get("failure_class") or "unknown") == status for r in results)
            for status in sorted({r.get("failure_class") or "unknown" for r in results})
        },
        "preservation": {
            "checked": len(preserved),
            "skipped_gold_edits_outside": len(checkable) - len(preserved),
            "tasks_with_unauthorized_edits": sum(1 for r in preserved if damage_count(r)),
            "unauthorized_cells_total": sum(r["unauthorized_cells"] for r in preserved),
            "unauthorized_sheet_changes_total": sum(
                len(r.get("sheets_removed") or []) + len(r.get("sheets_added") or []) for r in preserved
            ),
            "passed_but_damaged": sum(1 for r in preserved if r["passed"] and damage_count(r)),
        },
        "agent_seconds": {
            "median": round(statistics.median(durations), 1) if durations else None,
            # Tail latency is noise on small samples.
            "p90": percentile(durations, 0.9) if len(durations) >= 20 else None,
        },
        "tool_calls_mean": round(statistics.mean(r["tool_calls"] for r in ok), 1) if ok else None,
        "tool_errors_total": sum(r["tool_errors"] for r in ok),
        "tool_error_categories": tool_error_categories,
        "tokens_infra_ok": tokens,
        "tokens_per_task_mean": {k: round(v / max(len(ok), 1)) for k, v in tokens.items()},
    }


def keep_awake() -> int | None:
    """Hold off idle sleep for as long as this process runs.

    A 40-task run lost five tasks to idle sleep: twice before the AC timeout
    was turned off, and three times after, once the laptop was unplugged and
    the battery timeout applied. Power settings are the wrong lever; this is
    the documented Windows API media players use, it needs no settings change,
    and it lapses on its own when the process exits. Closing the lid still
    sleeps the machine.
    """
    if sys.platform != "win32":
        return None
    import ctypes

    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001
    call = ctypes.WinDLL("kernel32", use_last_error=True).SetThreadExecutionState
    call.argtypes = [ctypes.c_uint]
    call.restype = ctypes.c_uint
    previous_state = call(ES_CONTINUOUS | ES_SYSTEM_REQUIRED)
    if previous_state == 0:
        raise ctypes.WinError(ctypes.get_last_error())
    return previous_state


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--dataset", type=Path, required=True, help="spreadsheetbench_verified_400 directory")
    parser.add_argument("--spreadsheetbench", type=Path, default=PROJECT_ROOT.parent / "_sdks" / "spreadsheetbench",
                        help="SpreadsheetBench repo checkout (for evaluation/evaluation.py)")
    parser.add_argument("--run", required=True, help="run name; one directory per configuration")
    parser.add_argument("--model", default="sonnet", choices=["haiku", "sonnet", "opus"], help="model tier")
    parser.add_argument("--ids", nargs="*", help="run exactly these task ids")
    parser.add_argument("--sample", type=int, help="stratified random sample of this many tasks")
    parser.add_argument("--seed", type=int, default=0)
    parser.add_argument("--limit", type=int, default=10)
    parser.add_argument("--offset", type=int, default=0)
    parser.add_argument("--timeout", type=int, default=900, help="per-task agent timeout, seconds")
    parser.add_argument("--retry-infra", action="store_true", help="rerun tasks whose last attempt was an infra failure")
    args = parser.parse_args()
    previous_power_state = keep_awake()
    if previous_power_state is not None:
        print(f"Windows keep-awake request accepted; previous execution state=0x{previous_power_state:08x}", flush=True)

    sys.path.insert(0, str(args.spreadsheetbench / "evaluation"))
    from evaluation import compare_cell_value, cell_level_compare

    def compare_workbooks(*args):
        return compare_answer_workbooks(*args, cell_compare=cell_level_compare)

    args.compare_values = compare_cell_value
    tasks = json.loads((args.dataset / "dataset.json").read_text(encoding="utf-8"))
    selected = select_tasks(tasks, args)
    token = TOKEN_FILE.read_text(encoding="utf-8").strip()
    manifest = build_manifest(args, selected, daemon_request("/eval/info", token))
    manifest["runtime"] = {"keepAwakeRequestAccepted": previous_power_state is not None}

    run_dir = PROJECT_ROOT / "evals" / "runs" / args.run
    manifest_path = run_dir / "manifest.json"
    results_path = run_dir / "results.jsonl"
    if manifest_path.exists():
        existing = json.loads(manifest_path.read_text(encoding="utf-8"))
        if existing["config"] != manifest["config"]:
            changed = sorted(k for k in manifest["config"] if existing["config"].get(k) != manifest["config"][k])
            sys.exit(f"Run '{args.run}' was recorded with a different configuration ({', '.join(changed)}). "
                     "Use a new --run name.")
    elif results_path.exists():
        sys.exit(f"Run '{args.run}' predates run manifests; use a new --run name.")
    else:
        run_dir.mkdir(parents=True, exist_ok=True)
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        # The manifest only hashes uncommitted changes; keep the diff itself so
        # the run can be reproduced even if those changes are never committed.
        dirty_diff = uncommitted_patch()
        if dirty_diff:
            (run_dir / "uncommitted.patch").write_text(dirty_diff, encoding="utf-8")

    previous = {}
    if results_path.exists():
        for line in results_path.read_text(encoding="utf-8").splitlines():
            if line:
                row = json.loads(line)
                previous[row["id"]] = row  # the last attempt of each task wins
    addin = read_manifest(PROJECT_ROOT / "manifests" / "excel.xml")

    for index, task in enumerate(selected, 1):
        last = previous.get(str(task["id"]))
        if last and not (args.retry_infra and last["infra_status"] != "ok"):
            continue
        started = time.perf_counter()
        partial = {"id": str(task["id"]), "instruction_type": task["instruction_type"], "phases_s": {}}
        try:
            result = run_task(task, args.dataset, run_dir, args, token, addin, compare_workbooks, partial)
        except ExcelUnavailable as exc:
            # Stop without recording results; --retry-infra or a rerun picks these tasks up.
            print(f"[{index}/{len(selected)}] {task['id']}: run stopped — {exc}", flush=True)
            break
        except Exception as exc:
            # One broken task (COM error, unreadable workbook) must not end the run.
            result = harness_error_record(partial, exc)
            task_dir = run_dir / str(task["id"])
            task_dir.mkdir(parents=True, exist_ok=True)
            (task_dir / "attempt.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        with results_path.open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(result, ensure_ascii=False) + "\n")
        previous[result["id"]] = result
        print(
            f"[{index}/{len(selected)}] {result['id']}: {'PASS' if result['passed'] else 'FAIL'} "
            f"infra={result['infra_status']} agent={result['agent_status']} "
            f"outside_edits={result['unauthorized_cells']} tools={result['tool_calls']} "
            f"({time.perf_counter() - started:.0f}s)",
            flush=True,
        )

    results = [previous[str(t["id"])] for t in selected if str(t["id"]) in previous]
    summary = summarize(args.run, results)
    (run_dir / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
