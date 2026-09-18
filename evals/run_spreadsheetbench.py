"""Run SpreadsheetBench tasks through Excel Assistant in the live Excel instance.

For each task: copy the initial workbook into its own folder (fresh agent
session), tag it so the task pane auto-opens, open it in Excel, inject the
prompt through the daemon's /eval/run endpoint, save, and grade with the
official SpreadsheetBench comparison.

  uv run --project evals python evals/run_spreadsheetbench.py \
      --dataset <.../spreadsheetbench_verified_400> --run qwen-plus --limit 20
"""

import argparse
import json
import shutil
import sys
import time
import urllib.request
from pathlib import Path

from embed_taskpane import embed_taskpane, read_manifest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DAEMON = "http://127.0.0.1:47834"
TOKEN_FILE = Path.home() / ".claude" / "office-addins" / "bridge-token"

# Field descriptions follow SpreadsheetBench's PROMPT_NO_DF_RCT_FORMAT, adapted
# from "write Python to an output file" to editing the open workbook in place.
PROMPT = """You need to solve the given spreadsheet manipulation question, which contains three types of information:
- instruction: The question about spreadsheet manipulation.
- instruction_type: There are two values (Cell-Level Manipulation, Sheet-Level Manipulation) used to indicate whether the answer to this question applies only to specific cells or to the entire worksheet.
- answer_position: The position need to be modified or filled. For Cell-Level Manipulation questions, this field is filled with the cell position; for Sheet-Level Manipulation, it is the maximum range of cells you need to modify. You only need to modify or fill in values within the cell range specified by answer_position.

The workbook is already open in Excel. Make the changes directly in it; do not create other files. Do not ask clarifying questions — complete the task.

### instruction
{instruction}

### instruction_type
{instruction_type}

### answer_position
{answer_position}
"""


def qualified_answer_position(task: dict) -> str:
    position = task["answer_position"]
    sheet = task.get("answer_sheet")
    if not sheet or "!" in position:
        return position
    return ",".join(f"'{sheet}'!{part.strip()}" for part in position.split(","))


def post_json(path: str, payload: dict, token: str, timeout: float) -> dict:
    request = urllib.request.Request(
        DAEMON + path,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json", "x-bridge-token": token},
        method="POST",
    )
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    with opener.open(request, timeout=timeout) as response:
        return json.load(response)


def excel_app():
    import win32com.client

    try:
        app = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        app = win32com.client.Dispatch("Excel.Application")
    app.Visible = True
    return app


def run_task(task: dict, dataset: Path, run_dir: Path, args, token: str, addin: tuple[str, str], compare) -> dict:
    task_id = str(task["id"])
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

    answer_position = qualified_answer_position(task)
    prompt = PROMPT.format(
        instruction=task["instruction"],
        instruction_type=task["instruction_type"],
        answer_position=answer_position,
    )

    app = excel_app()
    app.DisplayAlerts = False
    workbook = app.Workbooks.Open(str(workbook_path), UpdateLinks=0)
    try:
        agent = post_json(
            "/eval/run",
            {"doc": str(workbook_path), "prompt": prompt, "model": args.model, "timeoutMs": args.timeout * 1000},
            token,
            timeout=args.timeout + 180,
        )
        workbook.Save()
    finally:
        try:
            workbook.Close(SaveChanges=False)
        finally:
            app.DisplayAlerts = True

    try:
        passed, _ = compare(str(golden_file), str(workbook_path), task["instruction_type"], answer_position)
    except Exception as exc:
        passed = False
        agent["grade_error"] = str(exc)

    return {
        "id": task_id,
        "instruction_type": task["instruction_type"],
        "passed": bool(passed),
        "status": agent.get("status"),
        "model": agent.get("model"),
        "tool_calls": len(agent.get("tools", [])),
        "duration_s": round(agent.get("durationMs", 0) / 1000, 1),
        "usage": agent.get("usage"),
        "error": agent.get("error") or agent.get("grade_error"),
        "final_text": agent.get("text", "")[-2000:],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--dataset", type=Path, required=True, help="spreadsheetbench_verified_400 directory")
    parser.add_argument("--spreadsheetbench", type=Path, default=PROJECT_ROOT.parent / "_sdks" / "spreadsheetbench",
                        help="SpreadsheetBench repo checkout (for evaluation/evaluation.py)")
    parser.add_argument("--run", required=True, help="run name, e.g. qwen3.7-plus")
    parser.add_argument("--model", default="sonnet", choices=["haiku", "sonnet", "opus"], help="model tier")
    parser.add_argument("--limit", type=int, default=10)
    parser.add_argument("--offset", type=int, default=0)
    parser.add_argument("--ids", nargs="*", help="run only these task ids")
    parser.add_argument("--timeout", type=int, default=900, help="per-task agent timeout, seconds")
    args = parser.parse_args()

    sys.path.insert(0, str(args.spreadsheetbench / "evaluation"))
    from evaluation import compare_workbooks

    tasks = json.loads((args.dataset / "dataset.json").read_text(encoding="utf-8"))
    for task in tasks:
        task["id"] = str(task["id"])
    tasks = [t for t in tasks if t["id"] in set(args.ids)] if args.ids else tasks[args.offset: args.offset + args.limit]

    run_dir = PROJECT_ROOT / "evals" / "runs" / args.run
    run_dir.mkdir(parents=True, exist_ok=True)
    results_path = run_dir / "results.jsonl"
    done = set()
    if results_path.exists():
        done = {str(json.loads(line)["id"]) for line in results_path.read_text(encoding="utf-8").splitlines() if line}

    token = TOKEN_FILE.read_text(encoding="utf-8").strip()
    addin = read_manifest(PROJECT_ROOT / "manifests" / "excel.xml")

    for index, task in enumerate(tasks, 1):
        if task["id"] in done:
            continue
        started = time.time()
        try:
            result = run_task(task, args.dataset, run_dir, args, token, addin, compare_workbooks)
        except Exception as exc:
            # One broken task (COM error, unreadable workbook) must not end the run.
            result = {
                "id": task["id"],
                "instruction_type": task["instruction_type"],
                "passed": False,
                "status": "harness_error",
                "model": None,
                "tool_calls": 0,
                "duration_s": round(time.time() - started, 1),
                "usage": None,
                "error": f"{type(exc).__name__}: {exc}",
                "final_text": "",
            }
        with results_path.open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(result, ensure_ascii=False) + "\n")
        print(f"[{index}/{len(tasks)}] {task['id']}: {'PASS' if result['passed'] else 'FAIL'} "
              f"({result['status']}, {result['tool_calls']} tools, {time.time() - started:.0f}s)", flush=True)

    results = [json.loads(line) for line in results_path.read_text(encoding="utf-8").splitlines() if line]
    summary = {
        "run": args.run,
        "models": sorted({r["model"] for r in results if r.get("model")}),
        "tasks": len(results),
        "passed": sum(r["passed"] for r in results),
        "pass_rate": round(sum(r["passed"] for r in results) / max(len(results), 1), 3),
        "by_type": {
            kind: {
                "tasks": len(group),
                "passed": sum(r["passed"] for r in group),
            }
            for kind in sorted({r["instruction_type"] for r in results})
            for group in [[r for r in results if r["instruction_type"] == kind]]
        },
        "statuses": {s: sum(r["status"] == s for r in results) for s in sorted({r["status"] for r in results})},
        "avg_duration_s": round(sum(r["duration_s"] for r in results) / max(len(results), 1), 1),
    }
    (run_dir / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
