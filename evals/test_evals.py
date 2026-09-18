"""Offline checks for the eval harness: run with
uv run --project evals python -m unittest discover -s evals
"""

import tempfile
import unittest
from argparse import Namespace
from pathlib import Path

import openpyxl

from run_spreadsheetbench import qualified_answer_position, select_tasks, summarize
from workbook_diff import authorized_ranges, unauthorized_edits


def same_value(a, b):
    return a == b


class WorkbookDiffTest(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp())

    def save(self, name, cells, sheets=("Data",)):
        wb = openpyxl.Workbook()
        wb.active.title = sheets[0]
        for extra in sheets[1:]:
            wb.create_sheet(extra)
        for (sheet, coord), value in cells.items():
            wb[sheet][coord] = value
        path = self.dir / name
        wb.save(path)
        return path

    def test_edits_inside_authorized_range_are_ignored(self):
        before = self.save("a.xlsx", {("Data", "A1"): 1, ("Data", "B2"): None})
        after = self.save("b.xlsx", {("Data", "A1"): 1, ("Data", "B2"): "=A1*2"})
        self.assertEqual(unauthorized_edits(before, after, "B2", same_value)["unauthorized_cells"], 0)

    def test_value_and_formula_changes_outside_range_are_counted(self):
        before = self.save("a.xlsx", {("Data", "A1"): 1, ("Data", "C3"): "=A1+1"})
        after = self.save("b.xlsx", {("Data", "A1"): 5, ("Data", "C3"): 2, ("Data", "D9"): "x"})
        diff = unauthorized_edits(before, after, "B2", same_value)
        self.assertEqual(diff["unauthorized_cells"], 3)
        self.assertEqual({s["cell"] for s in diff["samples"]}, {"'Data'!A1", "'Data'!C3", "'Data'!D9"})

    def test_ranges_are_scoped_to_their_sheet(self):
        before = self.save("a.xlsx", {("Data", "A1"): 1, ("Out", "A1"): None}, sheets=("Data", "Out"))
        after = self.save("b.xlsx", {("Data", "A1"): 1, ("Out", "A1"): 9}, sheets=("Data", "Out"))
        self.assertEqual(unauthorized_edits(before, after, "'Out'!A1:B2", same_value)["unauthorized_cells"], 0)
        # Unqualified positions refer to the first sheet, as in the official grader.
        self.assertEqual(unauthorized_edits(before, after, "A1:B2", same_value)["unauthorized_cells"], 1)

    def test_added_and_removed_sheets_are_reported(self):
        before = self.save("a.xlsx", {("Data", "A1"): 1}, sheets=("Data", "Old"))
        after = self.save("b.xlsx", {("Data", "A1"): 1}, sheets=("Data", "New"))
        diff = unauthorized_edits(before, after, "A1", same_value)
        self.assertEqual((diff["sheets_removed"], diff["sheets_added"]), (["Old"], ["New"]))

    def test_authorized_ranges_parse_multi_part_positions(self):
        boxes = authorized_ranges("'My Sheet'!A1:B2,C5", "First")
        self.assertEqual(boxes, {"My Sheet": [(1, 1, 2, 2)], "First": [(5, 3, 5, 3)]})


class SelectionTest(unittest.TestCase):
    tasks = [{"id": i, "instruction_type": "Cell" if i % 4 else "Sheet"} for i in range(400)]

    def test_stratified_sample_is_deterministic_and_keeps_the_mix(self):
        args = Namespace(ids=None, sample=40, seed=7, offset=0, limit=10)
        first, second = select_tasks(self.tasks, args), select_tasks(self.tasks, args)
        self.assertEqual(first, second)
        self.assertEqual(len(first), 40)
        self.assertEqual(sum(t["instruction_type"] == "Sheet" for t in first), 10)

    def test_explicit_ids_keep_the_requested_order(self):
        args = Namespace(ids=["5", "2"], sample=None, seed=0, offset=0, limit=10)
        self.assertEqual([t["id"] for t in select_tasks(self.tasks, args)], [5, 2])

    def test_answer_sheet_qualifies_every_part(self):
        task = {"answer_position": "A1:B2,D4", "answer_sheet": "Out"}
        self.assertEqual(qualified_answer_position(task), "'Out'!A1:B2,'Out'!D4")


class SummaryTest(unittest.TestCase):
    def test_infra_failures_stay_in_the_headline_denominator(self):
        ok = {"instruction_type": "Cell", "infra_status": "ok", "agent_status": "completed", "tool_calls": 3,
              "tool_errors": 0, "agent_duration_s": 10, "unauthorized_cells": 0}
        results = [
            {**ok, "passed": True},
            {**ok, "passed": True, "unauthorized_cells": 2},
            {**ok, "passed": False, "infra_status": "stalled"},
            {**ok, "passed": False},
        ]
        summary = summarize("t", results)
        self.assertEqual(summary["end_to_end_pass_rate"], 0.5)
        self.assertEqual(summary["infra_completion_rate"], 0.75)
        self.assertEqual(summary["pass_rate_given_infra_ok"], 0.667)
        self.assertEqual(summary["preservation"]["passed_but_damaged"], 1)
        self.assertIsNone(summary["agent_seconds"]["p90"])


if __name__ == "__main__":
    unittest.main()
