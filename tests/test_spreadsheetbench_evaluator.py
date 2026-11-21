"""Unit tests for the SpreadsheetBench evaluator, a port of the official cell comparison."""

from __future__ import annotations

import datetime

import openpyxl
import pytest

from skillopt.envs.spreadsheetbench.evaluator import (
    _col_name2num,
    _col_num2name,
    _compare_cell_value,
    _generate_cell_names,
    _transform_value,
    compare_workbooks,
    evaluate,
)


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (1, 1.0),
        (1.239, 1.24),
        (1.234, 1.23),
        ("3.14159", 3.14),  # numeric strings are compared as numbers
        ("abc", "abc"),
        (True, 1.0),
        (None, None),
        (datetime.datetime(2024, 1, 2, 18, 0), 45294.0),  # Excel serial day, rounded
        (datetime.datetime(2024, 1, 2, 6, 0), 45293.0),
    ],
)
def test_transform_value(value, expected):
    assert _transform_value(value) == expected


@pytest.mark.parametrize(
    ("gold", "pred", "equal"),
    [
        (5, "5", True),
        (1.0, "1.00", True),
        (5, 5.004, True),  # 2-decimal quantization, not a tolerance
        (5, 5.006, False),
        (None, "", True),
        ("", None, True),
        (None, None, True),
        (None, 0, False),
        ("a", "A", False),  # text comparison is case-sensitive
        (1, "one", False),  # a number never equals non-numeric text
    ],
)
def test_compare_cell_value(gold, pred, equal):
    assert _compare_cell_value(gold, pred) is equal


@pytest.mark.parametrize(("name", "number"), [("A", 1), ("Z", 26), ("AA", 27), ("AZ", 52), ("AAA", 703)])
def test_column_name_round_trip(name, number):
    assert _col_name2num(name) == number
    assert _col_num2name(number) == name


@pytest.mark.parametrize(
    ("cell_range", "expected"),
    [
        ("A1", ["A1"]),
        ("A1:B2", ["A1", "A2", "B1", "B2"]),
        ("Y1:AB1", ["Y1", "Z1", "AA1", "AB1"]),
    ],
)
def test_generate_cell_names(cell_range, expected):
    assert _generate_cell_names(cell_range) == expected


def _workbook(path, sheets: dict[str, dict]) -> str:
    wb = openpyxl.Workbook()
    for index, (name, cells) in enumerate(sheets.items()):
        ws = wb.active if index == 0 else wb.create_sheet(name)
        ws.title = name
        for cell, value in cells.items():
            ws[cell] = value
    wb.save(path)
    return str(path)


@pytest.fixture
def gold(tmp_path) -> str:
    return _workbook(tmp_path / "gold.xlsx", {"Data": {"A1": 1, "B2": "x"}, "Out": {"C3": 2.5}})


def test_matching_workbook_passes_across_sheets_and_ranges(tmp_path, gold):
    pred = _workbook(tmp_path / "pred.xlsx", {"Data": {"A1": 1.0, "B2": "x"}, "Out": {"C3": "2.50"}})

    assert compare_workbooks(gold, pred, "Data!A1:B2,Out!C3") == (True, "")
    assert compare_workbooks(gold, pred, "'Out'!C3") == (True, "")
    # Without a sheet name the first gold sheet is used.
    assert compare_workbooks(gold, pred, "A1:B2") == (True, "")


def test_mismatch_reports_the_first_wrong_cell(tmp_path, gold):
    pred = _workbook(tmp_path / "pred.xlsx", {"Data": {"A1": 1, "B2": "y"}, "Out": {"C3": 2.5}})

    assert compare_workbooks(gold, pred, "Data!A1:B2,Out!C3") == (False, "value@Data!B2: gt='x' pred='y'")
    assert compare_workbooks(gold, pred, "Out!C3") == (True, "")


def test_every_listed_range_is_checked(tmp_path, gold):
    pred = _workbook(tmp_path / "pred.xlsx", {"Data": {"A1": 1, "B2": "x"}, "Out": {"C3": 9}})

    assert compare_workbooks(gold, pred, "Data!A1:B2,Out!C3") == (False, "value@Out!C3: gt=2.5 pred=9")


def test_missing_sheet_missing_file_and_unreadable_file_fail(tmp_path, gold):
    no_out_sheet = _workbook(tmp_path / "pred.xlsx", {"Data": {"A1": 1, "B2": "x"}})
    unreadable = tmp_path / "bad.xlsx"
    unreadable.write_text("not a workbook", encoding="utf-8")

    assert compare_workbooks(gold, no_out_sheet, "Out!C3") == (False, "worksheet not found: Out")
    assert compare_workbooks(gold, str(tmp_path / "missing.xlsx"), "A1") == (False, "file not exist")
    ok, reason = compare_workbooks(gold, str(unreadable), "A1")
    assert ok is False and reason.startswith("load error:")


def test_evaluate_returns_the_comparison_result(tmp_path, gold):
    pred = _workbook(tmp_path / "pred.xlsx", {"Data": {"A1": 1, "B2": "x"}, "Out": {"C3": 2.5}})

    assert evaluate(pred, gold, "Cell-Level Manipulation", "Data!A1") == {
        "ok": True,
        "reason": "",
        "instruction_type": "Cell-Level Manipulation",
    }
