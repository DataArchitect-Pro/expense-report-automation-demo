"""Tests for the reusable, read-only workbook contract auditor."""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
from workbook_contract import ContractError, audit  # noqa: E402


def contract() -> dict:
    return {
        "schema_version": 1,
        "required_sheets": {"Summary": ["Metric", "Value"], "Raw_Data": ["Amount"]},
        "summary": {"sheet": "Summary", "key_column": "Metric", "value_column": "Value"},
        "forbid_formulas": True,
        "row_count_checks": [{"sheet": "Raw_Data", "summary_key": "Records"}],
        "sum_checks": [{"sheet": "Raw_Data", "column": "Amount", "summary_key": "Total", "tolerance": "0"}],
    }


def workbook(path: Path, *, count: int = 2, amount: float = 3.5, formula: bool = False) -> Path:
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    summary.append(["Metric", "Value"])
    summary.append(["Records", count])
    summary.append(["Total", amount + 1])
    raw = wb.create_sheet("Raw_Data")
    raw.append(["Amount", "Comment"])
    raw.append([amount, "=1+1" if formula else "plain text"])
    raw.append([1, "second"])
    wb.save(path)
    return path


def codes(result: dict) -> set[str]:
    return {item["code"] for item in result["findings"]}


def test_pass_and_artifact_hash(tmp_path: Path) -> None:
    result = audit(workbook(tmp_path / "good.xlsx"), contract())
    assert result["status"] == "PASS"
    assert len(result["artifact_sha256"]) == 64
    assert result["findings"] == []


def test_detect_count_mismatch(tmp_path: Path) -> None:
    result = audit(workbook(tmp_path / "bad.xlsx", count=3), contract())
    assert result["status"] == "FAIL"
    assert "COUNT_MISMATCH" in codes(result)


def test_detect_amount_mismatch(tmp_path: Path) -> None:
    result = audit(workbook(tmp_path / "bad.xlsx", amount=2.5), {
        **contract(), "sum_checks": [
            {"sheet": "Raw_Data", "column": "Amount", "summary_key": "Records"}
        ]
    })
    assert "SUM_MISMATCH" in codes(result)


def test_reject_formula(tmp_path: Path) -> None:
    result = audit(workbook(tmp_path / "formula.xlsx", formula=True), contract())
    assert "FORMULA_CELL" in codes(result)


def test_missing_sheet_reports_failure(tmp_path: Path) -> None:
    cfg = contract()
    cfg["required_sheets"]["DoesNotExist"] = ["Id"]
    assert "MISSING_SHEET" in codes(audit(workbook(tmp_path / "good.xlsx"), cfg))


def test_bad_contract_is_blocked(tmp_path: Path) -> None:
    from pytest import raises
    with raises(ContractError):
        audit(workbook(tmp_path / "good.xlsx"), {"schema_version": 2})


def test_cli_machine_readable_exit_codes(tmp_path: Path) -> None:
    good = workbook(tmp_path / "good.xlsx")
    rules = tmp_path / "contract.json"
    rules.write_text(json.dumps(contract()), encoding="utf-8")
    tool = ROOT / "tools" / "workbook_contract.py"
    args = [sys.executable, str(tool), str(good), "--contract", str(rules)]
    result = subprocess.run(args, capture_output=True, text=True, check=False)
    assert result.returncode == 0
    assert json.loads(result.stdout)["status"] == "PASS"
    rules.write_text("{}", encoding="utf-8")
    result = subprocess.run(args, capture_output=True, text=True, check=False)
    assert result.returncode == 2
    assert json.loads(result.stdout)["status"] == "BLOCKED"


def test_bundled_sample_passes_external_contract(tmp_path: Path) -> None:
    """Ensure the checked-in report contract matches the actual report generator."""
    sys.path.insert(0, str(ROOT / "src"))
    from generate_expense_report import run_report

    out = tmp_path / "output"
    assert run_report(ROOT / "input", out) == 0
    cfg = json.loads((ROOT / "contracts" / "expense-report.json").read_text(encoding="utf-8"))
    result = audit(out / "Expense_Report_2026-01.xlsx", cfg)
    assert result["status"] == "PASS", result["findings"]
