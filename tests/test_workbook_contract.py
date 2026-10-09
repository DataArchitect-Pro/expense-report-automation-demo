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


# --- Hardening tests (negative fixtures) -----------------------------------
import zipfile  # noqa: E402

import pytest  # noqa: E402
from workbook_contract import MAX_FINDINGS  # noqa: E402

TOOL = ROOT / "tools" / "workbook_contract.py"


def run_cli(*args: object) -> subprocess.CompletedProcess:
    return subprocess.run([sys.executable, str(TOOL), *map(str, args)], capture_output=True, text=True, check=False)


def write_contract(path: Path, cfg: dict | None = None) -> Path:
    path.write_text(json.dumps(cfg if cfg is not None else contract()), encoding="utf-8")
    return path


@pytest.mark.parametrize("bad_version", [True, 1.0, "1"])
def test_schema_version_must_be_integer_one(tmp_path: Path, bad_version: object) -> None:
    with pytest.raises(ContractError):
        audit(workbook(tmp_path / "g.xlsx"), {**contract(), "schema_version": bad_version})


def test_unknown_contract_keys_are_rejected(tmp_path: Path) -> None:
    """A typo such as 'forbid_formula' must not silently disable a check."""
    cfg = {**contract(), "forbid_formula": True}
    with pytest.raises(ContractError):
        audit(workbook(tmp_path / "g.xlsx"), cfg)
    cfg = contract()
    cfg["sum_checks"][0]["tolerence"] = "0.01"
    with pytest.raises(ContractError):
        audit(workbook(tmp_path / "g.xlsx"), cfg)


def test_check_column_must_be_declared_in_required_sheets(tmp_path: Path) -> None:
    cfg = contract()
    cfg["sum_checks"][0]["column"] = "NotDeclared"
    with pytest.raises(ContractError):
        audit(workbook(tmp_path / "g.xlsx"), cfg)


def test_text_typed_number_is_not_a_numeric_cell(tmp_path: Path) -> None:
    path = workbook(tmp_path / "t.xlsx")
    from openpyxl import load_workbook
    wb = load_workbook(path)
    wb["Raw_Data"]["A2"] = "3.5"  # text that merely looks numeric
    wb.save(path)
    result = audit(path, contract())
    assert "INVALID_SUM_CELL" in codes(result)


def test_non_finite_and_boolean_cells_are_invalid(tmp_path: Path) -> None:
    path = workbook(tmp_path / "t.xlsx")
    from openpyxl import load_workbook
    wb = load_workbook(path)
    wb["Raw_Data"]["A2"] = True
    wb["Raw_Data"]["A3"] = float("inf")
    wb.save(path)
    assert "INVALID_SUM_CELL" in codes(audit(path, contract()))


def test_duplicate_metric_with_mixed_key_types(tmp_path: Path) -> None:
    path = workbook(tmp_path / "d.xlsx")
    from openpyxl import load_workbook
    wb = load_workbook(path)
    wb["Summary"].append(["Records", 2])
    wb.save(path)
    assert "DUPLICATE_METRIC" in codes(audit(path, contract()))


def test_findings_are_capped_and_flagged(tmp_path: Path) -> None:
    from openpyxl import load_workbook
    path = workbook(tmp_path / "many.xlsx")
    wb = load_workbook(path)
    raw = wb["Raw_Data"]
    for _ in range(MAX_FINDINGS + 50):
        raw.append(["not-a-number", "x"])
    wb.save(path)
    result = audit(path, {**contract(), "row_count_checks": []})
    assert result["status"] == "FAIL"
    assert len(result["findings"]) <= MAX_FINDINGS + 1
    assert result["findings"][-1]["code"] == "FINDINGS_TRUNCATED"


def test_oversized_workbook_is_blocked(tmp_path: Path) -> None:
    good = workbook(tmp_path / "g.xlsx")
    with pytest.raises(ContractError):
        audit(good, contract(), max_bytes=100)
    proc = run_cli(good, "--contract", write_contract(tmp_path / "c.json"), "--max-bytes", "100")
    assert proc.returncode == 2 and json.loads(proc.stdout)["status"] == "BLOCKED"


def test_corrupt_sheet_xml_is_blocked_not_a_crash(tmp_path: Path) -> None:
    good = workbook(tmp_path / "g.xlsx")
    bad = tmp_path / "bad.xlsx"
    with zipfile.ZipFile(good) as zin, zipfile.ZipFile(bad, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "xl/worksheets/sheet2.xml":
                data = data[: len(data) // 2]
            zout.writestr(info, data)
    proc = run_cli(bad, "--contract", write_contract(tmp_path / "c.json"))
    assert proc.returncode == 2, proc.stderr
    assert json.loads(proc.stdout)["status"] == "BLOCKED"
    assert "Traceback" not in proc.stderr


def test_non_xlsx_and_directory_inputs_are_blocked(tmp_path: Path) -> None:
    rules = write_contract(tmp_path / "c.json")
    junk = tmp_path / "junk.xlsx"
    junk.write_text("not a zip", encoding="utf-8")
    assert run_cli(junk, "--contract", rules).returncode == 2
    assert run_cli(tmp_path, "--contract", rules).returncode == 2


def test_output_may_not_overwrite_inputs(tmp_path: Path) -> None:
    good = workbook(tmp_path / "g.xlsx")
    rules = write_contract(tmp_path / "c.json")
    before = good.read_bytes()
    proc = run_cli(good, "--contract", rules, "--output", good)
    assert proc.returncode == 2
    assert good.read_bytes() == before
    contract_before = rules.read_bytes()
    assert run_cli(good, "--contract", rules, "--output", rules).returncode == 2
    assert rules.read_bytes() == contract_before


def test_blocked_report_does_not_leak_paths(tmp_path: Path) -> None:
    secret_dir = tmp_path / "confidential-customer-dir"
    secret_dir.mkdir()
    missing = secret_dir / "payroll-secret.xlsx"
    proc = run_cli(missing, "--contract", tmp_path / "no-such-contract-name.json")
    assert proc.returncode == 2
    for needle in ("confidential-customer-dir", "payroll-secret", "no-such-contract-name"):
        assert needle not in proc.stdout and needle not in proc.stderr


def test_duplicate_json_keys_and_bom_contract(tmp_path: Path) -> None:
    good = workbook(tmp_path / "g.xlsx")
    dup = tmp_path / "dup.json"
    dup.write_text('{"schema_version": 1, "schema_version": 1}', encoding="utf-8")
    assert run_cli(good, "--contract", dup).returncode == 2
    bom = tmp_path / "bom.json"
    bom.write_bytes(b"\xef\xbb\xbf" + json.dumps(contract()).encode("utf-8"))
    assert run_cli(good, "--contract", bom).returncode == 0


def test_output_is_deterministic(tmp_path: Path) -> None:
    good = workbook(tmp_path / "g.xlsx", count=9)
    rules = write_contract(tmp_path / "c.json")
    first, second = run_cli(good, "--contract", rules), run_cli(good, "--contract", rules)
    assert first.returncode == second.returncode == 1
    assert first.stdout == second.stdout


def test_exit_code_fail_is_one(tmp_path: Path) -> None:
    proc = run_cli(workbook(tmp_path / "g.xlsx", count=9), "--contract", write_contract(tmp_path / "c.json"))
    assert proc.returncode == 1 and json.loads(proc.stdout)["status"] == "FAIL"
