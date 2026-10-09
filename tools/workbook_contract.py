"""Independent, configuration-driven postcondition checks for XLSX reports.

This module reads workbooks only. It never evaluates formulas, modifies the
artifact, imports application code, or interprets input as Python expressions.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import sys
import zipfile
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


MAX_FINDINGS = 200
DEFAULT_MAX_BYTES = 50 * 1024 * 1024  # compressed size cap; larger inputs are BLOCKED
MAX_UNCOMPRESSED_BYTES = 20 * DEFAULT_MAX_BYTES  # zip-bomb guard on declared sizes

_TOP_KEYS = {"schema_version", "required_sheets", "summary", "forbid_formulas", "row_count_checks", "sum_checks"}
_SUMMARY_KEYS = {"sheet", "key_column", "value_column"}
_ROW_COUNT_KEYS = {"sheet", "summary_key"}
_SUM_KEYS = {"sheet", "column", "summary_key", "tolerance"}


class ContractError(ValueError):
    """The contract is invalid or cannot be applied."""


def _digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def _number(value: Any, *, allow_text: bool = False) -> Decimal:
    """Convert a numeric cell to Decimal. Text is accepted only for contract values."""
    if isinstance(value, bool) or value is None:
        raise ValueError("not a numeric cell")
    if not allow_text and not isinstance(value, (int, float, Decimal)):
        raise ValueError("not a numeric cell")
    try:
        result = Decimal(str(value))
    except (InvalidOperation, ValueError) as exc:
        raise ValueError("not a numeric cell") from exc
    if not result.is_finite():
        raise ValueError("non-finite numeric cell")
    return result


def _reject_unknown(obj: dict, allowed: set[str], where: str) -> None:
    """Unknown keys are errors so a typo cannot silently disable a check."""
    unknown = sorted(str(key) for key in obj if key not in allowed)
    if unknown:
        raise ContractError(f"{where}: unknown key(s): {', '.join(unknown)}")


def _validate_contract(contract: Any) -> None:
    if not isinstance(contract, dict):
        raise ContractError("contract must be a JSON object")
    version = contract.get("schema_version")
    if isinstance(version, bool) or not isinstance(version, int) or version != 1:
        raise ContractError("schema_version must be integer 1")
    _reject_unknown(contract, _TOP_KEYS, "contract")
    sheets = contract.get("required_sheets")
    if not isinstance(sheets, dict) or not sheets:
        raise ContractError("required_sheets must be a non-empty object")
    for name, headers in sheets.items():
        if not isinstance(name, str) or not name or not isinstance(headers, list) or not all(
            isinstance(h, str) and h for h in headers
        ) or len(headers) != len(set(headers)):
            raise ContractError("each sheet needs a list of distinct header strings")
    summary = contract.get("summary")
    if not isinstance(summary, dict) or not all(
        isinstance(summary.get(key), str) and summary[key] for key in ("sheet", "key_column", "value_column")
    ):
        raise ContractError("summary requires sheet, key_column and value_column")
    _reject_unknown(summary, _SUMMARY_KEYS, "summary")
    if summary["sheet"] not in sheets:
        raise ContractError("summary sheet must be in required_sheets")
    if not isinstance(contract.get("forbid_formulas", False), bool):
        raise ContractError("forbid_formulas must be boolean")
    for group, fields in (
        ("row_count_checks", ("sheet", "summary_key")),
        ("sum_checks", ("sheet", "column", "summary_key")),
    ):
        checks = contract.get(group, [])
        if not isinstance(checks, list):
            raise ContractError(f"{group} must be a list")
        for check in checks:
            if not isinstance(check, dict) or not all(isinstance(check.get(k), str) and check[k] for k in fields):
                raise ContractError(f"{group} entries require: {', '.join(fields)}")
            _reject_unknown(check, _SUM_KEYS if group == "sum_checks" else _ROW_COUNT_KEYS, group)
            if check["sheet"] not in sheets:
                raise ContractError(f"{group}: sheet not in required_sheets")
            if group == "sum_checks" and check["column"] not in sheets[check["sheet"]]:
                raise ContractError("sum_checks: column must be listed in required_sheets for that sheet")
            if group == "sum_checks":
                try:
                    tolerance = _number(check.get("tolerance", "0"), allow_text=True)
                except ValueError as exc:
                    raise ContractError("sum tolerance must be finite and numeric") from exc
                if tolerance < 0:
                    raise ContractError("sum tolerance cannot be negative")


def _check_size(path: Path, max_bytes: int) -> None:
    if path.stat().st_size > max_bytes:
        raise ContractError(f"workbook exceeds size limit of {max_bytes} bytes")
    try:
        with zipfile.ZipFile(path) as archive:
            if sum(info.file_size for info in archive.infolist()) > MAX_UNCOMPRESSED_BYTES:
                raise ContractError("workbook uncompressed size exceeds limit")
    except zipfile.BadZipFile as exc:
        raise ContractError("cannot open XLSX: BadZipFile") from exc


def audit(workbook_path: Path, contract: dict[str, Any], *, max_bytes: int = DEFAULT_MAX_BYTES) -> dict[str, Any]:
    """Return JSON-serializable findings. PASS does not imply business approval."""
    _validate_contract(contract)
    path = Path(workbook_path)
    if not path.is_file():
        raise ContractError("workbook does not exist")
    _check_size(path, max_bytes)
    artifact_hash = _digest(path)
    findings: list[dict[str, str]] = []
    state = {"dropped": 0}

    def fail(code: str, location: str, detail: str) -> None:
        if len(findings) >= MAX_FINDINGS:
            state["dropped"] += 1
            return
        findings.append({"code": code, "location": location, "detail": detail})

    try:
        workbook = load_workbook(path, read_only=True, data_only=False)
    except Exception as exc:
        raise ContractError(f"cannot open XLSX: {type(exc).__name__}") from exc
    try:
        return _run_checks(workbook, contract, fail, findings, state, artifact_hash)
    except ContractError:
        raise
    except Exception as exc:  # corrupt content discovered while streaming rows
        raise ContractError(f"cannot read XLSX content: {type(exc).__name__}") from exc
    finally:
        workbook.close()


def _run_checks(workbook: Any, contract: dict[str, Any], fail: Any, findings: list[dict[str, str]],
                state: dict[str, int], artifact_hash: str) -> dict[str, Any]:
    rows_by_sheet: dict[str, list[dict[str, Any]]] = {}
    for name, required_headers in contract["required_sheets"].items():
        if name not in workbook.sheetnames:
            fail("MISSING_SHEET", name, "required worksheet not found")
            continue
        sheet = workbook[name]
        all_rows = sheet.iter_rows(values_only=True)
        headers = list(next(all_rows, ()))
        if len(headers) != len(set(headers)):
            fail("DUPLICATE_HEADER", name, "duplicate header in first row")
        missing = [header for header in required_headers if header not in headers]
        for header in missing:
            fail("MISSING_COLUMN", name, header)
        if missing or len(headers) != len(set(headers)):
            continue
        rows_by_sheet[name] = [
            dict(zip(headers, values)) for values in all_rows
            if any(value is not None for value in values)
        ]

    if contract.get("forbid_formulas", False):
        for sheet in workbook.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.data_type == "f":
                        fail("FORMULA_CELL", f"{sheet.title}!{cell.coordinate}", "formula found")

    summary_conf = contract["summary"]
    summary_name = summary_conf["sheet"]
    key_col = summary_conf["key_column"]
    val_col = summary_conf["value_column"]
    metrics: dict[str, Any] = {}
    if summary_name in rows_by_sheet:
        for row in rows_by_sheet[summary_name]:
            key = row.get(key_col)
            if key is None:
                continue
            key = str(key)
            if key in metrics:
                fail("DUPLICATE_METRIC", summary_name, "duplicate summary key")
            metrics[key] = row.get(val_col)

    for rule in contract.get("row_count_checks", []):
        name, key = rule["sheet"], rule["summary_key"]
        if name not in rows_by_sheet:
            continue
        if key not in metrics:
            fail("MISSING_METRIC", summary_name, key)
            continue
        try:
            expected = _number(metrics[key])
        except ValueError:
            fail("INVALID_METRIC", summary_name, key)
            continue
        actual = len(rows_by_sheet[name])
        if expected != actual:
            fail("COUNT_MISMATCH", name, f"summary key {key!r}: expected {expected}, observed {actual}")

    for rule in contract.get("sum_checks", []):
        name, key, col = rule["sheet"], rule["summary_key"], rule["column"]
        if name not in rows_by_sheet:
            continue
        if key not in metrics:
            fail("MISSING_METRIC", summary_name, key)
            continue
        try:
            expected = _number(metrics[key])
            tolerance = _number(rule.get("tolerance", "0"), allow_text=True)
        except ValueError:
            fail("INVALID_METRIC", summary_name, key)
            continue
        total = Decimal("0")
        invalid = False
        for index, row in enumerate(rows_by_sheet[name], start=2):
            try:
                total += _number(row.get(col))
            except ValueError:
                fail("INVALID_SUM_CELL", f"{name}!{col}:{index}", "non-numeric value")
                invalid = True
        if not invalid and abs(total - expected) > tolerance:
            fail("SUM_MISMATCH", name, f"summary key {key!r}: expected {expected}, observed {total}")

    if state["dropped"]:
        findings.append({"code": "FINDINGS_TRUNCATED", "location": "report",
                         "detail": f"{state['dropped']} further finding(s) omitted"})
    return {
        "schema_version": 1,
        "status": "FAIL" if findings else "PASS",
        "artifact_sha256": artifact_hash,
        "findings": findings,
    }


def _load_contract(path: Path) -> Any:
    def no_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        keys = [key for key, _ in pairs]
        if len(keys) != len(set(keys)):
            raise ContractError("contract JSON contains duplicate keys")
        return dict(pairs)

    def no_constants(name: str) -> Any:
        raise ContractError("contract JSON contains a non-standard constant")

    try:
        with path.open("r", encoding="utf-8-sig") as source:
            return json.load(source, object_pairs_hook=no_duplicates, parse_constant=no_constants)
    except OSError as exc:
        raise ContractError(f"cannot read contract: {type(exc).__name__}") from exc
    except ValueError as exc:
        if isinstance(exc, ContractError):
            raise
        raise ContractError(f"contract is not valid UTF-8 JSON: {type(exc).__name__}") from exc


def _same_file(left: Path, right: Path) -> bool:
    return left.resolve() == right.resolve() or (
        left.exists() and right.exists() and left.samefile(right)
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Read-only contract audit of a saved XLSX report")
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--contract", required=True, type=Path)
    parser.add_argument("--output", type=Path, help="Optional JSON result path (never the workbook or contract)")
    parser.add_argument("--max-bytes", type=int, default=DEFAULT_MAX_BYTES,
                        help="Maximum workbook size in bytes (default: %(default)s)")
    args = parser.parse_args(argv)
    # Messages in BLOCKED reports are fixed strings: they never echo paths or cell contents.
    def blocked(detail: str) -> dict[str, Any]:
        return {"schema_version": 1, "status": "BLOCKED", "findings": [
            {"code": "INPUT_ERROR", "location": "contract or workbook", "detail": detail}
        ]}

    write_output = args.output is not None
    try:
        if args.max_bytes <= 0:
            raise ContractError("--max-bytes must be positive")
        if args.output and (_same_file(args.output, args.workbook) or _same_file(args.output, args.contract)):
            write_output = False  # never overwrite an input, even to report the refusal
            raise ContractError("--output must not be the workbook or contract file")
        report = audit(args.workbook, _load_contract(args.contract), max_bytes=args.max_bytes)
    except ContractError as exc:
        report = blocked(str(exc))
    except OSError as exc:
        report = blocked(f"I/O error: {type(exc).__name__}")
    if write_output:
        try:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(json.dumps(report, indent=2, sort_keys=True, ensure_ascii=False) + "\n",
                                   encoding="utf-8")
        except OSError as exc:
            report = blocked(f"cannot write --output: {type(exc).__name__}")
    exit_code = {"PASS": 0, "FAIL": 1}.get(report["status"], 2)
    print(json.dumps(report, indent=2, sort_keys=True, ensure_ascii=False))
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
