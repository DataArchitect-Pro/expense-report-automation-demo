# Reusable XLSX contract auditor (CANDIDATE)

`tools/workbook_contract.py` is a standalone, **read-only** postcondition verifier for XLSX files, independent of the code that generated them. Rules are supplied in a JSON file, so the same checker can be reused with other spreadsheet reports and synthetic samples without copying application-specific Python tests.

```bash
python tools/workbook_contract.py output/Expense_Report_2026-01.xlsx --contract contracts/expense-report.json --output output/contract-audit.json
```

Options: `--max-bytes N` (default 52428800; larger workbooks are BLOCKED) and `--output PATH` (JSON copy of the result; refused if it is the workbook or contract file).

Exit status: `0` PASS; `1` FAIL (artifact violates the contract); `2` BLOCKED (invalid contract, missing/unreadable input). Results contain artifact SHA256 and deterministic findings; they are **not business approval**, and do not attest that the source data are complete or permitted. PASS only asserts the checks named in the contract.

Contract schema v1 fields:

- `required_sheets`: mapping of required worksheet names to required first-row column names.
- `summary`: sheet and columns holding metric-name / metric-value pairs.
- `forbid_formulas`: reject any formula cells; safe for reports expected to be data-only.
- `row_count_checks`: compare non-empty row counts in a sheet to a summary metric.
- `sum_checks`: compare the numeric sum of a named column to a summary metric with optional non-negative tolerance (decimal string is recommended).

Hardening behaviour (v1):

- Unknown contract keys, boolean/float `schema_version`, duplicate JSON keys, and `sum_checks.column` values not listed in `required_sheets` are contract errors (BLOCKED), so a typo cannot silently disable a check.
- Numeric checks accept only real numeric cells; text that merely looks numeric (e.g. `"3.5"`), booleans and NaN/Infinity are `INVALID_SUM_CELL`.
- Corrupt or oversized workbooks, unreadable files and unwritable `--output` give BLOCKED (exit 2), never a traceback with exit 1. BLOCKED messages are fixed strings and do not echo file paths or cell contents.
- At most 200 findings are listed; extra findings are summarised by one `FINDINGS_TRUNCATED` entry. Output uses sorted keys; findings follow contract order, so the same inputs give the same JSON.

The tool does not execute workbook formulas, access the network, change files, or make judgments about source-data authorization. It does not replace tests for raw-row identity, source-data provenance, business rules, or Windows/Excel compatibility. Do not run this tool with sensitive files in unapproved environments. See `tests/test_workbook_contract.py` for isolated positive and negative tests.

**Change status:** CANDIDATE / NON_AUTHORITATIVE. Draft PR for Human review; no automated merge or release.
