# Reusable XLSX contract auditor (CANDIDATE)

`tools/workbook_contract.py` is a standalone, **read-only** postcondition verifier for XLSX files, independent of the code that generated them. Rules are supplied in a JSON file, so the same checker can be reused with other spreadsheet reports and synthetic samples without copying application-specific Python tests.

```bash
python tools/workbook_contract.py output/Expense_Report_2026-01.xlsx --contract contracts/expense-report.json --output output/contract-audit.json
```

Exit status: `0` PASS; `1` FAIL (artifact violates the contract); `2` BLOCKED (invalid contract, missing/unreadable input). Results contain artifact SHA256 and deterministic findings; they are **not business approval**, and do not attest that the source data are complete or permitted. PASS only asserts the checks named in the contract.

Contract schema v1 fields:

- `required_sheets`: mapping of required worksheet names to required first-row column names.
- `summary`: sheet and columns holding metric-name / metric-value pairs.
- `forbid_formulas`: reject any formula cells; safe for reports expected to be data-only.
- `row_count_checks`: compare non-empty row counts in a sheet to a summary metric.
- `sum_checks`: compare the numeric sum of a named column to a summary metric with optional non-negative tolerance (decimal string is recommended).

The tool does not execute workbook formulas, access the network, change files, or make judgments about source-data authorization. It does not replace tests for raw-row identity, source-data provenance, business rules, or Windows/Excel compatibility. Do not run this tool with sensitive files in unapproved environments. See `tests/test_workbook_contract.py` for isolated positive and negative tests.

**Change status:** CANDIDATE / NON_AUTHORITATIVE. Draft PR for Human review; no automated merge or release.
