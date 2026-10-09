from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from generate_expense_report import aggregate, build_summary, validate_and_normalize_row


def base_row() -> dict[str, object]:
    return {
        "Date": "2026-01-15",
        "Office": " singapore ",
        "Department": "IT",
        "Employee": "Alex Tan",
        "ExpenseType": "Meal",
        "Amount": "$1,200.50",
        "Currency": "usd",
        "Description": "Team lunch",
    }


def error_text(row: dict[str, object], expected_office: str = "Singapore") -> str:
    return str(validate_and_normalize_row(row, expected_office)["ErrorType"])


def test_valid_row() -> None:
    result = validate_and_normalize_row(base_row(), "Singapore")

    assert result["IsValid"] is True
    assert result["Office"] == "Singapore"
    assert result["Currency"] == "USD"
    assert result["Amount"] == 1200.50


def test_invalid_office() -> None:
    row = {**base_row(), "Office": "Paris"}

    assert "Invalid Office" in error_text(row)


def test_invalid_department() -> None:
    row = {**base_row(), "Department": "Legal"}

    assert "Invalid Department" in error_text(row)


def test_invalid_amount() -> None:
    row = {**base_row(), "Amount": "USD 120.50"}

    assert "Invalid Amount" in error_text(row)


def test_non_positive_amount() -> None:
    row = {**base_row(), "Amount": "-1"}

    assert "Non-positive Amount" in error_text(row)


def test_unsupported_currency() -> None:
    row = {**base_row(), "Currency": "EUR"}

    assert "Unsupported Currency" in error_text(row)


def test_multiple_errors() -> None:
    row = {**base_row(), "Department": "", "Amount": "none", "Currency": "EUR"}

    assert error_text(row) == "Missing Department | Invalid Amount | Unsupported Currency"


def test_aggregation_reconciliation() -> None:
    valid_rows = pd.DataFrame(
        [
            {**base_row(), "Office": "Singapore", "Department": "IT", "ExpenseType": "Meal", "Amount": 100.0},
            {**base_row(), "Office": "Tokyo", "Department": "Admin", "ExpenseType": "Travel", "Amount": 250.0},
            {**base_row(), "Office": "Tokyo", "Department": "IT", "ExpenseType": "Meal", "Amount": 50.0},
        ]
    )
    by_office = aggregate(valid_rows, "Office")
    by_department = aggregate(valid_rows, "Department")
    by_expense_type = aggregate(valid_rows, "ExpenseType")
    summary = build_summary(
        "2026-01",
        files_processed=3,
        total_records=3,
        valid_records=3,
        error_records=0,
        total_amount=float(valid_rows["Amount"].sum()),
        by_office=by_office,
        by_department=by_department,
        by_expense_type=by_expense_type,
        generated_at="2026-01-31 12:00:00",
    )
    summary_total = float(summary.loc[summary["Metric"] == "Total Amount", "Value"].iloc[0])

    assert round(by_office["TotalAmount"].sum(), 2) == round(summary_total, 2)
    assert round(by_department["TotalAmount"].sum(), 2) == round(summary_total, 2)
    assert round(by_expense_type["TotalAmount"].sum(), 2) == round(summary_total, 2)
