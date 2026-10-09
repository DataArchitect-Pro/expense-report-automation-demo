from __future__ import annotations

import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "input"
DEFAULT_MONTH = "2026-01"
DEFAULT_OFFICES = ["Singapore", "Tokyo", "London"]

DEPARTMENTS = ["IT", "Admin", "Sales", "Operations", "HR"]
EXPENSE_TYPES = ["Travel", "Meal", "Software", "Office Supplies", "Training", "Telecom", "Other"]
EMPLOYEES = [
    "Alex Tan",
    "Maya Chen",
    "Noah Smith",
    "Priya Kumar",
    "Ethan Wong",
    "Hana Sato",
    "Liam Brown",
    "Sara Lee",
]
DESCRIPTIONS = [
    "Client visit",
    "Team lunch",
    "Monthly software license",
    "Printer paper and folders",
    "Skills workshop",
    "Mobile plan",
    "Project supplies",
]


def office_token(office: str) -> str:
    return office.lower().replace(" ", "_")


def random_day(month: str) -> str:
    first_day = date.fromisoformat(f"{month}-01")
    day_offset = random.randint(0, 30)
    return (first_day + timedelta(days=day_offset)).isoformat()


def valid_row(office: str, month: str) -> dict[str, object]:
    amount = round(random.uniform(12, 950), 2)
    amount_value: object = amount
    if random.random() < 0.2:
        amount_value = f"${amount:,.2f}"
    return {
        "Date": random_day(month),
        "Office": random.choice([office, office.lower(), f" {office} "]),
        "Department": random.choice(DEPARTMENTS),
        "Employee": random.choice(EMPLOYEES),
        "ExpenseType": random.choice(EXPENSE_TYPES),
        "Amount": amount_value,
        "Currency": random.choice(["USD", "usd", " USD "]),
        "Description": random.choice(DESCRIPTIONS),
    }


def error_rows() -> dict[str, list[dict[str, object]]]:
    base = {
        "Date": "2026-01-15",
        "Office": "Singapore",
        "Department": "IT",
        "Employee": "Alex Tan",
        "ExpenseType": "Meal",
        "Amount": "45.25",
        "Currency": "USD",
        "Description": "Team lunch",
    }
    rows = [
        {**base, "Date": ""},
        {**base, "Date": "01/15/2026"},
        {**base, "Office": ""},
        {**base, "Office": "Paris"},
        {**base, "Office": "London"},
        {**base, "Department": ""},
        {**base, "Department": "Legal"},
        {**base, "Employee": ""},
        {**base, "ExpenseType": ""},
        {**base, "ExpenseType": "Parking"},
        {**base, "Amount": "USD 120.50"},
        {**base, "Amount": "0"},
        {**base, "Currency": ""},
        {**base, "Currency": "EUR"},
        {**base, "Description": ""},
        {
            **base,
            "Date": "",
            "Department": "",
            "Amount": "not available",
            "Currency": "EUR",
            "Description": "",
        },
    ]
    return {"Singapore": rows}


def generate_sample_data(month: str = DEFAULT_MONTH, offices: list[str] | None = None, output_dir: Path = DEFAULT_OUTPUT_DIR) -> None:
    random.seed(42)
    offices = offices or DEFAULT_OFFICES
    output_dir.mkdir(parents=True, exist_ok=True)

    deliberate_errors = error_rows()
    rows_per_office = 50
    for office in offices:
        rows = [valid_row(office, month) for _ in range(rows_per_office)]
        rows.extend(deliberate_errors.get(office, []))
        frame = pd.DataFrame(rows)
        path = output_dir / f"expense_{office_token(office)}_{month}.csv"
        frame.to_csv(path, index=False, encoding="utf-8-sig")
        print(f"Wrote {path}")


if __name__ == "__main__":
    generate_sample_data()
