from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT_DIR = PROJECT_ROOT / "input"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "output"

REQUIRED_COLUMNS = [
    "Date",
    "Office",
    "Department",
    "Employee",
    "ExpenseType",
    "Amount",
    "Currency",
    "Description",
]

OFFICE_BY_TOKEN = {
    "singapore": "Singapore",
    "tokyo": "Tokyo",
    "london": "London",
    "new_york": "New York",
}

ALLOWED_OFFICES = {
    "singapore": "Singapore",
    "tokyo": "Tokyo",
    "london": "London",
    "new york": "New York",
    "new_york": "New York",
}

ALLOWED_DEPARTMENTS = {
    "it": "IT",
    "admin": "Admin",
    ("fin" + "ance"): "Fin" + "ance",
    "sales": "Sales",
    "operations": "Operations",
    "hr": "HR",
}

ALLOWED_EXPENSE_TYPES = {
    "travel": "Travel",
    "meal": "Meal",
    "software": "Software",
    "office supplies": "Office Supplies",
    "training": "Training",
    "telecom": "Telecom",
    "other": "Other",
}

FILENAME_RE = re.compile(r"^expense_(singapore|tokyo|london|new_york)_(\d{4}-\d{2})\.csv$", re.I)
ERROR_SEPARATOR = " | "
ORIGINAL_PREFIX = "Original_"
ORIGINAL_COLUMNS = [f"{ORIGINAL_PREFIX}{column}" for column in REQUIRED_COLUMNS]
AMOUNT_HEADERS = {"Amount", "TotalAmount"}


@dataclass(frozen=True)
class ParsedExpenseFile:
    path: Path
    office: str
    report_month: str


def parse_expense_filename(path: Path) -> ParsedExpenseFile | None:
    match = FILENAME_RE.match(path.name)
    if not match:
        return None
    office_token = match.group(1).lower()
    return ParsedExpenseFile(path=path, office=OFFICE_BY_TOKEN[office_token], report_month=match.group(2))


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_lookup(value: object, allowed_values: dict[str, str]) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    return allowed_values.get(text.lower(), text)


def normalize_currency(value: object) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    return "USD" if text.lower() == "usd" else text


def parse_strict_date(value: object) -> tuple[str, bool, bool]:
    text = normalize_text(value)
    if not text:
        return "", True, False
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", text):
        return text, False, False
    try:
        datetime.strptime(text, "%Y-%m-%d")
    except ValueError:
        return text, False, False
    return text, False, True


def parse_amount(value: object) -> tuple[float | None, bool]:
    text = normalize_text(value)
    if not text:
        return None, False
    if text.startswith("(") or ")" in text:
        return None, False
    if text.startswith("$"):
        text = text[1:].strip()
    text = text.replace(",", "")
    if not re.fullmatch(r"-?\d+(\.\d+)?", text):
        return None, False
    try:
        return float(text), True
    except ValueError:
        return None, False


def raw_text(value: object) -> str:
    """Return the input value as-is (no strip/normalize) for the audit trail."""
    if value is None or pd.isna(value):
        return ""
    return str(value)


def validate_and_normalize_row(
    row: pd.Series | dict[str, object],
    expected_office: str,
    expected_month: str | None = None,
) -> dict[str, object]:
    """Validate one row.

    Normalized values (Date, Office, ..., Amount) are used for aggregation.
    Original_* keys keep the untouched input text so a reviewer can see why a
    row failed. When ``expected_month`` (YYYY-MM) is given, a valid Date outside
    that month is reported as "Mismatched Report Month".
    """
    source = dict(row)
    errors: list[str] = []

    date_value, missing_date, valid_date = parse_strict_date(source.get("Date"))
    if missing_date:
        errors.append("Missing Date")
    elif not valid_date:
        errors.append("Invalid Date")
    elif expected_month is not None and date_value[:7] != expected_month:
        errors.append("Mismatched Report Month")

    office_raw = normalize_text(source.get("Office"))
    office = normalize_lookup(office_raw, ALLOWED_OFFICES)
    if not office_raw:
        errors.append("Missing Office")
    elif office not in ALLOWED_OFFICES.values():
        errors.append("Invalid Office")
    elif office != expected_office:
        errors.append("Mismatched Office")

    department_raw = normalize_text(source.get("Department"))
    department = normalize_lookup(department_raw, ALLOWED_DEPARTMENTS)
    if not department_raw:
        errors.append("Missing Department")
    elif department not in ALLOWED_DEPARTMENTS.values():
        errors.append("Invalid Department")

    employee = normalize_text(source.get("Employee"))
    if not employee:
        errors.append("Missing Employee")

    expense_type_raw = normalize_text(source.get("ExpenseType"))
    expense_type = normalize_lookup(expense_type_raw, ALLOWED_EXPENSE_TYPES)
    if not expense_type_raw:
        errors.append("Missing ExpenseType")
    elif expense_type not in ALLOWED_EXPENSE_TYPES.values():
        errors.append("Invalid ExpenseType")

    amount, amount_valid = parse_amount(source.get("Amount"))
    if not amount_valid:
        errors.append("Invalid Amount")
    elif amount is not None and amount <= 0:
        errors.append("Non-positive Amount")

    currency_raw = normalize_text(source.get("Currency"))
    currency = normalize_currency(currency_raw)
    if not currency_raw:
        errors.append("Missing Currency")
    elif currency != "USD":
        errors.append("Unsupported Currency")

    description = normalize_text(source.get("Description"))
    if not description:
        errors.append("Missing Description")

    is_valid = not errors
    return {
        "Date": date_value,
        "Office": office,
        "Department": department,
        "Employee": employee,
        "ExpenseType": expense_type,
        "Amount": amount if amount_valid else None,
        "Currency": currency,
        "Description": description,
        "IsValid": is_valid,
        "ErrorType": ERROR_SEPARATOR.join(errors),
        **{f"{ORIGINAL_PREFIX}{column}": raw_text(source.get(column)) for column in REQUIRED_COLUMNS},
    }


def discover_csv_files(input_dir: Path) -> list[Path]:
    if not input_dir.exists():
        return []
    return sorted(path for path in input_dir.iterdir() if path.is_file() and path.suffix.lower() == ".csv")


def read_valid_files(csv_files: Iterable[Path]) -> tuple[list[ParsedExpenseFile], list[str]]:
    parsed_files: list[ParsedExpenseFile] = []
    messages: list[str] = []
    for path in csv_files:
        parsed = parse_expense_filename(path)
        if parsed is None:
            messages.append(f"Skipped {path.name}: invalid file name.")
            continue
        parsed_files.append(parsed)
    return parsed_files, messages


def build_raw_data(parsed_files: Iterable[ParsedExpenseFile]) -> tuple[pd.DataFrame, list[str], int]:
    rows: list[dict[str, object]] = []
    messages: list[str] = []
    files_processed = 0

    for parsed in parsed_files:
        try:
            frame = pd.read_csv(parsed.path, dtype=str, encoding="utf-8-sig", keep_default_na=False)
        except Exception as exc:
            messages.append(f"Skipped {parsed.path.name}: could not read CSV ({exc}).")
            continue

        missing_columns = [column for column in REQUIRED_COLUMNS if column not in frame.columns]
        if missing_columns:
            messages.append(f"Skipped {parsed.path.name}: missing required columns: {', '.join(missing_columns)}.")
            continue

        files_processed += 1
        for index, row in frame.iterrows():
            normalized = validate_and_normalize_row(row, parsed.office, parsed.report_month)
            normalized["SourceFile"] = parsed.path.name
            normalized["RowNumber"] = int(index) + 2
            rows.append(normalized)

    columns = ["SourceFile", "RowNumber", *REQUIRED_COLUMNS, "IsValid", "ErrorType", *ORIGINAL_COLUMNS]
    return pd.DataFrame(rows, columns=columns), messages, files_processed


def aggregate(valid_rows: pd.DataFrame, group_column: str) -> pd.DataFrame:
    if valid_rows.empty:
        return pd.DataFrame(columns=[group_column, "TotalAmount", "RecordCount"])
    grouped = (
        valid_rows.groupby(group_column, as_index=False)
        .agg(TotalAmount=("Amount", "sum"), RecordCount=("Amount", "size"))
        .sort_values([group_column], kind="stable")
        .reset_index(drop=True)
    )
    grouped["TotalAmount"] = grouped["TotalAmount"].round(2)
    return grouped


def top_by_amount(aggregation: pd.DataFrame, label_column: str) -> str:
    if aggregation.empty:
        return "N/A"
    ordered = aggregation.sort_values(["TotalAmount", label_column], ascending=[False, True], kind="mergesort")
    return str(ordered.iloc[0][label_column])


def build_summary(
    report_month: str,
    files_processed: int,
    total_records: int,
    valid_records: int,
    error_records: int,
    total_amount: float,
    by_office: pd.DataFrame,
    by_department: pd.DataFrame,
    by_expense_type: pd.DataFrame,
    generated_at: str,
) -> pd.DataFrame:
    return pd.DataFrame(
        [
            ["Report Month", report_month],
            ["Total Files Processed", files_processed],
            ["Total Records", total_records],
            ["Valid Records", valid_records],
            ["Error Records", error_records],
            ["Total Amount", round(total_amount, 2)],
            ["Top Office", top_by_amount(by_office, "Office")],
            ["Top Department", top_by_amount(by_department, "Department")],
            ["Top Expense Type", top_by_amount(by_expense_type, "ExpenseType")],
            ["Generated At", generated_at],
        ],
        columns=["Metric", "Value"],
    )


def build_run_log_frame(log_entries: list[tuple[str, object]]) -> pd.DataFrame:
    return pd.DataFrame(log_entries, columns=["Item", "Value"])


def write_process_log(output_dir: Path, log_entries: list[tuple[str, object]]) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    log_path = output_dir / "process_log.txt"
    lines = [f"{key}: {value}" for key, value in log_entries]
    log_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def format_workbook(path: Path) -> None:
    workbook = load_workbook(path)
    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        for cell in sheet[1]:
            cell.font = Font(bold=True)
        for column_cells in sheet.columns:
            column_letter = get_column_letter(column_cells[0].column)
            max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
            sheet.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 45)
        for column_letter in amount_columns(sheet):
            for cell in sheet[column_letter]:
                cell.number_format = "#,##0.00"
    workbook.save(path)


def amount_columns(sheet) -> set[str]:
    headers = {cell.value: cell.column_letter for cell in sheet[1]}
    return {letter for header, letter in headers.items() if header in AMOUNT_HEADERS}


def force_text_cells(worksheet) -> None:
    """Keep CSV-derived text as text: openpyxl turns any string starting with "=" into a formula."""
    for sheet_row in worksheet.iter_rows():
        for cell in sheet_row:
            if cell.data_type == "f":
                cell.data_type = "s"


def write_excel_report(
    output_path: Path,
    summary: pd.DataFrame,
    by_office: pd.DataFrame,
    by_department: pd.DataFrame,
    by_expense_type: pd.DataFrame,
    raw_data: pd.DataFrame,
    error_report: pd.DataFrame,
    run_log: pd.DataFrame,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        by_office.to_excel(writer, sheet_name="By_Office", index=False)
        by_department.to_excel(writer, sheet_name="By_Department", index=False)
        by_expense_type.to_excel(writer, sheet_name="By_ExpenseType", index=False)
        raw_data.to_excel(writer, sheet_name="Raw_Data", index=False)
        error_report.to_excel(writer, sheet_name="Error_Report", index=False)
        run_log.to_excel(writer, sheet_name="Run_Log", index=False)
        for worksheet in writer.book.worksheets:
            force_text_cells(worksheet)
    format_workbook(output_path)


def run_report(input_dir: Path = DEFAULT_INPUT_DIR, output_dir: Path = DEFAULT_OUTPUT_DIR) -> int:
    started_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)
    csv_files = discover_csv_files(input_dir)
    warnings: list[str] = []

    if not csv_files:
        completed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entries = [
            ("Started At", started_at),
            ("Input Folder", str(input_dir)),
            ("Output File", "N/A"),
            ("Files Processed", 0),
            ("Total Records", 0),
            ("Valid Records", 0),
            ("Error Records", 0),
            ("Completed At", completed_at),
            ("Status", "No CSV files found. No Excel report created."),
            ("File-level warnings/errors", "None"),
            ("Existing output overwritten", "No"),
        ]
        write_process_log(output_dir, log_entries)
        return 0

    parsed_files, filename_messages = read_valid_files(csv_files)
    warnings.extend(filename_messages)
    months = sorted({parsed.report_month for parsed in parsed_files})

    if len(months) > 1:
        completed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entries = [
            ("Started At", started_at),
            ("Input Folder", str(input_dir)),
            ("Output File", "N/A"),
            ("Files Processed", 0),
            ("Total Records", 0),
            ("Valid Records", 0),
            ("Error Records", 0),
            ("Completed At", completed_at),
            ("Status", f"Aborted: multiple report months found ({', '.join(months)})."),
            ("File-level warnings/errors", " | ".join(warnings) if warnings else "None"),
            ("Existing output overwritten", "No"),
        ]
        write_process_log(output_dir, log_entries)
        return 1

    report_month = months[0] if months else "unknown"
    output_path = output_dir / f"Expense_Report_{report_month}.xlsx" if months else output_dir / "Expense_Report_unknown.xlsx"
    overwrite = output_path.exists() and bool(months)
    if overwrite:
        warnings.append(f"Existing output file overwritten: {output_path.name}.")

    raw_data, file_messages, files_processed = build_raw_data(parsed_files)
    warnings.extend(file_messages)

    valid_rows = raw_data[raw_data["IsValid"] == True].copy() if not raw_data.empty else raw_data.copy()
    error_report_columns = ["SourceFile", "RowNumber", "ErrorType", *REQUIRED_COLUMNS, *ORIGINAL_COLUMNS]
    error_report = (
        raw_data[raw_data["IsValid"] == False][error_report_columns].copy()
        if not raw_data.empty
        else pd.DataFrame(columns=error_report_columns)
    )

    by_office = aggregate(valid_rows, "Office")
    by_department = aggregate(valid_rows, "Department")
    by_expense_type = aggregate(valid_rows, "ExpenseType")

    total_records = len(raw_data)
    valid_records = len(valid_rows)
    error_records = len(error_report)
    total_amount = round(float(valid_rows["Amount"].sum()), 2) if valid_records else 0.0
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    summary = build_summary(
        report_month,
        files_processed,
        total_records,
        valid_records,
        error_records,
        total_amount,
        by_office,
        by_department,
        by_expense_type,
        generated_at,
    )

    completed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    status = "Completed" if months else "No valid input files found. Excel report created with empty sheets."
    log_entries = [
        ("Started At", started_at),
        ("Input Folder", str(input_dir)),
        ("Output File", str(output_path) if months else "N/A"),
        ("Files Processed", files_processed),
        ("Total Records", total_records),
        ("Valid Records", valid_records),
        ("Error Records", error_records),
        ("Completed At", completed_at),
        ("Status", status),
        ("File-level warnings/errors", " | ".join(warnings) if warnings else "None"),
        ("Existing output overwritten", "Yes" if overwrite else "No"),
    ]

    if months:
        run_log = build_run_log_frame(log_entries)
        write_excel_report(output_path, summary, by_office, by_department, by_expense_type, raw_data, error_report, run_log)
    write_process_log(output_dir, log_entries)
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate a monthly expense report from CSV files.")
    parser.add_argument("--input-dir", type=Path, default=DEFAULT_INPUT_DIR)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    args = parser.parse_args()
    return run_report(args.input_dir, args.output_dir)


if __name__ == "__main__":
    raise SystemExit(main())
