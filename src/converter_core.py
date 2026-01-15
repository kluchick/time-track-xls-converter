"""
Excel Time Tracker Converter - Core Business Logic

Core functionality for converting time tracking Excel files.
Separated from UI layer for better modularity and testability.
"""

from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from typing import Any, Optional, Tuple

import openpyxl
import pandas as pd

# Column mappings
INPUT_COLUMNS = ["Project Task", "Date", "Duration"]
OUTPUT_COLUMNS = ["Effort", "Description", "Date"]
EFFORTS_SHEET_NAME = "Efforts"
VALID_EXCEL_EXTENSIONS = {".xlsx", ".xls"}
TARGET_DAILY_HOURS = 8.0
DATA_START_ROW = 3


class ConversionError(Exception):
    """Custom exception for conversion errors."""


def normalize_date_to_date_object(date_value: Any) -> Optional[date]:
    """Convert various date formats to a Python date object."""
    if date_value is None:
        return None

    if hasattr(date_value, "date"):
        return date_value.date()

    if hasattr(date_value, "year"):
        return date_value

    try:
        return pd.to_datetime(date_value).date()
    except (ValueError, TypeError):
        return None


def validate_input_data(file_path: Path) -> Tuple[bool, str]:
    """
    Validate input file exists and has required columns.

    Returns (True, "") if valid, or (False, error_message) if invalid.
    """
    if not file_path.exists():
        return False, f"File not found: {file_path}"

    if not file_path.is_file():
        return False, f"Path is not a file: {file_path}"

    if file_path.suffix.lower() not in VALID_EXCEL_EXTENSIONS:
        return False, "File must be an Excel file (.xlsx or .xls)"

    try:
        df = pd.read_excel(file_path, nrows=0)
        missing_cols = [col for col in INPUT_COLUMNS if col not in df.columns]
        if missing_cols:
            return False, f"Missing required columns: {', '.join(missing_cols)}"
        return True, ""
    except Exception as e:
        return False, f"Error reading file: {e}"


def read_and_process_input(
    file_path: Path, extend_description: Optional[str] = None
) -> pd.DataFrame:
    """
    Read input Excel and process data.

    Returns processed DataFrame with columns: Effort, Description, Date.
    If extend_description is provided, adds filler rows for days with < 8 hours.
    """
    try:
        df = pd.read_excel(file_path)

        missing_cols = [col for col in INPUT_COLUMNS if col not in df.columns]
        if missing_cols:
            raise ConversionError(f"Missing required columns: {', '.join(missing_cols)}")

        df = df[INPUT_COLUMNS].copy()
        df = _clean_duration_column(df)
        df = df.dropna(subset=["Project Task", "Date", "Duration"])

        if df.empty:
            raise ConversionError("No valid data found in input file")

        processed = pd.DataFrame({
            "Effort": df["Duration"] / 60.0,
            "Description": df["Project Task"],
            "Date": pd.to_datetime(df["Date"]).dt.normalize(),
        })

        if extend_description and extend_description.strip():
            processed = extend_time_entries(processed, extend_description.strip())

        print(f"Processed {len(processed)} rows from input file")
        return processed

    except ConversionError:
        raise
    except Exception as error:
        raise ConversionError(f"Error processing input file: {error}") from error


def _clean_duration_column(df: pd.DataFrame) -> pd.DataFrame:
    """Convert Duration column to numeric and remove invalid rows."""
    df["Duration"] = pd.to_numeric(df["Duration"], errors="coerce")
    invalid_count = df["Duration"].isna().sum()
    if invalid_count > 0:
        print(f"Warning: Found {invalid_count} rows with invalid Duration values. They will be skipped.")
        df = df[df["Duration"].notna()].copy()
    return df


def extend_time_entries(processed_df: pd.DataFrame, extend_description: str) -> pd.DataFrame:
    """
    Extend time entries to ensure each day has at least 8 hours.

    Adds filler rows with extend_description for days with < 8 hours total.
    """
    processed_df = processed_df.sort_values("Date").reset_index(drop=True)
    daily_totals = processed_df.groupby("Date")["Effort"].sum()
    days_to_extend = daily_totals[daily_totals < TARGET_DAILY_HOURS]

    if days_to_extend.empty:
        return processed_df

    extended_rows = []
    current_date = None

    for _, row in processed_df.iterrows():
        row_date = row["Date"]

        if current_date is not None and current_date != row_date:
            filler = _create_filler_row(current_date, daily_totals, days_to_extend, extend_description)
            if filler:
                extended_rows.append(filler)

        extended_rows.append(row.to_dict())
        current_date = row_date

    # Handle the last day
    filler = _create_filler_row(current_date, daily_totals, days_to_extend, extend_description)
    if filler:
        extended_rows.append(filler)

    print(f"Extended {len(days_to_extend)} days to 8 hours with description: '{extend_description}'")
    return pd.DataFrame(extended_rows)


def _create_filler_row(
    date_value: Any,
    daily_totals: pd.Series,
    days_to_extend: pd.Series,
    description: str,
) -> Optional[dict]:
    """Create a filler row for a day that needs extension, or return None."""
    if date_value is None or date_value not in days_to_extend.index:
        return None

    missing_hours = TARGET_DAILY_HOURS - daily_totals[date_value]
    return {"Effort": missing_hours, "Description": description, "Date": date_value}


def write_to_output_file(template_bytes: bytes, data_df: pd.DataFrame, output_path: Path) -> int:
    """
    Load template from bytes, write data to Efforts sheet, save to output_path.

    Returns the number of duplicate rows that were skipped.
    """
    try:
        wb = openpyxl.load_workbook(BytesIO(template_bytes))
        ws = _get_or_create_efforts_sheet(wb)

        last_data_row = _find_last_data_row(ws)
        existing_rows = _read_existing_rows(ws, last_data_row)
        data_df, duplicates_count = _filter_duplicates(data_df, existing_rows)

        write_start_row = last_data_row + 1
        column_a_value = _get_column_a_value(ws)

        _write_data_rows(ws, data_df, write_start_row, column_a_value)
        wb.save(output_path)

        _log_write_results(len(data_df), duplicates_count, write_start_row, column_a_value)
        return duplicates_count

    except Exception as error:
        raise ConversionError(f"Error writing output file: {error}") from error


def _get_or_create_efforts_sheet(wb: openpyxl.Workbook):
    """Get the Efforts sheet, creating it if necessary."""
    if EFFORTS_SHEET_NAME not in wb.sheetnames:
        print(f"Created '{EFFORTS_SHEET_NAME}' sheet")
        return wb.create_sheet(EFFORTS_SHEET_NAME)
    return wb[EFFORTS_SHEET_NAME]


def _find_last_data_row(ws) -> int:
    """Find the last row with data in columns B, C, or D."""
    last_data_row = DATA_START_ROW - 1

    for row in range(DATA_START_ROW, ws.max_row + 1):
        has_data = any(
            ws.cell(row, col).value is not None and str(ws.cell(row, col).value).strip()
            for col in range(2, 5)
        )
        if has_data:
            last_data_row = row

    return last_data_row


def _read_existing_rows(ws, last_data_row: int) -> set:
    """Read existing rows from the worksheet for duplicate detection."""
    existing_rows = set()

    for row in range(DATA_START_ROW, last_data_row + 1):
        effort_val = _parse_float(ws.cell(row, 2).value)
        desc_val = _normalize_string(ws.cell(row, 3).value)
        date_val = normalize_date_to_date_object(ws.cell(row, 4).value)

        if effort_val is not None and desc_val and date_val is not None:
            existing_rows.add((effort_val, desc_val, date_val))

    return existing_rows


def _parse_float(value: Any) -> Optional[float]:
    """Safely parse a value to float."""
    if value is None:
        return None
    try:
        return float(value)
    except (ValueError, TypeError):
        return None


def _normalize_string(value: Any) -> Optional[str]:
    """Normalize a value to a stripped string or None."""
    if value is None:
        return None
    return str(value).strip() or None


def _filter_duplicates(
    data_df: pd.DataFrame, existing_rows: set
) -> Tuple[pd.DataFrame, int]:
    """Filter out duplicate rows from the data. Returns (filtered_df, duplicate_count)."""
    duplicates_count = 0
    unique_rows = []

    for _, row_data in data_df.iterrows():
        effort = float(row_data["Effort"])
        desc = str(row_data["Description"]).strip()
        date_normalized = normalize_date_to_date_object(row_data["Date"])

        row_tuple = (effort, desc, date_normalized) if date_normalized else None

        if row_tuple and row_tuple in existing_rows:
            duplicates_count += 1
        else:
            unique_rows.append(row_data)
            if row_tuple:
                existing_rows.add(row_tuple)

    if unique_rows:
        return pd.DataFrame(unique_rows).reset_index(drop=True), duplicates_count
    return pd.DataFrame(columns=OUTPUT_COLUMNS), duplicates_count


def _get_column_a_value(ws) -> Optional[str]:
    """Get the column A value from the first data row if it exists."""
    value = ws.cell(DATA_START_ROW, 1).value
    if value is not None and str(value).strip():
        return value
    return None


def _write_data_rows(
    ws, data_df: pd.DataFrame, start_row: int, column_a_value: Optional[str]
) -> None:
    """Write data rows to the worksheet."""
    for row_idx, (_, row_data) in enumerate(data_df.iterrows(), start=start_row):
        if column_a_value is not None:
            ws.cell(row_idx, 1).value = column_a_value

        for col_idx, col_name in enumerate(OUTPUT_COLUMNS, start=2):
            value = row_data[col_name]
            cell = ws.cell(row_idx, col_idx)
            cell.value = value

            if col_name == "Date" and isinstance(value, pd.Timestamp):
                cell.number_format = "YYYY-MM-DD"


def _log_write_results(
    rows_written: int, duplicates_count: int, start_row: int, column_a_value: Optional[str]
) -> None:
    """Log the results of the write operation."""
    if rows_written == 0:
        if duplicates_count > 0:
            print(f"No new rows added - all {duplicates_count} row(s) were duplicates")
        else:
            print("No rows to write")
        return

    columns = "A, B, C, D" if column_a_value else "B, C, D"
    print(f"Appended {rows_written} rows to '{EFFORTS_SHEET_NAME}' sheet starting from row {start_row} (columns {columns})")

    if column_a_value:
        print(f"  - Copied column A value '{column_a_value}' to all imported rows")

    if duplicates_count > 0:
        print(f"  - Skipped {duplicates_count} duplicate row(s)")


def convert_excel_file(
    input_path: Path,
    template_bytes: bytes,
    output_dir: Path,
    extend_description: Optional[str] = None,
) -> Tuple[Path, int]:
    """
    Main conversion function. Converts input Excel to output with timestamp.

    Returns (output_path, duplicate_count). If extend_description is provided,
    adds filler rows for days with < 8 hours.
    """
    try:
        is_valid, error_msg = validate_input_data(input_path)
        if not is_valid:
            raise ConversionError(error_msg)

        processed_df = read_and_process_input(input_path, extend_description=extend_description)

        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        output_path = output_dir / f"ets_time_records_{timestamp}.xlsx"
        output_dir.mkdir(parents=True, exist_ok=True)

        duplicates_count = write_to_output_file(template_bytes, processed_df, output_path)

        print("Conversion completed successfully!")
        print(f"Output file: {output_path}")

        return output_path, duplicates_count

    except ConversionError:
        raise
    except Exception as error:
        raise ConversionError(f"Unexpected error during conversion: {error}") from error
