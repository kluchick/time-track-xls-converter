# Excel Time Tracker Converter

Converts time tracking data from `input.xlsx` to a dated copy of `output.xlsx` on the Efforts sheet.

## Features

- Reads time tracking data from `input.xlsx`
- Converts Duration from minutes to hours
- Creates a dated copy of `output.xlsx` (format: `output_YYYY-MM-DD.xlsx`)
- Writes converted data to the "Efforts" sheet
- Preserves Excel formatting and structure

## Requirements

- Python 3.7+
- pandas >= 2.0.0
- openpyxl >= 3.1.0

## Installation

1. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Place your input file at `files/input.xlsx`
2. Ensure `files/output.xlsx` exists (template file)
3. Run the converter:
```bash
python src/converter.py
```

The script will create a new file `files/output_YYYY-MM-DD.xlsx` with converted data.

## Input File Format

The `input.xlsx` file must contain the following columns:
- **Project Task** - Task description
- **Date** - Date of the time entry
- **Duration** - Time spent in minutes (numeric)

## Output File Format

The script writes to the "Efforts" sheet in the output file with:
- **Effort** - Time spent in hours (Duration / 60)
- **Description** - Project Task value
- **Date** - Date value (same format as input)

## Behavior

- If the "Efforts" sheet doesn't exist, it will be created
- Existing data rows in the Efforts sheet are cleared (header row is preserved)
- Rows with invalid Duration values are skipped with a warning
- Empty input data creates an empty output file

## Error Handling

The script validates:
- Input file existence
- Required columns presence
- Data types (Duration must be numeric)
- Missing sheets (creates if needed)
