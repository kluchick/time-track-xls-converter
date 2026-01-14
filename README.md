# Excel Time Tracker Converter

A user-friendly desktop application for converting time tracking data from Excel files. Features a graphical interface with drag-and-drop support for easy file conversion.

## Features

- **Graphical User Interface** - Modern, intuitive UI built with Flet
- **Drag-and-Drop Support** - Simply drag your input Excel file to convert
- **Automatic Conversion** - Converts Duration from minutes to hours
- **Timestamped Output** - Creates output files with date and time (format: `output_YYYY-MM-DD_HH-MM-SS.xlsx`)
- **Smart Validation** - Validates input files before conversion
- **Output Location** - Saves converted files in the same directory as the application
- **Template Embedded** - No external template files needed
- **Cross-Platform** - Works on Windows, macOS, and Linux

## Requirements

- Python 3.7+
- pandas >= 2.0.0
- openpyxl >= 3.1.0
- flet >= 0.24.0

## Installation

1. Clone or download this repository
2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Running the Application

Simply run the UI application:
```bash
python src/ui_app.py
```

### Using the Interface

1. **Select File**:
   - Drag and drop your input Excel file onto the window, OR
   - Click "Browse Files" to select your file
2. **Convert**: Click the "Convert File" button
3. **View Output**: The converted file will be saved in the application directory with a timestamp
4. **Open Folder**: Click "Open Output Folder" to view the converted file

### Output Location

The converted file will be created in the **same directory where you run the application** (project root), not in the `files/` folder.

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

The application validates:
- Input file existence
- Required columns presence
- Data types (Duration must be numeric)
- File format (must be .xlsx or .xls)

Error messages are displayed in the UI with clear explanations of what went wrong.

## Architecture

The application is modular with three main components:

- **`src/ui_app.py`** - Flet-based graphical user interface
- **`src/converter_core.py`** - Core conversion business logic
- **`src/template_data.py`** - Embedded Excel template (base64 encoded)

This separation allows for easy testing and potential CLI/API implementations in the future.

## Development

To run in development mode:

```bash
# Activate virtual environment (if using venv)
source venv/bin/activate  # On Linux/macOS
venv\Scripts\activate     # On Windows

# Run the application
python src/ui_app.py
```

## Notes

- The Excel template is embedded in the application, so no external template file is needed
- The original CLI version (`converter.py`) has been replaced with the UI version
- Output files include both date and time in the filename for better uniqueness
- The application runs in a single window and is fully self-contained