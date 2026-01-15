# Excel Time Tracker Converter

Desktop application for converting time tracking data from Excel files. Simple graphical interface with drag-and-drop file support.

## Installation

The **dist** folder contains the Windows executable file (`ExcelTimeTrackerConverter-x.x.exe`).

1. Copy the `ExcelTimeTrackerConverter-x.x.exe` file from the **dist** folder to any convenient location on your computer
2. Double-click the file to launch it

No additional installation is required - all necessary components are already included in the executable file.

### Running from Source (Python)

If you prefer to run the application from Python source code:

**Requirements:**
- Python 3.7 or higher

**Installation:**

1. Create a virtual environment (recommended):
   ```bash
   python -m venv venv
   ```

2. Activate the virtual environment:
   - On Windows: `venv\Scripts\activate`
   - On Linux/macOS: `source venv/bin/activate`

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

**Running the Application:**

```bash
python src/ui_app.py
```

The application will work the same way as the executable version.

## Usage

### Launching the Application

Double-click the `ExcelTimeTrackerConverter-1.0.exe` file

### Using the Application

1. **Select File**:
   - Drag and drop your Excel file directly into the application window, OR
   - Click "Choose File..." to select a file via dialog, OR
   - Enter the file path manually in the text field

2. **Convert**: Click the "Convert File" button

3. **Result**: The converted file will be saved in the same folder where the application is located, with a timestamp in the filename

4. **Open Folder**: Click "Open Output Folder" to view the converted file

### Output File Format

Filename: `ets_time_records_YYYY-MM-DD_HH-MM-SS.xlsx` (e.g., `ets_time_records_2026-01-13_14-30-45.xlsx`)

## Input File Format

Your Excel file must contain the following columns:
- **Project Task** - Task description
- **Date** - Time entry date
- **Duration** - Time in minutes (numeric value)

## Output File Format

The application writes data to the "Efforts" sheet in the output file:
- **Effort** - Time in hours (Duration / 60)
- **Description** - Value from the Project Task column
- **Date** - Date value (same format as in the input file)

## Error Handling

The application validates:
- Input file existence
- Required columns presence
- Data types (Duration must be numeric)
- File format (must be .xlsx or .xls)

When errors occur, a clear message with an explanation of the problem is displayed in the interface.

## Loading Custom Template

You can use your own Excel template:

1. Click the **"Load Template"** button
2. Select a .xlsx file via the file selection dialog
3. The selected template will be saved and used for all subsequent conversions

To revert to the built-in template, delete the `custom_template.xlsx` file from the application folder (if it was created).

## Features

- If the "Efforts" sheet is missing in the output file, it will be created automatically
- Rows with invalid Duration values are skipped with a warning
- Empty input data creates an empty output file
- The Excel template is embedded in the application, no external template files are required
