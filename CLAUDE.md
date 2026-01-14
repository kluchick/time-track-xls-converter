# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A desktop application with graphical user interface for converting time tracking data from Excel files. The application features drag-and-drop support, validates input files, converts duration from minutes to hours, and generates timestamped output files with embedded template support.

## Development Setup

**Environment:**
```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment (Windows)
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

**Dependencies:**
- Python 3.7+
- pandas >= 2.0.0 (for Excel file reading/processing)
- openpyxl >= 3.1.0 (for Excel file manipulation)
- customtkinter >= 5.2.0 (for modern graphical user interface)
- tkinterdnd2 >= 0.3.0 (for native OS drag-and-drop support)

## Running the Application

```bash
# Run the UI application
python src/ui_app.py
```

**Usage:**
- **Drag-and-drop** input Excel file directly from Windows Explorer onto the drop zone
- Or **enter file path manually** in the text field
- Or click **"Choose File..."** to browse for a file using file dialog
- Application validates the file and shows status in UI
- Click "Convert File" to perform conversion
- Output is saved in the project root directory with timestamp
- User can open the output folder directly from the UI

**Loading Custom Templates:**
- Click **"Load Template"** button to select a custom Excel template
- Selected template is saved to `files/custom_template.xlsx`
- Custom template persists across app sessions
- To revert to built-in template, delete `files/custom_template.xlsx`

**Output Format:** `output_YYYY-MM-DD_HH-MM-SS.xlsx` (e.g., `output_2026-01-13_14-30-45.xlsx`)

## Architecture

**Modular three-component design:**

### 1. **UI Layer** (`src/ui_app.py`)
CustomTkinter + tkinterdnd2 based graphical interface that handles:
- **Native OS file drag-and-drop** from Windows Explorer
- File selection via manual input, default file button, or file picker
- Input validation and user feedback
- Conversion triggering (runs synchronously - fast enough for typical files)
- Status display and error handling
- Platform-specific folder opening

**Key Classes:**
- `ConverterApp`: Main application class managing UI state and events (inherits from `TkinterDnD.Tk`)

**Drag-and-Drop Implementation:**
- Powered by tkinterdnd2 library for native OS file dropping
- Visual feedback during drag: green border highlight on drag enter
- Handles edge cases: multiple files (takes first), non-Excel files (shows error), paths with spaces/Unicode
- Events: `<<Drop>>`, `<<DragEnter>>`, `<<DragLeave>>`

### 2. **Business Logic Layer** (`src/converter_core.py`)
Core conversion functionality separated from UI:

**Main Functions:**
- `validate_input_data(file_path)`: Returns `(is_valid: bool, error_msg: str)`
- `read_and_process_input(file_path)`: Returns processed DataFrame
- `write_to_output_file(template_bytes, data_df, output_path)`: Writes to Excel, copies column A from row 3 if present
- `convert_excel_file(input_path, template_bytes, output_dir)`: Main conversion orchestrator

**Exception Handling:**
- Custom `ConversionError` exception for all conversion failures
- Detailed error messages propagated to UI

### 3. **Template Layer** (`src/template_data.py`)
Embedded Excel template (base64 encoded) with custom template support:

**Functions:**
- `get_template_bytes()`: Returns template as bytes (custom if available, else embedded)
- `get_template_stream()`: Returns BytesIO stream for openpyxl
- `has_custom_template()`: Check if custom template exists
- `save_custom_template(source_path)`: Save custom template to disk
- `get_template_info()`: Get description of active template

**Custom Template:**
- Location: `files/custom_template.xlsx`
- Automatically used if present and valid
- Falls back to embedded template if custom template fails to load

**Template Size:** ~93KB (base64 encoded), 68,956 bytes decoded

**File Structure:**
```
src/
  ui_app.py           # CustomTkinter + tkinterdnd2 UI application (main entry point)
  converter_core.py   # Business logic (testable, UI-independent)
  template_data.py    # Base64-encoded Excel template
  converter.py        # (Legacy CLI version, not used)
files/
  input.xlsx          # Sample input file
  output.xlsx         # Original template (now embedded)
task/                 # Miscellaneous project files
venv/                 # Python virtual environment (gitignored)
```

## Column Mappings

**Input columns (input.xlsx):**
- Project Task → Description
- Date → Date (preserved)
- Duration (minutes) → Effort (hours)

**Output columns (output.xlsx "Efforts" sheet):**
- Column A: If filled in row 3 (first data row), value is copied to all imported rows; otherwise preserved
- Column B: Effort (Duration / 60)
- Column C: Description (Project Task)
- Column D: Date

## Error Handling

The converter handles:
- Missing input file
- Missing output template
- Missing required columns in input
- Non-numeric Duration values (skipped with warning)
- Missing "Efforts" sheet (created if needed)
- Empty input data (creates empty output file)

## Custom Template Management

**Loading Custom Template:**
- Click "Load Template" button in UI
- Select .xlsx file via file dialog
- Template is validated (must be valid Excel file)
- Saved to `files/custom_template.xlsx` (gitignored)

**Template Priority:**
1. Check for `files/custom_template.xlsx`
2. If exists and valid, use it
3. If missing or invalid, fall back to embedded template

**Reverting to Built-in Template:**
- Delete `files/custom_template.xlsx` manually from file system

## Code Modification Guidelines

### UI Changes (`src/ui_app.py`)
- Use `self.update()` after modifying UI state (CustomTkinter/tkinter method)
- For visibility: use `.pack()` / `.grid()` to show, `.pack_forget()` / `.grid_forget()` to hide widgets
- For state: use `.configure(state="disabled")` / `.configure(state="normal")`
- For text/color: use `.configure(text="...", text_color="...")`
- Disable interactive elements during processing to prevent race conditions
- Provide clear user feedback for all states (loading, success, error)
- Handle exceptions gracefully and display user-friendly error messages
- **Drag-and-drop events**:
  - `<<Drop>>` - file dropped on drop zone
  - `<<DragEnter>>` - file dragged over drop zone (change border to green)
  - `<<DragLeave>>` - file left drop zone (reset border color)
  - Handle edge cases in `on_file_dropped()`: multiple files, spaces in paths, Unicode paths

### Business Logic Changes (`src/converter_core.py`)
- Maintain the column mapping constants: `INPUT_COLUMNS`, `OUTPUT_COLUMNS`, `EFFORTS_SHEET_NAME`
- Always raise `ConversionError` for expected failures (propagates to UI)
- Preserve row 3 as data start (rows 1-2 are headers in Efforts sheet)
- Column A handling: If row 3 column A has a value, copy it to all imported rows; otherwise preserve existing values
- Validate all input data before processing
- Handle pandas datetime objects properly for Excel date formatting

### Template Updates
- To update the embedded template, re-encode `files/output.xlsx`:
  ```bash
  python -c "import base64; print(base64.b64encode(open('files/output.xlsx', 'rb').read()).decode())" > new_template.txt
  ```
- Replace `TEMPLATE_BASE64` in `src/template_data.py` with new encoded data
- Template must maintain the Efforts sheet structure (rows 1-2 headers, row 3+ data)

## Testing

**Manual UI Testing:**
```bash
python src/ui_app.py
```

**Core Logic Testing:**
```python
from pathlib import Path
from src.converter_core import convert_excel_file
from src.template_data import get_template_bytes

input_file = Path('files/input.xlsx')
template_bytes = get_template_bytes()
output_path = convert_excel_file(input_file, template_bytes, Path('.'))
print(f"Output: {output_path}")
```

**Validation Testing:**
```python
from pathlib import Path
from src.converter_core import validate_input_data

is_valid, error = validate_input_data(Path('files/input.xlsx'))
print(f"Valid: {is_valid}, Error: {error}")
```

## Common Development Tasks

**Add new validation rule:**
1. Update `validate_input_data()` in `src/converter_core.py`
2. Return `(False, "error message")` for validation failures
3. UI will automatically display the error message

**Change output filename format:**
1. Modify timestamp format in `convert_excel_file()` function
2. Update `datetime.now().strftime()` format string

**Add new UI element:**
1. Create CustomTkinter component in `setup_ui()` method (e.g., `ctk.CTkButton`, `ctk.CTkLabel`, `ctk.CTkEntry`)
2. Add to layout using `.pack()` or `.grid()` geometry manager
3. Wire up event handlers using `command=` parameter or `.bind()` method
4. Update state and call `self.update()` to refresh UI

**Modify column mappings:**
1. Update constants in `src/converter_core.py`: `INPUT_COLUMNS`, `OUTPUT_COLUMNS`
2. Update processing logic in `read_and_process_input()` and `write_to_output_file()`
3. Keep column indices consistent (B=2, C=3, D=4 in `write_to_output_file`)
