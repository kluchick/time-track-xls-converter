"""
Excel Time Tracker Converter - UI Application

CustomTkinter + tkinterdnd2 based graphical user interface for Excel time tracking conversion.
Supports native OS file drag-and-drop, manual file selection, and displays conversion status.
"""

import subprocess
import sys
from pathlib import Path
from tkinter import filedialog

import customtkinter as ctk
import openpyxl
from tkinterdnd2 import DND_FILES, TkinterDnD

from converter_core import ConversionError, convert_excel_file, validate_input_data
from template_data import (
    get_app_base_dir,
    get_template_bytes,
    get_template_info,
    save_custom_template,
)

# UI Color Constants
COLOR_PRIMARY = "#1976d2"
COLOR_PRIMARY_DARK = "#1565c0"
COLOR_SUCCESS = "green"
COLOR_ERROR = "red"
COLOR_WARNING = "orange"
COLOR_DROP_ZONE_BG = "#e3f2fd"
COLOR_DROP_ZONE_BORDER = "#90caf9"
COLOR_DROP_ZONE_ACTIVE = "#4caf50"
COLOR_TEMPLATE_BTN = "#7e57c2"
COLOR_TEMPLATE_BTN_HOVER = "#673ab7"

# Valid Excel extensions
VALID_EXCEL_EXTENSIONS = {".xlsx", ".xls"}


class ConverterApp(TkinterDnD.Tk):
    """Main application class for the Excel converter UI with drag-and-drop support."""

    def __init__(self) -> None:
        super().__init__()

        self.selected_file = None
        self.output_file = None
        self.template_info = None

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title("Excel Time Tracker Converter")
        self.geometry("650x750")
        self.resizable(False, False)
        self.configure(bg="#f0f0f0")

        self.setup_ui()

    def setup_ui(self) -> None:
        """Setup all UI components and layout."""
        self._create_title_section()
        self.create_drop_zone()
        self._create_selected_file_label()
        self._create_extend_time_section()
        self._create_separator()
        self._create_status_section()
        self._create_action_buttons()

    def _create_title_section(self) -> None:
        title_frame = ctk.CTkFrame(self, fg_color="transparent")
        title_frame.pack(pady=(20, 10), padx=20)

        ctk.CTkLabel(
            title_frame,
            text="Excel Time Tracker Converter",
            font=("Arial", 28, "bold"),
            text_color=COLOR_PRIMARY_DARK,
        ).pack()

        ctk.CTkLabel(
            title_frame,
            text="Convert time tracking data from minutes to hours",
            font=("Arial", 14),
            text_color="gray",
        ).pack(pady=(5, 0))

    def _create_selected_file_label(self) -> None:
        self.selected_file_label = ctk.CTkLabel(
            self,
            text="",
            font=("Arial", 12, "italic"),
            text_color="gray",
        )
        self.selected_file_label.pack(pady=(5, 0))

    def _create_extend_time_section(self) -> None:
        extend_frame = ctk.CTkFrame(self, fg_color="transparent")
        extend_frame.pack(pady=(5, 0), padx=20, fill="x")

        ctk.CTkLabel(
            extend_frame,
            text="Extend time entries (optional):",
            font=("Arial", 12),
            text_color="gray",
        ).pack(anchor="w", pady=(0, 5))

        self.extend_time_entry = ctk.CTkEntry(
            extend_frame,
            width=550,
            placeholder_text="Enter description for extended time entries",
            border_width=2,
            height=35,
        )
        self.extend_time_entry.pack(fill="x")

    def _create_separator(self) -> None:
        ctk.CTkFrame(self, height=2, fg_color="lightgray").pack(
            pady=20, padx=20, fill="x"
        )

    def _create_status_section(self) -> None:
        self.status_label = ctk.CTkLabel(
            self,
            text="Ready to convert",
            font=("Arial", 14, "bold"),
        )
        self.status_label.pack(pady=10)

        self.progress_bar = ctk.CTkProgressBar(self, width=400, mode="indeterminate")
        self.progress_bar.pack(pady=10)
        self.progress_bar.pack_forget()

        self.output_path_label = ctk.CTkLabel(
            self,
            text="",
            font=("Arial", 12),
            text_color=COLOR_SUCCESS,
        )
        self.output_path_label.pack()
        self.output_path_label.pack_forget()

        self.duplicates_label = ctk.CTkLabel(
            self,
            text="",
            font=("Arial", 11),
            text_color=COLOR_WARNING,
        )
        self.duplicates_label.pack()
        self.duplicates_label.pack_forget()

    def _create_action_buttons(self) -> None:
        button_container = ctk.CTkFrame(self, fg_color="transparent")
        button_container.pack(pady=(10, 20))

        self.convert_btn = ctk.CTkButton(
            button_container,
            text="Convert File",
            width=150,
            height=35,
            font=("Arial", 14, "bold"),
            command=self.on_convert_clicked,
            state="disabled",
        )
        self.convert_btn.grid(row=0, column=0, padx=10)

        self.open_folder_btn = ctk.CTkButton(
            button_container,
            text="Open Output Folder",
            width=150,
            height=35,
            font=("Arial", 13),
            command=self.on_open_folder_clicked,
            fg_color="gray",
            hover_color="darkgray",
        )
        self.open_folder_btn.grid(row=0, column=1, padx=10)
        self.open_folder_btn.grid_forget()

        self.reset_btn = ctk.CTkButton(
            button_container,
            text="Select Another File",
            width=150,
            height=35,
            font=("Arial", 13),
            command=self.on_reset_clicked,
            fg_color=COLOR_WARNING,
            hover_color="darkorange",
        )
        self.reset_btn.grid(row=0, column=2, padx=10)
        self.reset_btn.grid_forget()

    def create_drop_zone(self) -> None:
        """Create file drop zone with drag-and-drop support."""
        self.drop_frame = ctk.CTkFrame(
            self,
            width=550,
            height=300,
            corner_radius=10,
            fg_color=COLOR_DROP_ZONE_BG,
            border_width=2,
            border_color=COLOR_DROP_ZONE_BORDER,
        )
        self.drop_frame.pack(pady=20, padx=20)
        self.drop_frame.pack_propagate(False)

        ctk.CTkLabel(self.drop_frame, text="📄", font=("Arial", 60)).pack(pady=(30, 10))
        ctk.CTkLabel(
            self.drop_frame, text="Drop Excel file here", font=("Arial", 16, "bold")
        ).pack(pady=(0, 15))

        self.file_path_entry = ctk.CTkEntry(
            self.drop_frame,
            width=480,
            placeholder_text="Or enter file path manually",
            border_width=2,
            height=35,
        )
        self.file_path_entry.pack(pady=(0, 15))
        self.file_path_entry.bind("<KeyRelease>", self.on_file_path_changed)

        self._create_drop_zone_buttons()

        ctk.CTkLabel(
            self.drop_frame,
            text="Required columns: Project Task, Date, Duration",
            font=("Arial", 10),
            text_color="gray",
        ).pack(pady=(15, 0))

        self.template_info = ctk.CTkLabel(
            self.drop_frame,
            text="",
            font=("Arial", 10),
            text_color="purple",
            anchor="w",
            width=480,
        )
        self.template_info.pack(pady=(5, 10))
        self.update_template_info()

        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind("<<Drop>>", self.on_file_dropped)
        self.drop_frame.dnd_bind("<<DragEnter>>", self.on_drag_enter)
        self.drop_frame.dnd_bind("<<DragLeave>>", self.on_drag_leave)

    def _create_drop_zone_buttons(self) -> None:
        button_frame = ctk.CTkFrame(self.drop_frame, fg_color="transparent")
        button_frame.pack()

        ctk.CTkButton(
            button_frame,
            text="Load ETS template",
            width=180,
            command=self.on_load_template_clicked,
            fg_color=COLOR_TEMPLATE_BTN,
            hover_color=COLOR_TEMPLATE_BTN_HOVER,
        ).grid(row=0, column=0, padx=5)

        ctk.CTkButton(
            button_frame,
            text="Choose File...",
            width=180,
            command=self.on_browse_clicked,
            fg_color=COLOR_PRIMARY,
            hover_color=COLOR_PRIMARY_DARK,
        ).grid(row=0, column=1, padx=5)

    def _set_status(self, message: str, color: str) -> None:
        """Update the status label with a message and color."""
        self.status_label.configure(text=message, text_color=color)

    def _set_drop_zone_border(self, color: str) -> None:
        """Update the drop zone border color."""
        self.drop_frame.configure(border_color=color)

    def on_file_dropped(self, event) -> None:
        """Handle file drop event from drag-and-drop."""
        file_path_str = event.data.strip("{}")

        # Handle multiple files - take only the first one
        if " " in file_path_str and not Path(file_path_str).exists():
            file_path_str = file_path_str.split()[0].strip("{}")

        file_path = Path(file_path_str)

        self.file_path_entry.delete(0, "end")
        self.file_path_entry.insert(0, str(file_path))
        self.validate_and_set_file(file_path)
        self._set_drop_zone_border(COLOR_DROP_ZONE_BORDER)

    def on_drag_enter(self, event) -> None:
        """Highlight drop zone when file is dragged over it."""
        self._set_drop_zone_border(COLOR_DROP_ZONE_ACTIVE)

    def on_drag_leave(self, event) -> None:
        """Reset drop zone highlight when file leaves."""
        self._set_drop_zone_border(COLOR_DROP_ZONE_BORDER)

    def on_file_path_changed(self, event) -> None:
        """Handle manual file path entry changes."""
        file_path_text = self.file_path_entry.get().strip()
        if file_path_text:
            self.validate_and_set_file(Path(file_path_text))

    def on_load_template_clicked(self) -> None:
        """Open file dialog to load a custom template file."""
        file_path = filedialog.askopenfilename(
            title="Select Excel template file",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=str(get_app_base_dir()),
        )

        if not file_path:
            return

        template_path = Path(file_path)

        if template_path.suffix.lower() != ".xlsx":
            self._set_status("Error: Template must be .xlsx format", COLOR_ERROR)
            return

        validation_error = self._validate_excel_file(template_path)
        if validation_error:
            self._set_status(f"Invalid Excel file: {validation_error}", COLOR_ERROR)
            return

        save_error = self._save_template(template_path)
        if save_error:
            self._set_status(f"Error saving template: {save_error}", COLOR_ERROR)
            return

        self._set_status(f"Template loaded: {template_path.name}", COLOR_SUCCESS)
        self.update_template_info()

    def _validate_excel_file(self, file_path: Path) -> str:
        """Validate that a file is a valid Excel file. Returns error message or empty string."""
        try:
            wb = openpyxl.load_workbook(file_path)
            wb.close()
            return ""
        except Exception as ex:
            return str(ex)

    def _save_template(self, template_path: Path) -> str:
        """Save template file. Returns error message or empty string."""
        try:
            save_custom_template(template_path)
            return ""
        except Exception as ex:
            return str(ex)

    def update_template_info(self) -> None:
        """Update the template info label to show which template is active."""
        if self.template_info:
            self.template_info.configure(text=f"Template: {get_template_info()}")

    def on_browse_clicked(self) -> None:
        """Open file dialog for file selection."""
        file_path = filedialog.askopenfilename(
            title="Select input Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
            initialdir=str(get_app_base_dir()),
        )

        if file_path:
            self.file_path_entry.delete(0, "end")
            self.file_path_entry.insert(0, file_path)
            self.validate_and_set_file(Path(file_path))

    def validate_and_set_file(self, file_path: Path) -> None:
        """Validate selected file and update UI state accordingly."""
        if file_path.suffix.lower() not in VALID_EXCEL_EXTENSIONS:
            self._set_file_invalid("Error: File must be Excel format (.xlsx or .xls)")
            return

        is_valid, error_msg = validate_input_data(file_path)

        if is_valid:
            self._set_file_valid(file_path)
        else:
            self._set_file_invalid(f"Invalid file: {error_msg}")

    def _set_file_valid(self, file_path: Path) -> None:
        """Update UI for a valid file selection."""
        self.selected_file = file_path
        self.selected_file_label.configure(text=f"Selected: {file_path.name}")
        self._set_status("File validated successfully! Ready to convert.", COLOR_SUCCESS)
        self.convert_btn.configure(state="normal")

    def _set_file_invalid(self, error_message: str) -> None:
        """Update UI for an invalid file selection."""
        self.selected_file = None
        self.selected_file_label.configure(text="")
        self._set_status(error_message, COLOR_ERROR)
        self.convert_btn.configure(state="disabled")

    def on_convert_clicked(self) -> None:
        """Handle convert button click."""
        self._set_ui_converting()
        self.update()
        self._perform_conversion()

    def _set_ui_converting(self) -> None:
        """Set UI to converting state."""
        self.convert_btn.configure(state="disabled")
        self.file_path_entry.configure(state="disabled")
        self.extend_time_entry.configure(state="disabled")
        self.progress_bar.pack(pady=10)
        self.progress_bar.start()
        self._set_status("Converting... Please wait.", COLOR_PRIMARY)
        self.output_path_label.pack_forget()
        self.duplicates_label.pack_forget()
        self.open_folder_btn.grid_forget()
        self.reset_btn.grid_forget()

    def _perform_conversion(self) -> None:
        """Perform the file conversion and update UI with results."""
        extend_description = self.extend_time_entry.get().strip() or None

        try:
            output_path, duplicates_count = convert_excel_file(
                input_path=self.selected_file,
                template_bytes=get_template_bytes(),
                output_dir=get_app_base_dir(),
                extend_description=extend_description,
            )
            self._handle_conversion_success(output_path, duplicates_count)

        except ConversionError as ex:
            self._handle_conversion_error(f"Error: {ex}")

        except Exception as ex:
            self._handle_conversion_error(f"Unexpected error: {ex}")

    def _handle_conversion_success(self, output_path: Path, duplicates_count: int) -> None:
        """Update UI after successful conversion."""
        self.output_file = output_path
        self._set_status("Conversion completed successfully!", COLOR_SUCCESS)
        self.output_path_label.configure(text=f"Output saved: {output_path.name}")
        self.output_path_label.pack()

        if duplicates_count > 0:
            self.duplicates_label.configure(
                text=f"Skipped {duplicates_count} duplicate row(s)"
            )
            self.duplicates_label.pack()
        else:
            self.duplicates_label.pack_forget()

        self.open_folder_btn.grid(row=0, column=1, padx=10)
        self.reset_btn.grid(row=0, column=2, padx=10)
        self._stop_progress()

    def _handle_conversion_error(self, error_message: str) -> None:
        """Update UI after conversion error."""
        self._set_status(error_message, COLOR_ERROR)
        self._enable_input_controls()
        self._stop_progress()

    def _enable_input_controls(self) -> None:
        """Re-enable input controls after conversion."""
        self.convert_btn.configure(state="normal")
        self.file_path_entry.configure(state="normal")
        self.extend_time_entry.configure(state="normal")

    def _stop_progress(self) -> None:
        """Stop and hide the progress bar."""
        self.progress_bar.stop()
        self.progress_bar.pack_forget()

    def on_open_folder_clicked(self) -> None:
        """Open output folder in the system file explorer."""
        if not self.output_file:
            return

        output_dir = str(self.output_file.parent)

        commands_by_platform = {
            "win32": ["explorer", output_dir],
            "darwin": ["open", output_dir],
        }
        command = commands_by_platform.get(sys.platform, ["xdg-open", output_dir])

        try:
            subprocess.run(command, check=False)
        except Exception as ex:
            self._set_status(f"Could not open folder: {ex}", COLOR_WARNING)

    def on_reset_clicked(self) -> None:
        """Reset UI for a new conversion."""
        self.selected_file = None
        self.output_file = None
        self.selected_file_label.configure(text="")
        self._set_status("Ready to convert", "black")
        self.output_path_label.pack_forget()
        self.duplicates_label.pack_forget()
        self.convert_btn.configure(state="disabled")
        self.file_path_entry.configure(state="normal")
        self.file_path_entry.delete(0, "end")
        self.extend_time_entry.configure(state="normal")
        self.extend_time_entry.delete(0, "end")
        self.open_folder_btn.grid_forget()
        self.reset_btn.grid_forget()


def main() -> None:
    """Main entry point for the application."""
    app = ConverterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
