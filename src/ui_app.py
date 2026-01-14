"""
Excel Time Tracker Converter - UI Application

CustomTkinter + tkinterdnd2 based graphical user interface for Excel time tracking conversion.
Supports native OS file drag-and-drop, manual file selection, and displays conversion status.
"""

import customtkinter as ctk
from tkinterdnd2 import DND_FILES, TkinterDnD
from pathlib import Path
import subprocess
import sys
from tkinter import filedialog

from converter_core import convert_excel_file, validate_input_data, ConversionError
from template_data import get_template_bytes


class ConverterApp(TkinterDnD.Tk):
    """Main application class for the Excel converter UI with drag-and-drop support"""

    def __init__(self):
        """Initialize the converter application"""
        super().__init__()

        # Initialize state
        self.selected_file = None
        self.output_file = None

        # Configure theme (must be before creating widgets)
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        # Setup window
        self.title("Excel Time Tracker Converter")
        self.geometry("650x700")
        self.resizable(False, False)
        self.configure(bg="#f0f0f0")

        # Initialize UI
        self.setup_ui()

    def setup_ui(self):
        """Setup all UI components and layout"""

        # Title section
        title_frame = ctk.CTkFrame(self, fg_color="transparent")
        title_frame.pack(pady=(20, 10), padx=20)

        title = ctk.CTkLabel(
            title_frame,
            text="Excel Time Tracker Converter",
            font=("Arial", 28, "bold"),
            text_color="#1565c0"
        )
        title.pack()

        subtitle = ctk.CTkLabel(
            title_frame,
            text="Convert time tracking data from minutes to hours",
            font=("Arial", 14),
            text_color="gray"
        )
        subtitle.pack(pady=(5, 0))

        # Drop zone
        self.create_drop_zone()

        # Selected file display
        self.selected_file_label = ctk.CTkLabel(
            self,
            text="",
            font=("Arial", 12, "italic"),
            text_color="gray"
        )
        self.selected_file_label.pack(pady=(10, 0))

        # Separator
        separator = ctk.CTkFrame(self, height=2, fg_color="lightgray")
        separator.pack(pady=20, padx=20, fill="x")

        # Status area
        self.status_label = ctk.CTkLabel(
            self,
            text="Ready to convert",
            font=("Arial", 14, "bold")
        )
        self.status_label.pack(pady=10)

        self.progress_bar = ctk.CTkProgressBar(self, width=400, mode="indeterminate")
        self.progress_bar.pack(pady=10)
        self.progress_bar.pack_forget()  # Hide initially

        self.output_path_label = ctk.CTkLabel(
            self,
            text="",
            font=("Arial", 12),
            text_color="green"
        )
        self.output_path_label.pack()
        self.output_path_label.pack_forget()  # Hide initially

        # Action buttons
        button_container = ctk.CTkFrame(self, fg_color="transparent")
        button_container.pack(pady=20)

        self.convert_btn = ctk.CTkButton(
            button_container,
            text="Convert File",
            width=150,
            height=35,
            font=("Arial", 14, "bold"),
            command=self.on_convert_clicked,
            state="disabled"
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
            hover_color="darkgray"
        )
        self.open_folder_btn.grid(row=0, column=1, padx=10)
        self.open_folder_btn.grid_forget()  # Hide initially

        self.reset_btn = ctk.CTkButton(
            button_container,
            text="Select Another File",
            width=150,
            height=35,
            font=("Arial", 13),
            command=self.on_reset_clicked,
            fg_color="orange",
            hover_color="darkorange"
        )
        self.reset_btn.grid(row=0, column=2, padx=10)
        self.reset_btn.grid_forget()  # Hide initially

    def create_drop_zone(self):
        """Create file drop zone with drag-and-drop support"""

        # Main drop zone frame
        self.drop_frame = ctk.CTkFrame(
            self,
            width=550,
            height=280,
            corner_radius=10,
            fg_color="#e3f2fd",
            border_width=2,
            border_color="#90caf9"
        )
        self.drop_frame.pack(pady=20, padx=20)
        self.drop_frame.pack_propagate(False)  # Maintain fixed size

        # File icon
        file_icon = ctk.CTkLabel(
            self.drop_frame,
            text="📄",
            font=("Arial", 60)
        )
        file_icon.pack(pady=(30, 10))

        # Instruction text
        instruction_text = ctk.CTkLabel(
            self.drop_frame,
            text="Drop Excel file here",
            font=("Arial", 16, "bold")
        )
        instruction_text.pack(pady=(0, 15))

        # File path entry
        self.file_path_entry = ctk.CTkEntry(
            self.drop_frame,
            width=480,
            placeholder_text="Or enter file path manually",
            border_width=2,
            height=35
        )
        self.file_path_entry.pack(pady=(0, 15))
        self.file_path_entry.bind("<KeyRelease>", self.on_file_path_changed)

        # Button frame
        button_frame = ctk.CTkFrame(self.drop_frame, fg_color="transparent")
        button_frame.pack()

        default_btn = ctk.CTkButton(
            button_frame,
            text="Use files/input.xlsx",
            width=180,
            command=self.on_use_default_file,
            fg_color="#1976d2",
            hover_color="#1565c0"
        )
        default_btn.grid(row=0, column=0, padx=5)

        browse_btn = ctk.CTkButton(
            button_frame,
            text="Choose File...",
            width=180,
            command=self.on_browse_clicked,
            fg_color="#1976d2",
            hover_color="#1565c0"
        )
        browse_btn.grid(row=0, column=1, padx=5)

        # Hint text
        hint_label = ctk.CTkLabel(
            self.drop_frame,
            text="Required columns: Project Task, Date, Duration",
            font=("Arial", 10),
            text_color="gray"
        )
        hint_label.pack(pady=(15, 0))

        # Register drag-and-drop events
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind("<<Drop>>", self.on_file_dropped)
        self.drop_frame.dnd_bind("<<DragEnter>>", self.on_drag_enter)
        self.drop_frame.dnd_bind("<<DragLeave>>", self.on_drag_leave)

    def on_file_dropped(self, event):
        """
        Handle file drop event.

        Args:
            event: TkinterDnD event with file path data
        """
        # Parse file path from event
        # Format: {C:/path/to/file.xlsx} for paths with spaces
        file_path_str = event.data.strip('{}')

        # Handle multiple files (take only first)
        if ' ' in file_path_str and not Path(file_path_str).exists():
            # Multiple files dropped, split and take first
            file_path_str = file_path_str.split()[0].strip('{}')

        file_path = Path(file_path_str)

        # Update entry field
        self.file_path_entry.delete(0, 'end')
        self.file_path_entry.insert(0, str(file_path))

        # Validate file
        self.validate_and_set_file(file_path)

        # Reset visual feedback
        self.drop_frame.configure(border_color="#90caf9")

    def on_drag_enter(self, event):
        """
        Handle drag enter event (visual feedback).

        Args:
            event: TkinterDnD event
        """
        # Change border color to green to indicate drop zone is active
        self.drop_frame.configure(border_color="#4caf50")

    def on_drag_leave(self, event):
        """
        Handle drag leave event (reset visual feedback).

        Args:
            event: TkinterDnD event
        """
        # Reset border color
        self.drop_frame.configure(border_color="#90caf9")

    def on_file_path_changed(self, event):
        """
        Handle file path entry changes.

        Args:
            event: Tkinter event
        """
        file_path_text = self.file_path_entry.get().strip()
        if file_path_text:
            file_path = Path(file_path_text)
            self.validate_and_set_file(file_path)

    def on_use_default_file(self):
        """Use the default input file from files/input.xlsx"""
        default_path = Path(__file__).parent.parent / "files" / "input.xlsx"
        self.file_path_entry.delete(0, 'end')
        self.file_path_entry.insert(0, str(default_path))
        self.validate_and_set_file(default_path)

    def on_browse_clicked(self):
        """Open file dialog for file selection"""
        try:
            # Open file dialog
            file_path = filedialog.askopenfilename(
                title="Select input Excel file",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
                initialdir=str(Path(__file__).parent.parent / "files")
            )

            if file_path:
                self.file_path_entry.delete(0, 'end')
                self.file_path_entry.insert(0, file_path)
                self.validate_and_set_file(Path(file_path))

        except Exception as ex:
            self.status_label.configure(
                text=f"Error opening file dialog: {str(ex)}",
                text_color="red"
            )

    def validate_and_set_file(self, file_path: Path):
        """
        Validate selected file and update UI.

        Args:
            file_path: Path to selected file
        """
        # Check file extension first
        if file_path.suffix.lower() not in ['.xlsx', '.xls']:
            self.selected_file = None
            self.selected_file_label.configure(text="")
            self.status_label.configure(
                text="❌ Error: File must be Excel format (.xlsx or .xls)",
                text_color="red"
            )
            self.convert_btn.configure(state="disabled")
            return

        # Validate file content
        is_valid, error_msg = validate_input_data(file_path)

        if is_valid:
            self.selected_file = file_path
            self.selected_file_label.configure(text=f"Selected: {file_path.name}")
            self.status_label.configure(
                text="File validated successfully! Ready to convert.",
                text_color="green"
            )
            self.convert_btn.configure(state="normal")
        else:
            self.selected_file = None
            self.selected_file_label.configure(text="")
            self.status_label.configure(
                text=f"❌ Invalid file: {error_msg}",
                text_color="red"
            )
            self.convert_btn.configure(state="disabled")

    def on_convert_clicked(self):
        """Handle convert button click"""
        # Disable UI during conversion
        self.convert_btn.configure(state="disabled")
        self.file_path_entry.configure(state="disabled")
        self.progress_bar.pack(pady=10)
        self.progress_bar.start()
        self.status_label.configure(
            text="Converting... Please wait.",
            text_color="#1976d2"
        )
        self.output_path_label.pack_forget()
        self.open_folder_btn.grid_forget()
        self.reset_btn.grid_forget()
        self.update()

        # Perform conversion
        self.perform_conversion()

    def perform_conversion(self):
        """
        Perform conversion and update UI with results.
        """
        try:
            # Get script directory for output (project root)
            script_dir = Path(__file__).parent.parent

            # Get template bytes
            template_bytes = get_template_bytes()

            # Convert file
            output_path = convert_excel_file(
                input_path=self.selected_file,
                template_bytes=template_bytes,
                output_dir=script_dir
            )

            # Update UI on success
            self.output_file = output_path
            self.status_label.configure(
                text="✅ Conversion completed successfully!",
                text_color="green"
            )
            self.output_path_label.configure(text=f"Output saved: {output_path.name}")
            self.output_path_label.pack()
            self.open_folder_btn.grid(row=0, column=1, padx=10)
            self.reset_btn.grid(row=0, column=2, padx=10)
            self.progress_bar.stop()
            self.progress_bar.pack_forget()

        except ConversionError as ex:
            # Update UI on error
            self.status_label.configure(
                text=f"❌ Error: {str(ex)}",
                text_color="red"
            )
            self.convert_btn.configure(state="normal")
            self.file_path_entry.configure(state="normal")
            self.progress_bar.stop()
            self.progress_bar.pack_forget()

        except Exception as ex:
            # Update UI on unexpected error
            self.status_label.configure(
                text=f"❌ Unexpected error: {str(ex)}",
                text_color="red"
            )
            self.convert_btn.configure(state="normal")
            self.file_path_entry.configure(state="normal")
            self.progress_bar.stop()
            self.progress_bar.pack_forget()

    def on_open_folder_clicked(self):
        """Open output folder in file explorer"""
        if self.output_file:
            output_dir = self.output_file.parent

            # Platform-specific folder opening
            try:
                if sys.platform == "win32":
                    subprocess.run(["explorer", str(output_dir)], check=False)
                elif sys.platform == "darwin":
                    subprocess.run(["open", str(output_dir)], check=False)
                else:  # Linux
                    subprocess.run(["xdg-open", str(output_dir)], check=False)
            except Exception as ex:
                self.status_label.configure(
                    text=f"Could not open folder: {str(ex)}",
                    text_color="orange"
                )

    def on_reset_clicked(self):
        """Reset UI for new conversion"""
        self.selected_file = None
        self.output_file = None
        self.selected_file_label.configure(text="")
        self.status_label.configure(
            text="Ready to convert",
            text_color="black"
        )
        self.output_path_label.pack_forget()
        self.convert_btn.configure(state="disabled")
        self.file_path_entry.configure(state="normal")
        self.file_path_entry.delete(0, 'end')
        self.open_folder_btn.grid_forget()
        self.reset_btn.grid_forget()


def main():
    """Main entry point for the application"""
    app = ConverterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
