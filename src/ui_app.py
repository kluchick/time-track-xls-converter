"""
Excel Time Tracker Converter - UI Application

Flet-based graphical user interface for Excel time tracking conversion.
Supports drag-and-drop file selection and displays conversion status.
"""

import flet as ft
from pathlib import Path
import threading
import subprocess
import sys

from converter_core import convert_excel_file, validate_input_data, ConversionError
from template_data import get_template_bytes


class ConverterApp:
    """Main application class for the Excel converter UI"""

    def __init__(self, page: ft.Page):
        """
        Initialize the converter application.

        Args:
            page: Flet page object
        """
        self.page = page
        self.selected_file = None
        self.output_file = None

        # Setup page
        self.page.title = "Excel Time Tracker Converter"
        self.page.window_width = 650
        self.page.window_height = 550
        self.page.window_resizable = False
        self.page.padding = 20

        # Initialize UI components
        self.setup_ui()

    def setup_ui(self):
        """Setup all UI components and layout"""
        # No file picker needed - using TextField instead

        # Title
        title = ft.Text(
            "Excel Time Tracker Converter",
            size=28,
            weight=ft.FontWeight.BOLD,
            color=ft.Colors.BLUE_900
        )

        # Subtitle
        subtitle = ft.Text(
            "Convert time tracking data from minutes to hours",
            size=14,
            color=ft.Colors.GREY_700
        )

        # Drop zone for file selection
        self.drop_zone = self.create_drop_zone()

        # Selected file display
        self.selected_file_text = ft.Text(
            "",
            size=12,
            color=ft.Colors.GREY_600,
            visible=False,
            italic=True
        )

        # Status area
        self.status_text = ft.Text(
            "Ready to convert",
            size=14,
            weight=ft.FontWeight.W_500
        )

        self.progress_bar = ft.ProgressBar(
            width=400,
            visible=False
        )

        self.output_path_text = ft.Text(
            "",
            size=12,
            color=ft.Colors.GREEN_700,
            visible=False,
            selectable=True
        )

        # Action buttons
        self.convert_btn = ft.Button(
            "Convert File",
            icon="transform",
            on_click=self.on_convert_clicked,
            disabled=True,
            style=ft.ButtonStyle(
                bgcolor=ft.Colors.BLUE_700,
                color=ft.Colors.WHITE
            )
        )

        self.open_folder_btn = ft.Button(
            "Open Output Folder",
            icon="folder_open",
            on_click=self.on_open_folder_clicked,
            visible=False
        )

        self.reset_btn = ft.TextButton(
            "Select Another File",
            on_click=self.on_reset_clicked,
            visible=False
        )

        # Layout
        self.page.add(
            ft.Column([
                # Header
                ft.Container(
                    content=ft.Column([
                        title,
                        subtitle,
                    ], spacing=5),
                    padding=ft.Padding.only(bottom=20)
                ),

                # Drop zone
                self.drop_zone,

                # Selected file
                self.selected_file_text,

                # Divider
                ft.Divider(height=20, color=ft.Colors.GREY_300),

                # Status area
                ft.Container(
                    content=ft.Column([
                        self.status_text,
                        self.progress_bar,
                        self.output_path_text,
                    ], spacing=10, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                    padding=ft.Padding.symmetric(vertical=10)
                ),

                # Action buttons
                ft.Row([
                    self.convert_btn,
                    self.open_folder_btn,
                    self.reset_btn,
                ], alignment=ft.MainAxisAlignment.CENTER, spacing=15),
            ],
                spacing=15,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER
            )
        )

    def create_drop_zone(self):
        """
        Create file selection area.

        Returns:
            Container with file selection UI
        """
        # File path input field
        self.file_path_input = ft.TextField(
            label="File path",
            hint_text="Enter path to input.xlsx or click button below",
            width=500,
            on_change=self.on_file_path_changed,
            border_color=ft.Colors.BLUE_400
        )

        return ft.Container(
            content=ft.Column([
                ft.Text(
                    "📄",
                    size=60
                ),
                ft.Text(
                    "Select input Excel file",
                    size=16,
                    weight=ft.FontWeight.W_500
                ),
                self.file_path_input,
                ft.Row([
                    ft.Button(
                        "Use files/input.xlsx",
                        icon="folder",
                        on_click=self.on_use_default_file
                    ),
                    ft.Button(
                        "Choose File...",
                        icon="file_open",
                        on_click=self.on_browse_clicked
                    ),
                ], alignment=ft.MainAxisAlignment.CENTER, spacing=10),
                ft.Text(
                    "Required columns: Project Task, Date, Duration",
                    size=10,
                    color=ft.Colors.GREY_500,
                    italic=True
                ),
            ],
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=15
            ),
            border=ft.Border.all(2, ft.Colors.BLUE_200),
            border_radius=10,
            padding=30,
            bgcolor=ft.Colors.BLUE_50,
            width=550,
            alignment=ft.Alignment.CENTER
        )

    def on_file_path_changed(self, e):
        """
        Handle file path text field changes.

        Args:
            e: TextField change event
        """
        if self.file_path_input.value:
            file_path = Path(self.file_path_input.value.strip())
            self.validate_and_set_file(file_path)

    def on_use_default_file(self, e):
        """
        Use the default input file from files/input.xlsx.

        Args:
            e: Button click event
        """
        default_path = Path(__file__).parent.parent / "files" / "input.xlsx"
        self.file_path_input.value = str(default_path)
        self.page.update()
        self.validate_and_set_file(default_path)

    def on_browse_clicked(self, e):
        """
        Open file dialog using tkinter (built-in Python).

        Args:
            e: Button click event
        """
        try:
            import tkinter as tk
            from tkinter import filedialog

            # Create hidden root window
            root = tk.Tk()
            root.withdraw()
            root.wm_attributes('-topmost', 1)

            # Open file dialog
            file_path = filedialog.askopenfilename(
                title="Select input Excel file",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
                initialdir=str(Path(__file__).parent.parent / "files")
            )

            root.destroy()

            if file_path:
                self.file_path_input.value = file_path
                self.page.update()
                self.validate_and_set_file(Path(file_path))

        except Exception as ex:
            self.status_text.value = f"Error opening file dialog: {str(ex)}"
            self.status_text.color = ft.Colors.RED_700
            self.page.update()

    def validate_and_set_file(self, file_path: Path):
        """
        Validate selected file and update UI.

        Args:
            file_path: Path to selected file
        """
        is_valid, error_msg = validate_input_data(file_path)

        if is_valid:
            self.selected_file = file_path
            self.selected_file_text.value = f"Selected: {file_path.name}"
            self.selected_file_text.visible = True
            self.status_text.value = "File validated successfully! Ready to convert."
            self.status_text.color = ft.Colors.GREEN_700
            self.convert_btn.disabled = False
        else:
            self.selected_file = None
            self.selected_file_text.visible = False
            self.status_text.value = f"❌ Invalid file: {error_msg}"
            self.status_text.color = ft.Colors.RED_700
            self.convert_btn.disabled = True

        self.page.update()

    def on_convert_clicked(self, e):
        """
        Handle convert button click.

        Args:
            e: Button click event
        """
        # Disable UI during conversion
        self.convert_btn.disabled = True
        self.file_path_input.disabled = True
        self.progress_bar.visible = True
        self.status_text.value = "Converting... Please wait."
        self.status_text.color = ft.Colors.BLUE_700
        self.output_path_text.visible = False
        self.open_folder_btn.visible = False
        self.reset_btn.visible = False
        self.page.update()

        # Run conversion synchronously (it's fast enough)
        self.perform_conversion()

    def perform_conversion(self):
        """
        Perform conversion in background thread.
        Updates UI with results when complete.
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
            self.status_text.value = "✅ Conversion completed successfully!"
            self.status_text.color = ft.Colors.GREEN_700
            self.output_path_text.value = f"Output saved: {output_path.name}"
            self.output_path_text.visible = True
            self.open_folder_btn.visible = True
            self.reset_btn.visible = True
            self.progress_bar.visible = False
            self.page.update()

        except ConversionError as ex:
            # Update UI on error
            self.status_text.value = f"❌ Error: {str(ex)}"
            self.status_text.color = ft.Colors.RED_700
            self.convert_btn.disabled = False
            self.file_path_input.disabled = False
            self.progress_bar.visible = False
            self.page.update()

        except Exception as ex:
            # Update UI on unexpected error
            self.status_text.value = f"❌ Unexpected error: {str(ex)}"
            self.status_text.color = ft.Colors.RED_700
            self.convert_btn.disabled = False
            self.file_path_input.disabled = False
            self.progress_bar.visible = False
            self.page.update()

    def on_open_folder_clicked(self, e):
        """
        Open output folder in file explorer.

        Args:
            e: Button click event
        """
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
                self.status_text.value = f"Could not open folder: {str(ex)}"
                self.status_text.color = ft.Colors.ORANGE_700
                self.page.update()

    def on_reset_clicked(self, e):
        """
        Reset UI for new conversion.

        Args:
            e: Button click event
        """
        self.selected_file = None
        self.output_file = None
        self.selected_file_text.visible = False
        self.status_text.value = "Ready to convert"
        self.status_text.color = None
        self.output_path_text.visible = False
        self.convert_btn.disabled = True
        self.file_path_input.disabled = False
        self.open_folder_btn.visible = False
        self.reset_btn.visible = False
        self.page.update()


def main():
    """Main entry point for the application"""
    def startup(page: ft.Page):
        ConverterApp(page)

    ft.run(startup)


if __name__ == "__main__":
    main()
