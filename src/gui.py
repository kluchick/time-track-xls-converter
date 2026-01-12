"""
GUI for Excel Time Tracker Converter using Flet.

Provides a drag-and-drop interface for converting Excel files.
"""

import asyncio
import sys
from pathlib import Path
from tkinter import filedialog
import tkinter as tk

import flet as ft

# Add src directory to path for imports
sys.path.insert(0, str(Path(__file__).parent))
from converter import convert_excel


async def main(page: ft.Page):
    """Main GUI function."""
    # Configure page
    page.title = "Excel Time Tracker Converter"
    page.theme_mode = ft.ThemeMode.DARK
    page.window.width = 600
    page.window.height = 300
    page.window.resizable = False
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    
    # Status text
    status_text = ft.Text(
        "Enter Excel file path or click 'Browse' to select file",
        size=14,
        color=ft.Colors.GREY_400,
        text_align=ft.TextAlign.CENTER,
    )
    
    # File path input
    file_path_input = ft.TextField(
        label="Excel File Path",
        hint_text="C:\\path\\to\\file.xlsx",
        width=500,
        border_color=ft.Colors.GREY_700,
    )
    
    def browse_file(e):
        """Open file dialog to select Excel file."""
        # Hide tkinter root window
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        
        # Open file dialog
        file_path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        root.destroy()
        
        if file_path:
            file_path_input.value = file_path
            page.update()
    
    def process_file_from_input(e):
        """Process file from text input."""
        file_path_str = file_path_input.value.strip()
        if file_path_str:
            file_path = Path(file_path_str)
            if file_path.exists() and file_path.suffix.lower() == '.xlsx':
                page.run_task(process_file, file_path)
            else:
                status_text.value = f"Error: File not found or not an Excel file"
                status_text.color = ft.Colors.RED_400
                page.update()
        else:
            status_text.value = "Please enter a file path"
            status_text.color = ft.Colors.RED_400
            page.update()
    
    # Browse button
    browse_button = ft.Button(
        "📂 Browse",
        on_click=browse_file,
    )
    
    # Convert button
    convert_button = ft.Button(
        "Convert",
        on_click=process_file_from_input,
    )
    
    async def process_file(file_path: Path):
        """Process the selected Excel file."""
        try:
            # Update UI
            status_text.value = "Processing..."
            status_text.color = ft.Colors.BLUE_400
            page.update()
            
            # Convert file (run in executor to avoid blocking UI)
            loop = asyncio.get_event_loop()
            output_path = await loop.run_in_executor(
                None, convert_excel, file_path
            )
            
            # Success
            status_text.value = f"Success! Output saved to:\n{output_path.name}"
            status_text.color = ft.Colors.GREEN_400
            drop_container.border = ft.Border.all(2, ft.Colors.GREEN_400)
            page.update()
            
        except Exception as ex:
            # Error
            status_text.value = f"Error: {str(ex)}"
            status_text.color = ft.Colors.RED_400
            page.update()
    
    # Layout
    page.add(
        ft.Column(
            [
                ft.Text(
                    "Excel Time Tracker Converter",
                    size=24,
                    weight=ft.FontWeight.BOLD,
                    text_align=ft.TextAlign.CENTER,
                ),
                status_text,
                file_path_input,
                ft.Row(
                    [browse_button, convert_button],
                    alignment=ft.MainAxisAlignment.CENTER,
                    spacing=10,
                ),
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=20,
        )
    )


if __name__ == "__main__":
    ft.run(main)
