# main.py
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
from psd_tools import PSDImage
import os
import math
import tempfile
import shutil
from threading import Thread
from datetime import datetime
from pathlib import Path
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from PyPDF2 import PdfMerger

# Import from other parts
from preview_windows import PreviewWindow, BatchPreviewWindow
from pdf_creator import PDFCreator, PDFHelper
from psd_tools import PSDImage

class PSDAssembler(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Constants
        self.CARD_WIDTH = 63.5  # mm
        self.CARD_HEIGHT = 88.0  # mm
        self.BLEED = 2.5  # mm
        self.A4_WIDTH = 210  # mm
        self.A4_HEIGHT = 297  # mm
        self.CARDS_PER_SHEET = 9  # Fixed 3x3 grid

        # Initialize variables
        self.recto_files = []
        self.verso_file = None
        self.output_directory = None
        self.processing_error = None

        # Set up GUI
        self.setup_window()
        self.setup_gui()

    def setup_window(self):
        """Initialize the main window"""
        self.title("PSD to PDF Sheet Assembler")
        self.geometry("800x700")
        self.resizable(True, True)

        # Configure grid
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

    def setup_gui(self):
        """Set up the main GUI elements"""
        # Main frame
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)

        # Title
        title_label = ctk.CTkLabel(
            self.main_frame,
            text="PSD to PDF Sheet Assembler",
            font=("Arial", 20, "bold")
        )
        title_label.pack(pady=(0, 20))

        # Setup frames
        self.setup_file_selection()
        self.setup_settings()
        self.setup_processing_controls()

    def setup_file_selection(self):
        """Set up the file selection section"""
        file_frame = ctk.CTkFrame(self.main_frame)
        file_frame.pack(fill="x", padx=20, pady=10)

        # Recto files selection
        recto_frame = ctk.CTkFrame(file_frame)
        recto_frame.pack(fill="x", pady=5)

        recto_button = ctk.CTkButton(
            recto_frame,
            text="Select Recto Files",
            command=self.select_recto_files
        )
        recto_button.pack(side="left", padx=5)

        self.recto_preview_button = ctk.CTkButton(
            recto_frame,
            text="Preview Files",
            command=self.preview_recto_files,
            state="disabled"
        )
        self.recto_preview_button.pack(side="left", padx=5)

        self.recto_label = ctk.CTkLabel(recto_frame, text="No files selected")
        self.recto_label.pack(side="left", padx=5)

        # Verso file selection
        verso_frame = ctk.CTkFrame(file_frame)
        verso_frame.pack(fill="x", pady=5)

        verso_button = ctk.CTkButton(
            verso_frame,
            text="Select Verso File",
            command=self.select_verso_file
        )
        verso_button.pack(side="left", padx=5)

        self.verso_preview_button = ctk.CTkButton(
            verso_frame,
            text="Preview File",
            command=self.preview_verso_file,
            state="disabled"
        )
        self.verso_preview_button.pack(side="left", padx=5)

        self.verso_label = ctk.CTkLabel(verso_frame, text="No file selected")
        self.verso_label.pack(side="left", padx=5)

        # Output directory selection
        output_frame = ctk.CTkFrame(file_frame)
        output_frame.pack(fill="x", pady=5)

        output_button = ctk.CTkButton(
            output_frame,
            text="Select Output Directory",
            command=self.select_output_directory
        )
        output_button.pack(side="left", padx=5)

        self.output_label = ctk.CTkLabel(output_frame, text="No directory selected")
        self.output_label.pack(side="left", padx=5)

    def setup_settings(self):
        """Set up the settings section"""
        settings_frame = ctk.CTkFrame(self.main_frame)
        settings_frame.pack(fill="x", padx=20, pady=10)

        # Settings title
        settings_label = ctk.CTkLabel(
            settings_frame,
            text="Settings",
            font=("Arial", 12, "bold")
        )
        settings_label.pack(pady=5)

        # DPI selection
        dpi_frame = ctk.CTkFrame(settings_frame)
        dpi_frame.pack(fill="x", pady=5)

        dpi_label = ctk.CTkLabel(dpi_frame, text="Output DPI:")
        dpi_label.pack(side="left", padx=5)

        self.dpi_var = ctk.StringVar(value="300")
        dpi_menu = ctk.CTkOptionMenu(
            dpi_frame,
            values=["150", "300", "600"],
            variable=self.dpi_var
        )
        dpi_menu.pack(side="left", padx=5)

        # Print options
        self.reg_marks_var = ctk.BooleanVar(value=True)
        reg_marks_cb = ctk.CTkCheckBox(
            settings_frame,
            text="Add Registration Marks",
            variable=self.reg_marks_var
        )
        reg_marks_cb.pack(pady=5)

        self.color_bars_var = ctk.BooleanVar(value=True)
        color_bars_cb = ctk.CTkCheckBox(
            settings_frame,
            text="Add Color Bars",
            variable=self.color_bars_var
        )
        color_bars_cb.pack(pady=5)

        self.optimize_var = ctk.BooleanVar(value=True)
        optimize_cb = ctk.CTkCheckBox(
            settings_frame,
            text="Optimize PDF Size",
            variable=self.optimize_var
        )
        optimize_cb.pack(pady=5)

    def setup_processing_controls(self):
        """Set up the processing controls section"""
        # Process button
        self.process_button = ctk.CTkButton(
            self.main_frame,
            text="Process Files",
            command=self.start_processing,
            state="disabled"
        )
        self.process_button.pack(pady=20)

        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(self.main_frame)
        self.progress_bar.set(0)
        self.progress_bar.pack(pady=10, fill="x", padx=20)
        self.progress_bar.pack_forget()

        # Status label
        self.status_label = ctk.CTkLabel(
            self.main_frame,
            text=""
        )
        self.status_label.pack(pady=5)

    def select_recto_files(self):
        """Handle recto file selection"""
        files = filedialog.askopenfilenames(
            title="Select Recto PSD Files",
            filetypes=[("PSD files", "*.psd"), ("All files", "*.*")]
        )

        if files:
            try:
                self.recto_files = files
                self.recto_label.configure(text=f"Selected {len(files)} recto files")
                self.recto_preview_button.configure(state="normal")
                self.update_process_button()

            except Exception as e:
                messagebox.showerror("Error", str(e))

    def select_verso_file(self):
        """Handle verso file selection"""
        file = filedialog.askopenfilename(
            title="Select Verso File",
            filetypes=[("PSD files", "*.psd"), ("All files", "*.*")]
        )

        if file:
            try:
                self.verso_file = file
                self.verso_label.configure(text=f"Selected: {Path(file).name}")
                self.verso_preview_button.configure(state="normal")
                self.update_process_button()

            except Exception as e:
                messagebox.showerror("Error", str(e))

    def select_output_directory(self):
        """Handle output directory selection"""
        directory = filedialog.askdirectory(title="Select Output Directory")

        if directory:
            if os.access(directory, os.W_OK):
                self.output_directory = directory
                self.output_label.configure(text=f"Output: {Path(directory).name}")
                self.update_process_button()
            else:
                messagebox.showerror("Error", "Selected directory is not writable")

    def preview_recto_files(self):
        """Show batch preview for recto files"""
        if self.recto_files:
            BatchPreviewWindow(self, self.recto_files)

    def preview_verso_file(self):
        """Show preview for verso file"""
        if self.verso_file:
            PreviewWindow(self, self.verso_file, "Verso Preview")

    def update_process_button(self):
        """Update process button state and show warnings if needed"""
        if self.recto_files and self.verso_file and self.output_directory:
            # Calculate needed sheets
            total_sheets = math.ceil(len(self.recto_files) / self.CARDS_PER_SHEET)

            if len(self.recto_files) % self.CARDS_PER_SHEET != 0:
                warning = f"Warning: Number of recto files ({len(self.recto_files)}) "
                warning += f"is not a multiple of {self.CARDS_PER_SHEET}. "
                warning += "The last sheet will be partially empty."

                self.status_label.configure(
                    text=warning,
                    text_color="orange"
                )
            else:
                self.status_label.configure(
                    text=f"Ready to process {len(self.recto_files)} cards across {total_sheets} sheets",
                    text_color="white"
                )

            self.process_button.configure(state="normal")
        else:
            self.process_button.configure(state="disabled")
            self.status_label.configure(
                text="Please select all required files and output directory",
                text_color="white"
            )

    def start_processing(self):
        """Start the processing operation"""
        # Show progress bar
        self.progress_bar.pack(pady=10, fill="x", padx=20)
        self.progress_bar.set(0)

        # Disable controls during processing
        self.process_button.configure(state="disabled")

        # Reset error state
        self.processing_error = None

        # Start processing thread
        self.processing_thread = Thread(target=self.process_files)
        self.processing_thread.start()

        # Start monitoring
        self.monitor_processing()

    def monitor_processing(self):
        """Monitor the processing thread"""
        if self.processing_thread.is_alive():
            self.after(100, self.monitor_processing)
        else:
            # Hide progress bar
            self.progress_bar.pack_forget()

            # Check for errors
            if self.processing_error:
                # Show error in main thread
                self.after(0, lambda: messagebox.showerror(
                    "Error",
                    f"An error occurred: {str(self.processing_error)}"
                ))
                self.status_label.configure(
                    text=f"Error: {str(self.processing_error)}",
                    text_color="red"
                )
            else:
                # Show completion message
                self.after(0, lambda: messagebox.showinfo(
                    "Complete",
                    "Processing completed successfully!"
                ))
                self.status_label.configure(
                    text="Processing completed successfully!",
                    text_color="white"
                )

            # Re-enable controls
            self.process_button.configure(state="normal")

    def update_status(self, message, color="white"):
        """Thread-safe status update"""
        self.after(0, lambda: self.status_label.configure(
            text=message,
            text_color=color
        ))

    def update_progress(self, value):
        """Thread-safe progress update"""
        self.after(0, lambda: self.progress_bar.set(value))

    def process_files(self):
        """Process the PSD files and create PDF output"""
        try:
            pdf_creator = PDFCreator()
            pdf_creator.set_optimization(self.optimize_var.get())

            # Create output PDF
            output_pdf = os.path.join(self.output_directory, "cards.pdf")

            # Process all files in batch
            pdf_creator.process_batch(
                recto_files=self.recto_files,
                verso_file=self.verso_file,
                output_path=output_pdf,
                card_width=self.CARD_WIDTH,
                card_height=self.CARD_HEIGHT,
                bleed=self.BLEED,
                dpi=int(self.dpi_var.get()),
                reg_marks=self.reg_marks_var.get(),
                color_bars=self.color_bars_var.get(),
                progress_callback=lambda progress, message: (
                    self.update_progress(progress),
                    self.update_status(message)
                )
            )

        except Exception as e:
            self.processing_error = str(e)

def main():
    """Main entry point of the application"""
    try:
        # Set appearance mode and color theme
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Create and run the application
        app = PSDAssembler()
        app.mainloop()

    except Exception as e:
        # Show error message if something goes wrong during startup
        messagebox.showerror(
            "Fatal Error",
            f"An unexpected error occurred: {str(e)}\n\n"
            "Please check that all required libraries are installed:\n"
            "- customtkinter\n"
            "- Pillow\n"
            "- psd-tools\n"
            "- reportlab\n"
            "- PyPDF2"
        )

if __name__ == "__main__":
    main()
