# preview_windows.py
import customtkinter as ctk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
from psd_tools import PSDImage
import os
from datetime import datetime

class PreviewWindow(ctk.CTkToplevel):
    def load_psd(self):
        try:
            psd = PSDImage.open(self.psd_path)
            self.psd = psd
            self.image = psd.topil()
            self.update_preview()
            self.update_info()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load PSD: {str(e)}")
            self.destroy()

    def update_preview(self, *args):
        try:
            if self.image is None:
                raise ValueError("No image loaded")

            # Apply zoom
            zoom = int(self.zoom_var.get().rstrip('%')) / 100
            if zoom != 1:
                new_size = (int(self.image.width * zoom), int(self.image.height * zoom))
                display_image = self.image.resize(new_size, Image.Resampling.LANCZOS)
            else:
                display_image = self.image

            # Convert to RGB if necessary
            if display_image.mode != 'RGB':
                display_image = display_image.convert('RGB')

            # Update canvas
            self.photo = ImageTk.PhotoImage(display_image)
            self.canvas.delete("all")
            self.canvas.create_image(0, 0, image=self.photo, anchor="nw")
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        except Exception as e:
            messagebox.showerror("Error", f"Failed to update preview: {str(e)}")

class BatchPreviewWindow(ctk.CTkToplevel):
    def __init__(self, parent, psd_files):
        super().__init__(parent)

        self.title("Batch Preview")
        self.geometry("1000x800")
        self.psd_files = psd_files

        # Make window modal
        self.transient(parent)
        self.grab_set()

        self.setup_ui()
        self.load_previews()

    def setup_ui(self):
        # Main frame
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Controls frame
        self.controls_frame = ctk.CTkFrame(self.main_frame)
        self.controls_frame.pack(fill="x", padx=5, pady=5)

        # Sort options
        sort_label = ctk.CTkLabel(self.controls_frame, text="Sort by:")
        sort_label.pack(side="left", padx=5)

        self.sort_var = ctk.StringVar(value="name")
        sort_menu = ctk.CTkOptionMenu(
            self.controls_frame,
            values=["name", "size", "date"],
            variable=self.sort_var,
            command=self.resort_previews
        )
        sort_menu.pack(side="left", padx=5)

        # Scrollable frame for thumbnails
        self.scroll_frame = ctk.CTkScrollableFrame(
            self.main_frame,
            label_text="PSD Files Preview"
        )
        self.scroll_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Grid frame for thumbnails
        self.grid_frame = ctk.CTkFrame(self.scroll_frame)
        self.grid_frame.pack(fill="both", expand=True)

    def load_previews(self):
        THUMBNAIL_SIZE = (200, 200)
        GRID_COLUMNS = 4

        for i, psd_path in enumerate(self.psd_files):
            try:
                # Create frame for thumbnail
                thumb_frame = ctk.CTkFrame(self.grid_frame)
                row = i // GRID_COLUMNS
                col = i % GRID_COLUMNS
                thumb_frame.grid(row=row, column=col, padx=5, pady=5)

                # Load and create thumbnail
                psd = PSDImage.open(psd_path)
                img = psd.compose()

                # Convert to RGB if necessary
                if img.mode in ['RGBA', 'LA'] or (img.mode == 'P' and 'transparency' in img.info):
                    img = img.convert('RGB')

                # Create thumbnail
                img.thumbnail(THUMBNAIL_SIZE, Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(img)

                # Create image label
                image_label = ttk.Label(thumb_frame, image=photo)
                image_label.image = photo  # Keep reference
                image_label.pack(padx=5, pady=5)

                # Add filename
                name_label = ctk.CTkLabel(
                    thumb_frame,
                    text=os.path.basename(psd_path),
                    wraplength=180
                )
                name_label.pack(padx=5, pady=(0, 5))

                # Add info
                info_text = f"{psd.width}x{psd.height}px\n"
                info_text += f"{psd.color_mode.name}, {len(psd.layers)} layers"
                info_label = ctk.CTkLabel(
                    thumb_frame,
                    text=info_text,
                    font=("Arial", 10)
                )
                info_label.pack(padx=5, pady=(0, 5))

                # Add preview button
                preview_button = ctk.CTkButton(
                    thumb_frame,
                    text="Full Preview",
                    command=lambda p=psd_path: PreviewWindow(self, p)
                )
                preview_button.pack(padx=5, pady=(0, 5))

            except Exception as e:
                # Create error placeholder
                error_frame = ctk.CTkFrame(
                    self.grid_frame,
                    fg_color="red"
                )
                error_frame.grid(row=row, column=col, padx=5, pady=5)

                error_label = ctk.CTkLabel(
                    error_frame,
                    text=f"Error loading\n{os.path.basename(psd_path)}:\n{str(e)}",
                    wraplength=180
                )
                error_label.pack(padx=5, pady=5)

    def resort_previews(self, *args):
        # Clear current previews
        for widget in self.grid_frame.winfo_children():
            widget.destroy()

        # Sort files based on selected criterion
        if self.sort_var.get() == "name":
            sorted_files = sorted(self.psd_files)
        elif self.sort_var.get() == "size":
            sorted_files = sorted(self.psd_files, key=os.path.getsize)
        else:  # date
            sorted_files = sorted(self.psd_files, key=os.path.getmtime)

        # Reload previews with sorted files
        self.psd_files = sorted_files
        self.load_previews()
