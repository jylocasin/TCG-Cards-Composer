# pdf_creator.py
import os
import tempfile
import shutil
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from PIL import Image
from psd_tools import PSDImage
from PyPDF2 import PdfMerger
import math

class PDFCreator:
    def __init__(self, width_mm=None, height_mm=None):
        """
        Initialize PDFCreator with default A4 dimensions
        """
        self.width_mm = width_mm if width_mm is not None else 210  # A4 default
        self.height_mm = height_mm if height_mm is not None else 297  # A4 default
        self.temp_dir = tempfile.mkdtemp()
        self.optimize = True

    def set_optimization(self, optimize):
        """Set PDF optimization flag"""
        self.optimize = optimize

    def process_batch(self, recto_files, verso_file, output_path, card_width=63.5,
                     card_height=88.0, bleed=2.5, dpi=300, reg_marks=True,
                     color_bars=True, progress_callback=None):
        """Process a batch of files creating multiple sheets"""
        try:
            # Convert tuple to list if necessary
            recto_files = list(recto_files)

            # Fixed 3x3 grid (9 cards per sheet)
            cards_per_sheet = 9
            total_sheets = math.ceil(len(recto_files) / cards_per_sheet)

            print(f"\nProcessing {len(recto_files)} cards across {total_sheets} sheets")
            generated_pdfs = []

            # Process each sheet
            for sheet_num in range(total_sheets):
                if progress_callback:
                    progress = (sheet_num) / total_sheets
                    message = f"Processing sheet {sheet_num + 1} of {total_sheets}"
                    progress_callback(progress, message)

                # Calculate the range of cards for this sheet
                start_idx = sheet_num * cards_per_sheet
                end_idx = min(start_idx + cards_per_sheet, len(recto_files))
                current_recto_files = recto_files[start_idx:end_idx]

                print(f"\nSheet {sheet_num + 1} of {total_sheets}:")
                print(f"Processing cards {start_idx + 1} to {end_idx}")

                # Create recto sheet
                recto_pdf = os.path.join(self.temp_dir, f"sheet_{sheet_num:03d}_recto.pdf")
                self.create_sheet(
                    current_recto_files,
                    recto_pdf,
                    card_width=card_width,
                    card_height=card_height,
                    bleed=bleed,
                    dpi=dpi,
                    is_verso=False,
                    reg_marks=reg_marks,
                    color_bars=color_bars
                )
                generated_pdfs.append(recto_pdf)

                # Create verso sheet
                verso_pdf = os.path.join(self.temp_dir, f"sheet_{sheet_num:03d}_verso.pdf")
                verso_files = [verso_file] * len(current_recto_files)
                self.create_sheet(
                    verso_files,
                    verso_pdf,
                    card_width=card_width,
                    card_height=card_height,
                    bleed=bleed,
                    dpi=dpi,
                    is_verso=True,
                    reg_marks=reg_marks,
                    color_bars=color_bars
                )
                generated_pdfs.append(verso_pdf)

            if progress_callback:
                progress_callback(0.9, "Merging PDFs...")

            # Merge all PDFs
            self.merge_pdfs(generated_pdfs, output_path)

            if progress_callback:
                progress_callback(1.0, f"Complete! Created {total_sheets} sheets ({total_sheets*2} pages)")

        except Exception as e:
            print(f"Error in process_batch: {str(e)}")
            raise

    def create_sheet(self, sheet_files, output_path, card_width=63.5, card_height=88.0,
                    bleed=2.5, dpi=300, is_verso=False, reg_marks=True, color_bars=True):
        """Create a single sheet of cards in a 3x3 grid"""
        try:
            # Fixed 3x3 grid size
            grid_size = 3

            # Calculate grid dimensions
            total_grid_width = grid_size * (card_width + 2 * bleed)
            total_grid_height = grid_size * (card_height + 2 * bleed)
            margin_x = (self.width_mm - total_grid_width) / 2
            margin_y = (self.height_mm - total_grid_height) / 2

            # Create new PDF document
            c = canvas.Canvas(output_path, pagesize=A4)

            # Process each card position (maximum 9 cards)
            for i, psd_file in enumerate(sheet_files[:grid_size * grid_size]):
                row = i // grid_size
                col = i % grid_size

                if is_verso:
                    # Mirror positions for verso side
                    x = self.width_mm - (margin_x + (col * (card_width + 2 * bleed)) + (card_width + 2 * bleed))
                    y = self.height_mm - (margin_y + (row * (card_height + 2 * bleed)) + (card_height + 2 * bleed))
                else:
                    x = margin_x + (col * (card_width + 2 * bleed))
                    y = self.height_mm - (margin_y + (row * (card_height + 2 * bleed)) + (card_height + 2 * bleed))

                print(f"  Placing card {i+1} at position ({row+1}, {col+1})")

                # Process and place image
                image = self.handle_psd_file(psd_file, dpi)
                self.place_image(
                    c, image, x, y,
                    card_width + 2 * bleed,
                    card_height + 2 * bleed,
                    dpi
                )

            # Add cut lines and marks
            self.add_cut_lines(
                c, margin_x, margin_y,
                total_grid_width, total_grid_height,
                card_width, card_height,
                grid_size, bleed
            )

            if reg_marks:
                self.add_registration_marks(
                    c, margin_x, margin_y,
                    total_grid_width, total_grid_height
                )

            if color_bars:
                self.add_color_bars(
                    c, margin_x, margin_y - 10,
                    total_grid_width
                )

            c.showPage()
            c.save()
            return output_path

        except Exception as e:
            print(f"Error creating sheet: {str(e)}")
            raise

    def handle_psd_file(self, psd_path, target_dpi):
        """Process a PSD file and return a PIL Image"""
        try:
            psd = PSDImage.open(psd_path)
            image = psd.topil()

            if image is None:
                raise ValueError(f"Could not process PSD file: {psd_path}")

            # Convert to RGB if necessary
            if image.mode in ['RGBA', 'LA'] or (image.mode == 'P' and 'transparency' in image.info):
                background = Image.new('RGB', image.size, (255, 255, 255))
                if image.mode == 'RGBA':
                    background.paste(image, mask=image.split()[3])
                else:
                    background.paste(image, mask=image.split()[1])
                image = background
            elif image.mode == 'CMYK':
                image = image.convert('RGB')

            return image

        except Exception as e:
            raise ValueError(f"Error processing {os.path.basename(psd_path)}: {str(e)}")

    def place_image(self, canvas, image, x, y, width, height, target_dpi):
        """Place an image on the PDF canvas with proper scaling"""
        try:
            # Calculate target size in pixels
            target_width = int(width * target_dpi / 25.4)  # mm to inches * dpi
            target_height = int(height * target_dpi / 25.4)

            # Create temporary file
            with tempfile.NamedTemporaryFile(suffix='.jpg', delete=False) as tmp_file:
                # Resize image if needed
                if image.size != (target_width, target_height):
                    image = image.resize(
                        (target_width, target_height),
                        Image.Resampling.LANCZOS
                    )

                # Save with appropriate quality
                quality = 95 if self.optimize else 100
                image.save(
                    tmp_file.name,
                    'JPEG',
                    quality=quality,
                    dpi=(target_dpi, target_dpi)
                )

                # Place image
                canvas.drawImage(
                    tmp_file.name,
                    x * mm,
                    y * mm,
                    width=width * mm,
                    height=height * mm
                )

            # Clean up
            os.unlink(tmp_file.name)

        except Exception as e:
            raise Exception(f"Error placing image: {str(e)}")

    def add_cut_lines(self, canvas, margin_x, margin_y, grid_width, grid_height,
                     card_width, card_height, grid_size, bleed):
        """Add cut lines to the PDF page"""
        # Set properties for trim lines
        canvas.setStrokeColorRGB(0, 0, 0)
        canvas.setLineWidth(0.25)
        canvas.setDash([])  # Solid line

        # Draw trim lines
        for i in range(grid_size + 1):
            # Vertical lines
            x = margin_x + (i * (card_width + 2 * bleed))
            canvas.line(
                x * mm, margin_y * mm,
                x * mm, (margin_y + grid_height) * mm
            )

            # Horizontal lines
            y = margin_y + (i * (card_height + 2 * bleed))
            canvas.line(
                margin_x * mm, y * mm,
                (margin_x + grid_width) * mm, y * mm
            )

        # Draw bleed lines (dashed)
        canvas.setDash([2, 2])
        canvas.setLineWidth(0.15)
        canvas.setStrokeColorRGB(0.5, 0.5, 0.5)  # Gray color for bleed lines

        for i in range(1, grid_size):
            # Vertical bleed lines
            x = margin_x + (i * (card_width + 2 * bleed))
            canvas.line(
                (x - bleed) * mm, margin_y * mm,
                (x - bleed) * mm, (margin_y + grid_height) * mm
            )
            canvas.line(
                (x + bleed) * mm, margin_y * mm,
                (x + bleed) * mm, (margin_y + grid_height) * mm
            )

            # Horizontal bleed lines
            y = margin_y + (i * (card_height + 2 * bleed))
            canvas.line(
                margin_x * mm, (y - bleed) * mm,
                (margin_x + grid_width) * mm, (y - bleed) * mm
            )
            canvas.line(
                margin_x * mm, (y + bleed) * mm,
                (margin_x + grid_width) * mm, (y + bleed) * mm
            )

    def add_registration_marks(self, canvas, margin_x, margin_y, grid_width, grid_height):
        """Add registration marks to the PDF page"""
        canvas.setDash([])  # Solid line
        canvas.setLineWidth(0.25)
        canvas.setStrokeColorRGB(0, 0, 0)

        reg_size = 5  # 5mm registration marks

        def draw_reg_mark(x, y):
            # Draw cross
            canvas.line(
                (x - reg_size) * mm, y * mm,
                (x + reg_size) * mm, y * mm
            )
            canvas.line(
                x * mm, (y - reg_size) * mm,
                x * mm, (y + reg_size) * mm
            )
            # Add circle for precise alignment
            canvas.circle(x * mm, y * mm, 0.5 * mm, stroke=1, fill=0)

        # Corner registration marks
        corners = [
            (margin_x, margin_y),  # Bottom left
            (margin_x + grid_width, margin_y),  # Bottom right
            (margin_x, margin_y + grid_height),  # Top left
            (margin_x + grid_width, margin_y + grid_height)  # Top right
        ]

        for x, y in corners:
            draw_reg_mark(x, y)

        # Center registration marks
        center_x = margin_x + (grid_width / 2)
        center_y = margin_y + (grid_height / 2)

        draw_reg_mark(center_x, margin_y)  # Bottom
        draw_reg_mark(center_x, margin_y + grid_height)  # Top
        draw_reg_mark(margin_x, center_y)  # Left
        draw_reg_mark(margin_x + grid_width, center_y)  # Right

    def add_color_bars(self, canvas, x, y, width):
        """Add CMYK color bars for print verification"""
        bar_height = 5 * mm
        bar_width = width / 4

        # Reset line properties
        canvas.setDash([])
        canvas.setLineWidth(0.25)

        # Cyan
        canvas.setFillColorCMYK(1, 0, 0, 0)
        canvas.rect(x * mm, y * mm, bar_width * mm, bar_height, fill=1)

        # Magenta
        canvas.setFillColorCMYK(0, 1, 0, 0)
        canvas.rect((x + bar_width) * mm, y * mm, bar_width * mm, bar_height, fill=1)

        # Yellow
        canvas.setFillColorCMYK(0, 0, 1, 0)
        canvas.rect((x + bar_width * 2) * mm, y * mm, bar_width * mm, bar_height, fill=1)

        # Black
        canvas.setFillColorCMYK(0, 0, 0, 1)
        canvas.rect((x + bar_width * 3) * mm, y * mm, bar_width * mm, bar_height, fill=1)

        # Reset fill color
        canvas.setFillColorRGB(0, 0, 0)


    def merge_pdfs(self, pdf_files, output_path):
        """Merge multiple PDF files into one"""
        try:
            merger = PdfMerger()

            # Sort files to ensure correct order
            sorted_files = sorted(pdf_files)

            # Add each PDF file to the merger
            for pdf_file in sorted_files:
                if os.path.exists(pdf_file):
                    merger.append(pdf_file)
                else:
                    print(f"Warning: PDF file not found: {pdf_file}")

            # Write the merged file
            with open(output_path, 'wb') as output_file:
                merger.write(output_file)

            merger.close()

        except Exception as e:
            raise Exception(f"Error merging PDFs: {str(e)}")

    def cleanup(self):
        """Clean up temporary files"""
        try:
            shutil.rmtree(self.temp_dir)
        except Exception as e:
            print(f"Warning: Could not clean up temporary directory: {str(e)}")

class PDFHelper:
    @staticmethod
    def validate_and_resize_psd(psd_path, expected_width_mm, expected_height_mm, tolerance_percent=5):
        """
        Validate PSD file dimensions and resize if needed
        Args:
            psd_path: Path to PSD file
            expected_width_mm: Expected width in millimeters (including bleed)
            expected_height_mm: Expected height in millimeters (including bleed)
            tolerance_percent: Acceptable deviation percentage
        Returns:
            PIL Image: Resized image if needed
        """
        try:
            psd = PSDImage.open(psd_path)
            # Convert PSD to PIL Image
            composed_image = psd.topil()

            if composed_image is None:
                raise ValueError(f"Could not open PSD file: {psd_path}")

            # Get current dimensions in mm
            width_mm = psd.width / 11.811  # Convert pixels to mm (300 DPI reference)
            height_mm = psd.height / 11.811

            # Calculate tolerance in mm
            tolerance_mm = min(expected_width_mm, expected_height_mm) * (tolerance_percent / 100)

            # Check if dimensions need adjustment
            width_diff = abs(width_mm - expected_width_mm)
            height_diff = abs(height_mm - expected_height_mm)

            if width_diff > tolerance_mm or height_diff > tolerance_mm:
                # Calculate target dimensions at 300 DPI
                target_width = int(expected_width_mm * 11.811)
                target_height = int(expected_height_mm * 11.811)

                # Convert to RGB if necessary
                if composed_image.mode in ['RGBA', 'LA'] or (composed_image.mode == 'P' and 'transparency' in composed_image.info):
                    composed_image = composed_image.convert('RGB')

                # Resize image
                resized_image = composed_image.resize(
                    (target_width, target_height),
                    Image.Resampling.LANCZOS
                )

                print(f"Resized {os.path.basename(psd_path)} from "
                      f"{width_mm:.1f}x{height_mm:.1f}mm to "
                      f"{expected_width_mm:.1f}x{expected_height_mm:.1f}mm")

                return resized_image

            return composed_image

        except Exception as e:
            raise ValueError(f"Error processing {os.path.basename(psd_path)}: {str(e)}")

    @staticmethod
    def process_psd(psd_path):
        """Process a PSD file and return a PIL Image"""
        try:
            psd = PSDImage.open(psd_path)
            image = psd.topil()

            if image is None:
                raise ValueError(f"Could not process PSD file: {psd_path}")

            # Convert to RGB if necessary
            if image.mode in ['RGBA', 'LA'] or (image.mode == 'P' and 'transparency' in image.info):
                # Create white background for transparency
                background = Image.new('RGB', image.size, (255, 255, 255))
                if image.mode == 'RGBA':
                    background.paste(image, mask=image.split()[3])
                else:
                    background.paste(image, mask=image.split()[1])
                image = background
            elif image.mode == 'CMYK':
                image = image.convert('RGB')

            return image

        except Exception as e:
            raise ValueError(f"Error processing {os.path.basename(psd_path)}: {str(e)}")