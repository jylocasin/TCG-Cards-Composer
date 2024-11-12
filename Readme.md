# PSD to PDF Sheet Assembler

A Python application that takes PSD files and arranges them in a 3x3 grid on A4 sheets, creating a print-ready PDF with proper recto/verso pairing. Perfect for card game prototyping and similar print projects.

## Features

- Arranges cards in a 3x3 grid on A4 sheets
- Handles both recto (front) and verso (back) sides
- Supports multiple sheets automatically
- Includes registration marks and color bars
- Maintains proper bleed areas (2.5mm)
- Provides real-time preview of PSD files
- Supports high-resolution output (up to 600 DPI)
- Creates print-ready PDFs with proper page ordering

## Technical Specifications

- Card Size: 63.5mm × 88.0mm
- Bleed: 2.5mm
- Sheet Size: A4 (210mm × 297mm)
- Grid: 3×3 (9 cards per sheet)
- Supported DPI: 150, 300, 600
- Input Format: Adobe Photoshop (PSD)
- Output Format: PDF

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/psd-to-pdf-assembler.git
cd psd-to-pdf-assembler
```

2. Create a virtual environment (recommended):
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install requirements:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the application:
```bash
python psd-assembler.py
```

2. Select your input files:
   - Click "Select Recto Files" to choose the front side PSD files
   - Click "Select Verso File" to choose the back side PSD file
   - Click "Select Output Directory" to choose where to save the PDF

3. Configure settings:
   - Choose output DPI (150, 300, or 600)
   - Toggle registration marks
   - Toggle color bars
   - Toggle PDF optimization

4. Click "Process Files" to create the PDF

## Code Structure

The application consists of three main Python files:

### psd-assembler.py
Main application file containing the GUI and core logic:
- `PSDAssembler`: Main application class
  - Handles file selection
  - Manages GUI
  - Controls processing flow

### pdf_creator.py
PDF generation engine:
- `PDFCreator`: Core PDF creation class
  - Processes PSD files
  - Creates PDF pages
  - Handles layout and positioning
  - Manages registration marks and color bars

### preview_windows.py
Preview functionality:
- `PreviewWindow`: Single file preview
- `BatchPreviewWindow`: Multiple file preview

## Requirements

- Python 3.8 or higher
- See requirements.txt for Python package dependencies

## Images Requirements

- PSD files should be:
  - In the correct dimensions (63.5mm × 88.0mm + 2.5mm bleed)
  - RGB or CMYK color mode
  - Flattened or with layers (layers will be composited)

## Output

The application creates a multi-page PDF with:
- Recto pages followed by matching verso pages
- 9 cards per page in a 3×3 grid
- Registration marks for proper alignment
- Color bars for print verification
- Proper bleed handling

Example output for different quantities:
- 36 cards = 8 pages (4 recto + 4 verso)
- 54 cards = 12 pages (6 recto + 6 verso)
- 72 cards = 16 pages (8 recto + 8 verso)

## Notes

- For best results, ensure your PSD files include proper bleed areas
- The application automatically handles partial sheets when the card count isn't a multiple of 9
- Progress is shown in real-time during processing
- Preview functionality helps verify correct file selection
- All measurements are in millimeters for print accuracy

## Troubleshooting

Common warnings that can be safely ignored:
- "Unknown image resource"
- "Unknown key: b'CAI'"
- "Unknown tagged block"

These are informational messages from the PSD processing library and don't affect output quality.

## License

[Your chosen license]

## Contributing

[Your contribution guidelines]

## Authors

[Your name/organization]
