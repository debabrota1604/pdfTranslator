# PDF Layout-Preserving Segmentation and Reinsertion Engine

A deterministic PDF layout-preserving text extraction and reinsertion engine in Python.

## Features

- **Extract** text blocks with bounding boxes, font info, and layout metadata
- **Preserve** layout features including page size, margins, column structure
- **Replace** text while maintaining original positioning
- **Auto-scale** fonts when translated text overflows bounding boxes
- **Deterministic** output for reproducible results

## Installation

```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
# Windows:
.\venv\Scripts\activate
# Linux/Mac:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

## Usage

### CLI Commands

#### Extract Layout
```bash
python main.py extract input.pdf layout.json
```

Extract text layout from a PDF file and save to JSON.

Options:
- `--template, -t`: Also create a translation template file

#### Rebuild PDF
```bash
python main.py rebuild input.pdf layout.json translations.json output.pdf
```

Rebuild PDF with translated text.

Options:
- `--min-font-size`: Minimum font size threshold (default: 6.0)
- `--font-step`: Font size reduction step (default: 0.5)

#### Full Pipeline
```bash
python main.py process input.pdf layout.json translations.json output.pdf
```

Extract layout and rebuild with translations in one command.

#### Create Translation Template
```bash
python main.py template layout.json template.json
```

Create a translation template from existing layout JSON.

### Programmatic Usage

```python
from pdf_layout.extractor import extract_pdf_layout
from pdf_layout.rebuilder import rebuild_pdf
import json

# Step 1: Extract layout
document = extract_pdf_layout("input.pdf", "layout.json")

# Step 2: Create translations (block_id -> translated text)
translations = {}
for page in document.pages:
    for block in page.blocks:
        translations[block.block_id] = translate(block.text)  # Your translation function

# Save translations
with open("translations.json", "w") as f:
    json.dump(translations, f, indent=2)

# Step 3: Rebuild PDF
rebuild_pdf(
    pdf_path="input.pdf",
    layout_path="layout.json",
    translations_path="translations.json",
    output_path="output.pdf"
)
```

## Output Format

### Layout JSON Structure
```json
{
  "source_file": "input.pdf",
  "pages": [
    {
      "page_number": 1,
      "width": 595,
      "height": 842,
      "rotation": 0,
      "blocks": [
        {
          "block_id": "p1_b0",
          "bbox": [72.0, 72.0, 300.0, 90.0],
          "text": "Hello World",
          "font_name": "Helvetica",
          "font_size": 12.0,
          "color": "#000000",
          "writing_direction": "ltr",
          "line_height": 14.0
        }
      ]
    }
  ]
}
```

### Translations JSON Structure
```json
{
  "p1_b0": "Translated text for block 0",
  "p1_b1": "Translated text for block 1"
}
```

## Project Structure

```
pdf_translator/
├── pdf_layout/
│   ├── __init__.py
│   ├── extractor.py      # PDF text extraction
│   ├── segmenter.py      # Block segmentation logic
│   └── rebuilder.py      # PDF rebuilding with translations
├── utils/
│   ├── __init__.py
│   └── font_utils.py     # Font mapping and metrics
├── tests/
│   ├── __init__.py
│   └── test_roundtrip.py # Comprehensive tests
├── main.py               # CLI entry point
├── requirements.txt
└── README.md
```

## Determinism Guarantees

- **Stable block IDs**: Block IDs are deterministic based on position (sorted top-to-bottom, left-to-right)
- **Consistent output**: Same input + translations = same output PDF
- **No random seeds**: All operations are fully deterministic
- **Rounded coordinates**: Floating point values rounded to 3 decimals

## Font Handling

- Extracts original font names from PDF
- Maps to PyMuPDF built-in fonts (Helvetica, Times, Courier)
- Falls back to Helvetica for unknown fonts
- Preserves font size and color

## Font Scaling Algorithm

When translated text overflows the bounding box:

1. Start with original font size
2. Measure text height using invisible rendering
3. If overflow: reduce font size by 0.5pt
4. Repeat until text fits or minimum size (6pt) reached

## Testing

```bash
# Run all tests
python -m pytest tests/ -v

# Run with coverage
python -m pytest tests/ --cov=pdf_layout --cov=utils
```

## Dependencies

- **PyMuPDF** (fitz): PDF processing
- **ReportLab**: PDF generation (optional, for advanced rendering)
- **pycairo**: Cairo rendering fallback (optional)
- **pytest**: Testing framework

## Limitations

- Does not handle embedded fonts (uses built-in fallbacks)
- Does not reconstruct SmartArt or vector graphics
- Does not reflow tables (text replaced within cell bounding boxes)
- Writing direction detection is basic (ltr/rtl/ttb)

## License

MIT License
