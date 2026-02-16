# PDF Layout-Preserving Translation Engine

A deterministic PDF layout-preserving text extraction and translation reinsertion engine in Python.

## Features

- **Extract** text blocks with bounding boxes, font info, and layout metadata
- **Preserve** layout features including page size, margins, column structure
- **Replace** text while maintaining original positioning
- **Auto-scale** fonts when translated text overflows bounding boxes
- **Unicode support** for Hindi, Arabic, CJK and other scripts
- **Multiple pipelines**: Direct PDF, Office roundtrip (DOCX/PPTX/XLSX), XLIFF
- **Auto-detect** source format from PDF metadata

## Installation

```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
# Windows:
.\venv\Scripts\activate
# Linux/Mac:
source venv/bin/activate

# Install core dependencies
pip install PyMuPDF

# Install optional dependencies for Office pipeline
pip install pdf2docx python-pptx openpyxl

# For Office → PDF conversion, install LibreOffice
# https://www.libreoffice.org/
```

## Quick Start

```bash
# Minimal workflow - just 2 commands!
python main.py extract document.pdf -l Spanish
python main.py merge document.pdf
# → Creates document_translated.pdf
```

## CLI Reference

### Commands

| Command | Description |
|---------|-------------|
| `info` | Detect PDF source format (Word/PowerPoint/Excel) |
| `extract` | Extract layout and generate translation file |
| `merge` | Merge translations and rebuild PDF |
| `topdf` | Convert edited Office file to PDF |

### Extract Command

```bash
python main.py extract <input.pdf> [options]
```

**Options:**
| Option | Default | Description |
|--------|---------|-------------|
| `-l, --language` | Hindi | Target language for translation |
| `-p, --pipeline` | direct | Pipeline: `direct`, `office`, `xliff` |
| `--office-format` | auto | Office format: `auto`, `docx`, `pptx`, `xlsx` |
| `--source-language` | en | Source language code for XLIFF |

**Output files:**
- `input.pdf_layout.json` - Layout and block data
- `input.pdf_translate.txt` - Source text for translation
- `input.pdf_translated.txt` - Template for translated text

### Merge Command

```bash
python main.py merge <input.pdf> [output.pdf] [options]
```

**Options:**
| Option | Default | Description |
|--------|---------|-------------|
| `-o, --output` | `{input}_translated.pdf` | Output PDF path |
| `-p, --pipeline` | auto | Auto-detects from layout.json |
| `--min-font-size` | 6.0 | Minimum font size threshold |
| `--font-step` | 0.5 | Font size reduction step |
| `--render-method` | 1 | Text render method (line-by-line) |

**Intermediate files (Office pipeline):**
The Office pipeline automatically keeps intermediate files:
- `input.docx` / `input.pptx` / `input.xlsx` - Original converted from PDF
- `input_translated.docx/pptx/xlsx` - With translations applied

These files can be edited manually and re-converted using `topdf`.

**Alternative input:** Pass the layout JSON directly:
```bash
python main.py merge input.pdf_layout.json
```

### Info Command

```bash
python main.py info document.pdf
```

Detects source format from PDF metadata and recommends pipeline.

### ToPDF Command

```bash
python main.py topdf <office_file> [output.pdf]
```

Converts edited Office files (DOCX/PPTX/XLSX) back to PDF using LibreOffice.

**Use case:** After running `merge`, you may want to manually adjust text boxes in PowerPoint/Word, then regenerate the PDF.

**Workflow for manual text refitting:**
```bash
# 1. Generate translated Office file
python main.py merge presentation.pdf
# → Creates presentation_translated.pptx

# 2. Open in PowerPoint, manually adjust text boxes, fonts, etc.
# 3. Regenerate PDF from edited file
python main.py topdf presentation_translated.pptx
# → Creates presentation_translated.pdf

# Or with custom output name:
python main.py topdf presentation_translated.pptx final_output.pdf
```

## Pipelines

### Direct PDF (default)
Fast, direct PDF text replacement using PyMuPDF.

```bash
python main.py extract doc.pdf -l Hindi
python main.py merge doc.pdf
```

**Pros:** Fast, no external dependencies
**Cons:** Limited text reflow

### Office Roundtrip
Converts PDF → Office format → translates XML → converts back to PDF.

```bash
python main.py extract doc.pdf --pipeline office -l French
python main.py merge doc.pdf
```

Requires LibreOffice for Office → PDF conversion.

**Auto-detects format from metadata:**
- DOCX for Word documents
- PPTX for PowerPoint presentations  
- XLSX for Excel spreadsheets

### XLIFF Export
Generates industry-standard XLIFF for professional CAT tools.

```bash
python main.py extract doc.pdf --pipeline xliff -l German
# Edit input.pdf.xlf in your CAT tool
python main.py merge doc.pdf
```

## Translation File Format

Text is tagged with `<N>text</N>` format where N is a number:

```
<0>Hello World</0>
<1>This is a paragraph.\nWith multiple lines.</1>
<2>Another block</2>
```

- `\n` represents line breaks within a block
- Tags must be preserved exactly during translation
- Do not add/remove lines

## Programmatic API

### Direct PDF Pipeline

```python
from pdf_layout.pipelines import create_direct_pdf_pipeline

# Create pipeline
pipeline = create_direct_pdf_pipeline(
    target_language="Hindi",
    min_font_size=6.0,
    font_step=0.5,
)

# Extract text and layout
from pathlib import Path
result = pipeline.extract(Path("input.pdf"))
print(f"Layout: {result.layout_path}")
print(f"Translate: {result.translate_path}")

# After translation, merge back
pipeline.merge(
    input_path=Path("input.pdf"),
    output_path=Path("output.pdf"),
    translated_path=result.translated_template_path,
    layout_path=result.layout_path,
)
```

### Office Roundtrip Pipeline

```python
from pdf_layout.pipelines import (
    create_office_roundtrip_pipeline,
    OfficeFormat,
)

# Auto-detect format from PDF metadata
pipeline = create_office_roundtrip_pipeline(
    target_language="French",
    office_format=OfficeFormat.AUTO,
    keep_intermediate=True,
)

# Or force specific format
pipeline = create_office_roundtrip_pipeline(
    target_language="Spanish",
    office_format=OfficeFormat.PPTX,
)
```

### Office XML Handlers

Direct access to Office document XML:

```python
from pdf_layout.pipelines import (
    DocxXMLHandler,
    PptxXMLHandler,
    XlsxXMLHandler,
    get_handler,
)
from pathlib import Path

# Auto-detect handler from extension
handler = get_handler(Path("document.docx"))

# Extract text segments
extraction = handler.extract()
for segment in extraction.segments:
    print(f"{segment.block_id}: {segment.text}")

# Update with translations
translations = {"p0": "Translated text", "p1": "More text"}
handler.update(
    output_path=Path("translated.docx"),
    translations=translations,
    extraction=extraction,
)
```

### Source Format Detection

```python
from pdf_layout.source_detector import (
    detect_source_format,
    get_recommended_pipeline,
    SourceFormat,
)
from pathlib import Path

# Detect source format
info = detect_source_format(Path("document.pdf"))
print(f"Format: {info.format.value}")  # docx, pptx, xlsx
print(f"Confidence: {info.confidence:.0%}")
print(f"Creator: {info.creator}")

# Get recommended pipeline
pipeline = get_recommended_pipeline(info)  # "docx", "pptx", "xlsx", or "direct"
```

### Low-Level Extraction

```python
from pdf_layout.extractor import PDFExtractor

with PDFExtractor("input.pdf") as extractor:
    document = extractor.extract()
    
    for page in document.pages:
        print(f"Page {page.page_number}: {page.width}x{page.height}")
        for block in page.blocks:
            print(f"  {block.block_id}: {block.text[:50]}...")
            print(f"    Font: {block.font_name} {block.font_size}pt")
            print(f"    BBox: {block.bbox}")
```

### Low-Level Rebuilding

```python
from pdf_layout.rebuilder_unicode import (
    PDFRebuilder,
    RebuildConfig,
    RenderMethod,
)
from pathlib import Path

config = RebuildConfig(
    render_method=RenderMethod.LINE_BY_LINE,
    min_font_size=6.0,
    font_step=0.5,
)

rebuilder = PDFRebuilder(
    pdf_path=Path("input.pdf"),
    layout_path=Path("layout.json"),
    translations_path=Path("translations.json"),
    config=config,
)

rebuilder.rebuild(Path("output.pdf"))
```

## Output Format

### Layout JSON Structure

```json
{
  "source_file": "input.pdf",
  "pipeline": "direct",
  "target_language": "Hindi",
  "pages": [
    {
      "page_number": 1,
      "width": 595,
      "height": 842,
      "blocks": [
        {
          "block_id": "p1_b0",
          "bbox": [72.0, 72.0, 300.0, 90.0],
          "text": "Hello World",
          "font_name": "Helvetica",
          "font_size": 12.0,
          "color": "#000000",
          "lines": [
            {
              "bbox": [72.0, 72.0, 150.0, 86.0],
              "spans": [{"text": "Hello World", "size": 12.0}]
            }
          ]
        }
      ]
    }
  ],
  "block_order": ["p1_b0", "p1_b1"]
}
```

### Translation File Format

```
<0>Source text block 0</0>
<1>Source text block 1\nWith line break</1>
<2>Source text block 2</2>
```

## Project Structure

```
pdf_layout/
├── __init__.py
├── extractor.py           # PDF text extraction
├── segmenter.py           # Block segmentation
├── rebuilder.py           # Legacy rebuilder
├── rebuilder_unicode.py   # Unicode-aware rebuilder
├── font_utils.py          # Font mapping/metrics
├── translation_io.py      # Translation file I/O
├── source_detector.py     # PDF source format detection
└── pipelines/
    ├── __init__.py
    ├── base.py            # Pipeline base classes
    ├── direct_pdf.py      # Direct PDF pipeline
    ├── office_roundtrip.py # Office conversion pipeline
    ├── office_xml.py      # Office XML handlers
    ├── docx_roundtrip.py  # Legacy DOCX pipeline
    └── xliff_format.py    # XLIFF export pipeline
tests/
├── test_roundtrip.py      # Comprehensive tests
main.py                    # CLI entry point
```

## Dependencies

| Package | Purpose | Required |
|---------|---------|----------|
| PyMuPDF | PDF processing | ✅ |
| pdf2docx | PDF → DOCX conversion | Office pipeline |
| python-pptx | PPTX handling | Office pipeline |
| openpyxl | XLSX handling | Office pipeline |
| LibreOffice | Office → PDF conversion | Office pipeline |

## Testing

```bash
# Run all tests
python -m pytest tests/ -v

# Run with coverage
python -m pytest tests/ --cov=pdf_layout
```

## Limitations

- Direct pipeline: Limited text reflow for length changes
- Office pipeline: Requires LibreOffice for PDF generation
- Does not handle embedded fonts (uses Nirmala UI for Unicode)
- Does not reconstruct SmartArt or vector graphics
- Table text replaced within cell bounding boxes

## License

MIT License
