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

# Install optional dependencies for CAT pipeline (OASIS XLIFF compliance)
pip install translate-toolkit

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
| `-p, --pipeline` | cat | Pipeline: `direct`, `office`, `xliff`, `cat`, `pikepdf`, `html` |
| `--office-format` | auto | Office format: `auto`, `docx`, `pptx`, `xlsx` |
| `--cat-format` | moses | CAT output: `moses`, `xliff` (for cat pipeline) |
| `--source-language` | en | Source language code for XLIFF |
| `--encoding` | utf-8 | Text file encoding for Moses/text output |

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
| `-p, --pipeline` | auto | Auto-detects from layout.json (`direct`, `office`, `xliff`, `cat`, `pikepdf`, `html`) |
| `--encoding` | utf-8 | Text file encoding for Moses/text input |
| `--min-font-size` | 6.0 | Minimum font size threshold |
| `--font-step` | 0.5 | Font size reduction step |
| `--render-method` | 1 | Text render method (line-by-line) |

**Intermediate files (Office/CAT pipeline):**
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

### Overview

| # | Pipeline | CLI Flag | Speed | Best For |
|---|----------|----------|-------|----------|
| 1 | Direct PDF | `--pipeline direct` | **~0.3s** | Unicode text (Hindi, Bengali, Arabic, CJK) |
| 2 | Office Roundtrip | `--pipeline office` | ~30s | Complex layouts, Office-origin PDFs |
| 3 | XLIFF | `--pipeline xliff` | ~2s | Professional CAT tools |
| 4 | Office CAT | `--pipeline cat` (default) | ~30s | Moses/XLIFF + Office conversion |
| 5 | PikePDF Low-Level | `--pipeline pikepdf` | **~0.1s** | Latin-only text, maximum speed |
| 6 | HTML Intermediate | `--pipeline html` | ~0.3s | Visual preview, browser editing |

### 1. Direct PDF (Redact+Insert) - Fastest

```
PDF → Extract text+positions → [Translate] → Redact+Insert → PDF
```

```bash
python main.py extract doc.pdf --pipeline direct -l Hindi
python main.py merge doc.pdf
```

**How it works:**
- Uses PyMuPDF to extract text blocks with exact bounding boxes
- Redacts (removes) original text with white overlay
- Inserts translations at exact same positions using TextWriter
- Embeds Unicode fonts (Nirmala UI) for non-Latin scripts (Bengali, Hindi, etc.)
- Auto-scales font size when translations are longer than original

**Pros:** 
- Fastest (~0.3s for typical PDF)
- No external dependencies beyond PyMuPDF
- Preserves images, vectors, and original layout perfectly

**Cons:** 
- Limited text reflow for significantly longer translations
- Text may be truncated if box is too small

### 2. Office Roundtrip

```
PDF → Office (DOCX/PPTX/XLSX) → Extract XML text → [Translate] → Update XML → Office → PDF
```

```bash
python main.py extract doc.pdf --pipeline office -l French
python main.py merge doc.pdf
```

**How it works:**
- Auto-detects source format from PDF metadata (Word, PowerPoint, Excel)
- Converts PDF to Office format using pdf2docx/python-pptx/openpyxl
- Extracts text from Office XML structure
- Updates XML with translations preserving formatting
- Converts back to PDF via LibreOffice
- Extracts and preserves images, SmartArt graphics
- Converts PDF annotations to Word comments

**Auto-detects format:**
- DOCX for Word documents
- PPTX for PowerPoint presentations  
- XLSX for Excel spreadsheets

**Requires:** LibreOffice for Office → PDF conversion

**Pros:**
- Best formatting preservation for Office-origin PDFs
- Handles complex layouts, tables, images
- Intermediate Office files can be manually edited

**Cons:** 
- Slower (~30s)
- Requires LibreOffice installation

### 3. XLIFF Export

```
PDF → Extract text → Generate XLIFF 1.2 → [CAT Tool Translation] → Parse XLIFF → PDF
```

```bash
python main.py extract doc.pdf --pipeline xliff -l German
# Edit input.pdf.xlf in your CAT tool (SDL Trados, memoQ, OmegaT)
python main.py merge doc.pdf
```

**How it works:**
- Extracts text blocks from PDF
- Generates OASIS XLIFF 1.2 compliant file using translate-toolkit
- Compatible with professional CAT tools
- Parses translated XLIFF and rebuilds PDF

**Pros:**
- Industry-standard format
- Works with any CAT tool
- Translation memory integration

**Cons:**
- Requires translate-toolkit for full compliance

### 4. Office CAT Pipeline (Default)

```
PDF → Office → Moses/XLIFF text files → [Translate] → Update Office → PDF
```

```bash
# Moses format (simple parallel text files)
python main.py extract doc.pdf --pipeline cat --cat-format moses -l Spanish
# Edit *_target.txt (one segment per line)
python main.py merge doc.pdf

# XLIFF format
python main.py extract doc.pdf --pipeline cat --cat-format xliff -l French
# Edit *.xlf in your CAT tool
python main.py merge doc.pdf
```

**How it works:**
- Combines Office conversion with CAT-friendly output formats
- **Moses format**: Simple parallel text files (`source.txt`, `target.txt`)
  - One segment per line
  - `<br>` markers for line breaks within segments
  - Easy to edit in any text editor
- **XLIFF format**: XML for professional CAT tools
- Updates Office XML with translations
- Converts to PDF via LibreOffice

**Moses format output:**
- `*_source.txt` - Source text (one segment per line)
- `*_target.txt` - Target text (translate this file)
- `*_source.mapping.json` - Segment ID mapping

**XLIFF format output:**
- `*.xlf` - XLIFF 1.2 file (OASIS compliant when translate-toolkit is installed)

**Pros:** 
- Best balance of quality and translator-friendly format
- Works with Moses MT systems
- Industry-standard XLIFF option

**Cons:** 
- Slower (~30s)
- Requires LibreOffice

### 5. PikePDF Low-Level Pipeline - Maximum Speed

```
PDF → Parse content streams → [Translate] → Rewrite streams → PDF
```

```bash
python main.py extract doc.pdf --pipeline pikepdf -l Spanish
python main.py merge doc.pdf
```

**How it works:**
- Directly parses PDF content streams (raw PDF operators)
- Extracts text from Tj/TJ text operators
- Rewrites content streams with translations
- Preserves all PDF structure exactly

**Performance:**
- Extraction: ~0.07s
- Merge: ~0.1s
- Fastest possible PDF manipulation

**Output files:**
- `*_source.txt` - Source text (one operator per line)
- `*_target.txt` - Target text (translate this file)
- `*_layout.json` - Operator positions and metadata

**Pros:** 
- **Absolutely fastest** - directly manipulates PDF bytes
- Perfect structure preservation
- No intermediate format conversion

**Cons:** 
- **Unicode NOT supported** (Bengali, Hindi, Arabic, CJK)
- Only works with Latin/ASCII text
- May have issues with complex font encodings

**Best for:**
- European language translations (Spanish, French, German, etc.)
- Maximum throughput requirements
- PDFs with simple text structure

**NOT recommended for:**
- Indian languages (Hindi, Bengali, Tamil, etc.)
- Arabic, Hebrew, or CJK languages
- Any non-Latin scripts

For Unicode languages, use `--pipeline direct` instead.

### 6. HTML Intermediate Pipeline - Visual Preview

```
PDF → HTML (CSS positioned) → [Translate] → PDF
```

```bash
python main.py extract doc.pdf --pipeline html -l Bengali
python main.py merge doc.pdf
```

**How it works:**
- Extracts text with exact positions from PDF
- Generates HTML with absolute-positioned CSS divs
- Includes Google Fonts for Unicode support
- Creates visual preview HTML (openable in browser)
- Renders back to PDF via PyMuPDF

**Output files:**
- `*_preview.html` - Visual preview (open in browser to see layout)
- `*_source.txt` - Source text (one segment per line)
- `*_target.txt` - Target text (translate this file)
- `*_translated.html` - Preview with translations
- `*_translated.pdf` - Final output

**Pros:** 
- Visual preview before final PDF
- Edit/inspect in browser
- Full Unicode support with embedded fonts
- CSS-based layout preservation

**Cons:** 
- Slightly larger output files
- Layout may differ slightly from original

**Best for:**
- Documents where you want visual verification
- When manual position adjustments might be needed
- Unicode translations (Hindi, Bengali, Arabic, etc.)

### Pipeline Auto-Detection

During merge, the pipeline is **automatically detected** from the layout JSON file:

```bash
# These are equivalent - pipeline auto-detected from layout file:
python main.py merge doc.pdf
python main.py merge doc.pdf --pipeline auto

# Explicit pipeline (overrides auto-detection):
python main.py merge doc.pdf --pipeline direct
```

The layout JSON stores which pipeline was used during extraction:
```json
{
  "pipeline": "direct_pdf",
  ...
}
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
    ├── direct_pdf.py      # Direct PDF pipeline (fastest, Unicode)
    ├── office_roundtrip.py # Office conversion pipeline
    ├── office_cat.py      # Office + CAT format pipeline
    ├── office_xml.py      # Office XML handlers (DOCX/PPTX/XLSX)
    ├── docx_roundtrip.py  # Legacy DOCX pipeline
    ├── xliff_format.py    # XLIFF export pipeline
    ├── pikepdf_lowlevel.py # Low-level PDF stream pipeline (fastest, Latin-only)
    └── html_intermediate.py # HTML intermediate pipeline (visual preview)
utils/
├── font_utils.py          # Font mapping/metrics
tests/
├── test_roundtrip.py      # Comprehensive tests
main.py                    # CLI entry point
```

## Dependencies

| Package | Purpose | Required For |
|---------|---------|-------------|
| PyMuPDF | PDF processing, text extraction | ✅ Core (all pipelines) |
| pdf2docx | PDF → DOCX conversion | `--pipeline office`, `--pipeline cat` |
| python-pptx | PowerPoint handling | `--pipeline office`, `--pipeline cat` |
| openpyxl | Excel handling | `--pipeline office`, `--pipeline cat` |
| translate-toolkit | XLIFF/TMX compliance | `--pipeline xliff` (OASIS standard) |
| pikepdf | Low-level PDF manipulation | `--pipeline pikepdf` |
| LibreOffice | Office → PDF conversion | `--pipeline office`, `--pipeline cat` |
| weasyprint | HTML → PDF (optional) | `--pipeline html` (optional, falls back to PyMuPDF) |

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
