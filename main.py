#!/usr/bin/env python3
"""
PDF Layout-Preserving Translation Engine.

Supports multiple translation pipelines:
- direct (default): Direct PDF text replacement using PyMuPDF
- office/docx: PDF -> Office (auto-detect DOCX/PPTX/XLSX) -> Translate -> PDF
- xliff: Generate XLIFF format for professional CAT tools
- cat: PDF -> Office -> Moses/XLIFF format -> Office -> PDF

Two-stage workflow:
  1. extract input.pdf  -> generates layout + translation file
  2. merge input.pdf output.pdf  -> rebuilds PDF with translations
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pdf_layout.pipelines import (
    PipelineType,
    create_pipeline,
    create_direct_pdf_pipeline,
    create_office_roundtrip_pipeline,
    create_xliff_pipeline,
    create_office_cat_pipeline,
    create_pikepdf_pipeline,
    create_html_pipeline,
    OfficeFormat,
    CATFormat,
    RenderMethod,
)
from pdf_layout.source_detector import (
    detect_source_format,
    get_recommended_pipeline,
    SourceFormat,
)


# Map CLI names to pipeline types
PIPELINE_MAP = {
    "auto": None,  # Auto-detect from layout file
    "direct": PipelineType.DIRECT_PDF,
    "1": PipelineType.DIRECT_PDF,
    "office": PipelineType.DOCX_ROUNDTRIP,
    "docx": PipelineType.DOCX_ROUNDTRIP,
    "2": PipelineType.DOCX_ROUNDTRIP,
    "xliff": PipelineType.XLIFF,
    "3": PipelineType.XLIFF,
    "cat": PipelineType.OFFICE_CAT,
    "moses": PipelineType.OFFICE_CAT,
    "4": PipelineType.OFFICE_CAT,
    "pikepdf": PipelineType.PIKEPDF_LOWLEVEL,
    "lowlevel": PipelineType.PIKEPDF_LOWLEVEL,
    "5": PipelineType.PIKEPDF_LOWLEVEL,
    "html": PipelineType.HTML_INTERMEDIATE,
    "6": PipelineType.HTML_INTERMEDIATE,
}

CAT_FORMAT_MAP = {
    "moses": CATFormat.MOSES,
    "xliff": CATFormat.XLIFF,
}

RENDER_METHOD_MAP = {
    "1": RenderMethod.REDACT_INSERT,
    "redact": RenderMethod.REDACT_INSERT,
    "redact-insert": RenderMethod.REDACT_INSERT,
    "2": RenderMethod.LINE_BY_LINE,
    "line-by-line": RenderMethod.LINE_BY_LINE,
    "3": RenderMethod.TEXTBOX_REFLOW,
    "reflow": RenderMethod.TEXTBOX_REFLOW,
}

OFFICE_FORMAT_MAP = {
    "auto": OfficeFormat.AUTO,
    "docx": OfficeFormat.DOCX,
    "pptx": OfficeFormat.PPTX,
    "xlsx": OfficeFormat.XLSX,
}


def info_command(args: argparse.Namespace) -> int:
    """Show PDF source format information."""
    input_path = Path(args.input_pdf)
    
    if not input_path.exists():
        print(f"Error: File not found: {input_path}", file=sys.stderr)
        return 1
    
    info = detect_source_format(input_path)
    
    print(f"Source Format Detection")
    print(f"=" * 50)
    print(f"File: {input_path}")
    print(f"Detected Format: {info.format.value}")
    print(f"Confidence: {info.confidence:.0%}")
    print(f"Creator: {info.creator or 'Unknown'}")
    print(f"Producer: {info.producer or 'Unknown'}")
    print(f"Details: {info.details}")
    print()
    print(f"Recommended Pipeline: {get_recommended_pipeline(info)}")
    
    if info.format == SourceFormat.WORD:
        print(f"  -> Use: python main.py extract {input_path} --pipeline office")
    elif info.format == SourceFormat.POWERPOINT:
        print(f"  -> Use: python main.py extract {input_path} --pipeline office --office-format pptx")
    elif info.format == SourceFormat.EXCEL:
        print(f"  -> Use: python main.py extract {input_path} --pipeline office --office-format xlsx")
    else:
        print(f"  -> Use: python main.py extract {input_path} --pipeline direct")
    
    return 0


def extract_command(args: argparse.Namespace) -> int:
    """Extract layout and generate translation file."""
    input_path = Path(args.input_pdf)
    
    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}", file=sys.stderr)
        return 1
    
    # Parse pipeline type
    pipeline_type = PIPELINE_MAP.get(args.pipeline.lower(), PipelineType.DIRECT_PDF)
    target_lang = args.language if args.language else "Hindi"
    
    # Parse render method for direct PDF pipeline
    render_method = RENDER_METHOD_MAP.get(
        getattr(args, 'render_method', '1'),
        RenderMethod.LINE_BY_LINE
    )
    
    try:
        # Create pipeline based on type
        if pipeline_type == PipelineType.DIRECT_PDF:
            pipeline = create_direct_pdf_pipeline(
                target_language=target_lang,
                render_method=render_method,
                min_font_size=getattr(args, 'min_font_size', 6.0),
                font_step=getattr(args, 'font_step', 0.5),
            )
        elif pipeline_type == PipelineType.DOCX_ROUNDTRIP:
            # Use Office roundtrip with auto-detection or specified format
            office_format = OFFICE_FORMAT_MAP.get(
                getattr(args, 'office_format', 'auto'),
                OfficeFormat.AUTO
            )
            pipeline = create_office_roundtrip_pipeline(
                target_language=target_lang,
                office_format=office_format,
            )
        elif pipeline_type == PipelineType.XLIFF:
            pipeline = create_xliff_pipeline(
                target_language=target_lang,
                source_language=getattr(args, 'source_language', 'en'),
            )
        elif pipeline_type == PipelineType.OFFICE_CAT:
            # Office CAT pipeline with Moses/XLIFF output
            office_format = OFFICE_FORMAT_MAP.get(
                getattr(args, 'office_format', 'auto'),
                OfficeFormat.AUTO
            )
            cat_format = CAT_FORMAT_MAP.get(
                getattr(args, 'cat_format', 'moses'),
                CATFormat.MOSES
            )
            pipeline = create_office_cat_pipeline(
                target_language=target_lang,
                source_language=getattr(args, 'source_language', 'en'),
                office_format=office_format,
                cat_format=cat_format,
                encoding=getattr(args, 'encoding', 'utf-8'),
            )
        elif pipeline_type == PipelineType.PIKEPDF_LOWLEVEL:
            pipeline = create_pikepdf_pipeline(
                target_language=target_lang,
            )
        elif pipeline_type == PipelineType.HTML_INTERMEDIATE:
            pipeline = create_html_pipeline(
                target_language=target_lang,
            )
        else:
            pipeline = create_direct_pdf_pipeline(target_language=target_lang)
        
        print(f"Pipeline: {pipeline.name}")
        print(f"Extracting: {input_path}")
        
        # Run extraction
        result = pipeline.extract(input_path)
        
        print(f"  Layout: {result.layout_path}")
        print(f"  Translate: {result.translate_path}")
        print(f"  Template: {result.translated_template_path}")
        
        if result.extra_files:
            for name, path in result.extra_files.items():
                print(f"  {name.capitalize()}: {path}")
        
        # Count blocks for prompt
        import json
        layout = json.loads(result.layout_path.read_text(encoding="utf-8"))
        if "pages" in layout:
            total_blocks = sum(len(p.get("blocks", [])) for p in layout["pages"])
        else:
            total_blocks = len(layout.get("blocks", []))
        
        # Output prompt
        print()
        print("=" * 60)
        print("LLM PROMPT (copy this):")
        print("=" * 60)
        print(pipeline.get_translation_prompt(total_blocks))
        print("=" * 60)
        print()
        print(f"Next: Translate {result.translate_path} -> {result.translated_template_path}")
        print(f"Then: python main.py merge {input_path}")
        print(f"      (auto-creates {input_path.stem}_translated.pdf)")
        
        return 0
        
    except ImportError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


def merge_command(args: argparse.Namespace) -> int:
    """Merge translations and rebuild document."""
    input_path = Path(args.input_pdf)
    
    # Check if input is a layout JSON file
    if input_path.suffix == '.json' and input_path.exists():
        return _merge_from_layout(input_path, args)
    
    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}", file=sys.stderr)
        return 1
    
    # Auto-derive output path if not specified (-o takes precedence over positional)
    output_pdf = getattr(args, 'output_pdf_opt', None) or args.output_pdf
    if output_pdf:
        output_path = Path(output_pdf)
    else:
        output_path = input_path.parent / f"{input_path.stem}_translated.pdf"
    
    # Try to auto-detect pipeline from existing layout file
    layout_path = input_path.parent / f"{input_path.name}_layout.json"
    pipeline_type = None
    target_lang = getattr(args, 'language', None)
    
    if layout_path.exists() and args.pipeline == 'auto':
        import json
        layout = json.loads(layout_path.read_text(encoding="utf-8"))
        stored_pipeline = layout.get("pipeline", "direct")
        if stored_pipeline == "office_roundtrip":
            pipeline_type = PipelineType.DOCX_ROUNDTRIP
        elif stored_pipeline == "xliff":
            pipeline_type = PipelineType.XLIFF
        elif stored_pipeline == "office_cat":
            pipeline_type = PipelineType.OFFICE_CAT
        elif stored_pipeline == "pikepdf_lowlevel":
            pipeline_type = PipelineType.PIKEPDF_LOWLEVEL
        elif stored_pipeline == "html_intermediate":
            pipeline_type = PipelineType.HTML_INTERMEDIATE
        else:
            pipeline_type = PipelineType.DIRECT_PDF
        target_lang = target_lang or layout.get("target_language", "Hindi")
        print(f"  Auto-detected pipeline: {stored_pipeline}")
    
    if pipeline_type is None:
        pipeline_type = PIPELINE_MAP.get(args.pipeline.lower(), PipelineType.DIRECT_PDF)
    target_lang = target_lang or 'Hindi'
    
    # Parse render method
    render_method = RENDER_METHOD_MAP.get(
        getattr(args, 'render_method', '1'),
        RenderMethod.LINE_BY_LINE
    )
    
    try:
        # Create pipeline
        if pipeline_type == PipelineType.DIRECT_PDF:
            pipeline = create_direct_pdf_pipeline(
                target_language=target_lang,
                render_method=render_method,
                min_font_size=args.min_font_size,
                font_step=args.font_step,
            )
        elif pipeline_type == PipelineType.DOCX_ROUNDTRIP:
            # Use Office roundtrip with auto-detection
            office_format = OFFICE_FORMAT_MAP.get(
                getattr(args, 'office_format', 'auto'),
                OfficeFormat.AUTO
            )
            pipeline = create_office_roundtrip_pipeline(
                target_language=target_lang,
                office_format=office_format,
            )
        elif pipeline_type == PipelineType.XLIFF:
            pipeline = create_xliff_pipeline(target_language=target_lang)
        elif pipeline_type == PipelineType.OFFICE_CAT:
            office_format = OFFICE_FORMAT_MAP.get(
                getattr(args, 'office_format', 'auto'),
                OfficeFormat.AUTO
            )
            cat_format = CAT_FORMAT_MAP.get(
                getattr(args, 'cat_format', 'moses'),
                CATFormat.MOSES
            )
            pipeline = create_office_cat_pipeline(
                target_language=target_lang,
                office_format=office_format,
                cat_format=cat_format,
                encoding=getattr(args, 'encoding', 'utf-8'),
            )
        elif pipeline_type == PipelineType.PIKEPDF_LOWLEVEL:
            pipeline = create_pikepdf_pipeline(
                target_language=target_lang,
            )
        elif pipeline_type == PipelineType.HTML_INTERMEDIATE:
            pipeline = create_html_pipeline(
                target_language=target_lang,
            )
        else:
            pipeline = create_direct_pdf_pipeline(target_language=target_lang)
        
        # Get paths
        paths = pipeline.derive_paths(input_path)
        
        # Check required files
        if not paths["layout"].exists():
            print(f"Error: Layout not found: {paths['layout']}", file=sys.stderr)
            print(f"  Run: python main.py extract {input_path} --pipeline {args.pipeline}")
            return 1
        
        if not paths["translated"].exists():
            print(f"Error: Translated file not found: {paths['translated']}", file=sys.stderr)
            return 1
        
        print(f"Pipeline: {pipeline.name}")
        print(f"Parsing: {paths['translated']}")
        
        # Run merge
        result = pipeline.merge(
            input_path=input_path,
            output_path=output_path,
            translated_path=paths["translated"],
            layout_path=paths["layout"],
        )
        
        print(f"  Processed {result.blocks_processed} blocks")
        
        if result.warnings:
            for warning in result.warnings:
                print(f"  Warning: {warning}")
        
        print(f"Done: {result.output_path}")
        return 0
        
    except ImportError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


def _merge_from_layout(layout_path: Path, args: argparse.Namespace) -> int:
    """Merge using layout JSON file directly."""
    import json
    
    layout = json.loads(layout_path.read_text(encoding="utf-8"))
    
    # Get source file from layout
    source_file = Path(layout.get("source_file", ""))
    if not source_file.exists():
        # Try relative to layout file
        source_file = layout_path.parent / source_file.name
    
    if not source_file.exists():
        print(f"Error: Source file not found: {layout.get('source_file')}", file=sys.stderr)
        return 1
    
    # Derive output path (-o takes precedence over positional)
    output_pdf = getattr(args, 'output_pdf_opt', None) or args.output_pdf
    if output_pdf:
        output_path = Path(output_pdf)
    else:
        output_path = source_file.parent / f"{source_file.stem}_translated.pdf"
    
    # Get pipeline from layout
    stored_pipeline = layout.get("pipeline", "direct")
    if stored_pipeline == "office_roundtrip":
        pipeline_type = PipelineType.DOCX_ROUNDTRIP
    elif stored_pipeline == "xliff":
        pipeline_type = PipelineType.XLIFF
    elif stored_pipeline == "office_cat":
        pipeline_type = PipelineType.OFFICE_CAT
    elif stored_pipeline == "pikepdf_lowlevel":
        pipeline_type = PipelineType.PIKEPDF_LOWLEVEL
    elif stored_pipeline == "html_intermediate":
        pipeline_type = PipelineType.HTML_INTERMEDIATE
    else:
        pipeline_type = PipelineType.DIRECT_PDF
    
    target_lang = layout.get("target_language", "Hindi")
    
    # Parse render method
    render_method = RENDER_METHOD_MAP.get(
        getattr(args, 'render_method', '1'),
        RenderMethod.LINE_BY_LINE
    )
    
    try:
        # Create pipeline
        if pipeline_type == PipelineType.DIRECT_PDF:
            pipeline = create_direct_pdf_pipeline(
                target_language=target_lang,
                render_method=render_method,
                min_font_size=args.min_font_size,
                font_step=args.font_step,
            )
        elif pipeline_type == PipelineType.DOCX_ROUNDTRIP:
            office_format = OFFICE_FORMAT_MAP.get(
                layout.get("office_format", "auto"),
                OfficeFormat.AUTO
            )
            pipeline = create_office_roundtrip_pipeline(
                target_language=target_lang,
                office_format=office_format,
            )
        elif pipeline_type == PipelineType.XLIFF:
            pipeline = create_xliff_pipeline(target_language=target_lang)
        elif pipeline_type == PipelineType.OFFICE_CAT:
            office_format = OFFICE_FORMAT_MAP.get(
                layout.get("office_format", "auto"),
                OfficeFormat.AUTO
            )
            cat_format = CAT_FORMAT_MAP.get(
                layout.get("cat_format", "moses"),
                CATFormat.MOSES
            )
            pipeline = create_office_cat_pipeline(
                target_language=target_lang,
                source_language=layout.get("source_language", "en"),
                office_format=office_format,
                cat_format=cat_format,
                encoding=layout.get("encoding", "utf-8"),
            )
        elif pipeline_type == PipelineType.PIKEPDF_LOWLEVEL:
            pipeline = create_pikepdf_pipeline(
                target_language=target_lang,
            )
        elif pipeline_type == PipelineType.HTML_INTERMEDIATE:
            pipeline = create_html_pipeline(
                target_language=target_lang,
            )
        else:
            pipeline = create_direct_pdf_pipeline(target_language=target_lang)
        
        # Get translated file path
        translated_path = layout_path.parent / f"{source_file.name}_translated.txt"
        if not translated_path.exists():
            print(f"Error: Translated file not found: {translated_path}", file=sys.stderr)
            return 1
        
        print(f"Pipeline: {pipeline.name}")
        print(f"Source: {source_file}")
        print(f"Layout: {layout_path}")
        print(f"Parsing: {translated_path}")
        
        # Run merge
        result = pipeline.merge(
            input_path=source_file,
            output_path=output_path,
            translated_path=translated_path,
            layout_path=layout_path,
        )
        
        print(f"  Processed {result.blocks_processed} blocks")
        
        if result.warnings:
            for warning in result.warnings:
                print(f"  Warning: {warning}")
        
        print(f"Done: {result.output_path}")
        return 0
        
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


def topdf_command(args: argparse.Namespace) -> int:
    """Convert edited Office file to PDF using LibreOffice.
    
    Workflow:
    1. User runs merge to get _translated.pptx
    2. User opens in PowerPoint/Word, manually adjusts text boxes
    3. User runs: python main.py topdf edited_file.pptx
    4. Gets fresh PDF from the manually edited Office file
    """
    import subprocess
    import shutil
    from pathlib import Path
    
    input_path = Path(args.input_file)
    
    if not input_path.exists():
        print(f"Error: File not found: {input_path}", file=sys.stderr)
        return 1
    
    # Validate extension
    ext = input_path.suffix.lower()
    if ext not in ('.docx', '.pptx', '.xlsx'):
        print(f"Error: Unsupported file type: {ext}", file=sys.stderr)
        print("Supported: .docx, .pptx, .xlsx", file=sys.stderr)
        return 1
    
    # Determine output path
    if args.output_pdf:
        output_path = Path(args.output_pdf)
    else:
        output_path = input_path.with_suffix('.pdf')
    
    # Find LibreOffice
    lo_paths = [
        Path("C:/Program Files/LibreOffice/program/soffice.exe"),
        Path("C:/Program Files (x86)/LibreOffice/program/soffice.exe"),
        Path("/usr/bin/libreoffice"),
        Path("/usr/bin/soffice"),
        Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
    ]
    
    libreoffice_path = None
    for path in lo_paths:
        if path.exists():
            libreoffice_path = path
            break
    
    if not libreoffice_path:
        print("Error: LibreOffice not found.", file=sys.stderr)
        print("Install from: https://www.libreoffice.org/", file=sys.stderr)
        return 1
    
    print(f"Converting: {input_path}")
    print(f"Output: {output_path}")
    
    # Convert
    cmd = [
        str(libreoffice_path),
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(output_path.parent),
        str(input_path),
    ]
    
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        
        if result.returncode != 0:
            print(f"Error: LibreOffice conversion failed: {result.stderr}", file=sys.stderr)
            return 1
        
        # LibreOffice names output based on input filename
        generated = input_path.with_suffix(".pdf")
        generated_in_outdir = output_path.parent / generated.name
        
        if generated_in_outdir != output_path and generated_in_outdir.exists():
            shutil.move(str(generated_in_outdir), str(output_path))
        
        output_size = output_path.stat().st_size / 1024
        print(f"Done: {output_path} ({output_size:.1f} KB)")
        return 0
        
    except subprocess.TimeoutExpired:
        print("Error: LibreOffice conversion timed out", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


def main() -> int:
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="PDF Layout-Preserving Translation Engine",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Pipelines:
  direct (1)   Direct PDF text replacement (Unicode-friendly)
  office (2)   PDF -> Office (auto-detect DOCX/PPTX/XLSX) -> Translate -> PDF
  xliff (3)    Generate XLIFF format for CAT tools
  cat (4)      PDF -> Office -> Moses/XLIFF -> Office -> PDF
  pikepdf (5)  Low-level content stream manipulation (fastest, Latin-only)
  html (6)     PDF -> HTML (CSS positioned) -> Translate -> PDF

CAT Formats (for cat pipeline):
  moses        Parallel text files (source.txt, target.txt) - line-by-line
  xliff        XLIFF 1.2 format for professional CAT tools

Workflow:
  1. python main.py extract input.pdf -l Hindi
     Creates: input.pdf_layout.json, input.pdf_translate.txt, input.pdf_translated.txt
  
  2. Translate input.pdf_translate.txt -> input.pdf_translated.txt
     (Use LLM or CAT tool with the provided prompt)
  
  3. python main.py merge input.pdf
     Creates: input_translated.pdf (auto-detects pipeline from layout)

Examples:
  # Minimal workflow (just 2 commands!)
  python main.py extract doc.pdf -l Spanish
  python main.py merge doc.pdf
  
  # Or use layout file directly
  python main.py merge doc.pdf_layout.json
  
  # Specify output name
  python main.py merge doc.pdf -o translated.pdf
  
  # Check source format first
  python main.py info document.pdf
  
  # Office roundtrip (auto-detects Word/PowerPoint/Excel)
  python main.py extract document.pdf --pipeline office -l French
  python main.py merge document.pdf
  
  # Force specific Office format
  python main.py extract presentation.pdf --pipeline office --office-format pptx
  
  # XLIFF for CAT tools
  python main.py extract document.pdf --pipeline xliff -l German
""",
    )
    
    subparsers = parser.add_subparsers(dest="command", help="Commands")
    
    # Info command
    info_parser = subparsers.add_parser(
        "info",
        help="Detect source format of PDF (Word/PowerPoint/Excel)",
    )
    info_parser.add_argument("input_pdf", help="Input PDF file")
    info_parser.set_defaults(func=info_command)
    
    # Common options
    pipeline_help = "Translation pipeline: direct/1, office/2, xliff/3, cat/4, pikepdf/5, html/6"
    office_format_help = "Office format: auto (default), docx, pptx, xlsx"
    cat_format_help = "CAT output format: moses (default), xliff"
    
    # Extract command
    extract_parser = subparsers.add_parser(
        "extract",
        help="Extract layout and generate translation file",
    )
    extract_parser.add_argument("input_pdf", help="Input PDF file")
    extract_parser.add_argument(
        "-l", "--language",
        default="Hindi",
        help="Target language (default: Hindi)",
    )
    extract_parser.add_argument(
        "--pipeline", "-p",
        type=str,
        default="cat",
        choices=["direct", "1", "office", "docx", "2", "xliff", "3", "cat", "moses", "4", "pikepdf", "lowlevel", "5", "html", "6"],
        help=pipeline_help,
    )
    extract_parser.add_argument(
        "--office-format",
        type=str,
        default="auto",
        choices=["auto", "docx", "pptx", "xlsx"],
        help=office_format_help,
    )
    extract_parser.add_argument(
        "--cat-format",
        type=str,
        default="moses",
        choices=["moses", "xliff"],
        help=cat_format_help,
    )
    extract_parser.add_argument(
        "--source-language",
        type=str,
        default="en",
        help="Source language code for XLIFF (default: en)",
    )
    extract_parser.add_argument(
        "--encoding",
        type=str,
        default="utf-8",
        help="Text file encoding for Moses/text output (default: utf-8)",
    )
    extract_parser.set_defaults(func=extract_command)
    
    # Merge command
    merge_parser = subparsers.add_parser(
        "merge",
        help="Merge translations and rebuild PDF",
    )
    merge_parser.add_argument(
        "input_pdf",
        help="Input PDF file or layout JSON file (auto-detects settings)",
    )
    merge_parser.add_argument(
        "output_pdf",
        nargs="?",
        default=None,
        help="Output PDF (default: {input}_translated.pdf)",
    )
    merge_parser.add_argument(
        "-o", "--output",
        dest="output_pdf_opt",
        help="Output PDF file (alternative to positional arg)",
    )
    merge_parser.add_argument(
        "--pipeline", "-p",
        type=str,
        default="auto",
        choices=["auto", "direct", "1", "office", "docx", "2", "xliff", "3", "cat", "moses", "4", "pikepdf", "lowlevel", "5", "html", "6"],
        help="Pipeline (default: auto-detect from layout)",
    )
    merge_parser.add_argument(
        "--office-format",
        type=str,
        default="auto",
        choices=["auto", "docx", "pptx", "xlsx"],
        help=office_format_help,
    )
    merge_parser.add_argument(
        "--cat-format",
        type=str,
        default="moses",
        choices=["moses", "xliff"],
        help=cat_format_help,
    )
    merge_parser.add_argument(
        "--min-font-size",
        type=float,
        default=6.0,
        help="Minimum font size (default: 6.0)",
    )
    merge_parser.add_argument(
        "--font-step",
        type=float,
        default=0.5,
        help="Font size step (default: 0.5)",
    )
    merge_parser.add_argument(
        "--render-method",
        type=str,
        default="1",
        choices=["1", "line-by-line"],
        help="Text render method for direct pipeline (default: 1)",
    )
    merge_parser.add_argument(
        "--encoding",
        type=str,
        default="utf-8",
        help="Text file encoding for Moses/text output (default: utf-8)",
    )
    merge_parser.set_defaults(func=merge_command)
    
    # ToPDF command - convert edited Office files to PDF
    topdf_parser = subparsers.add_parser(
        "topdf",
        help="Convert edited Office file (DOCX/PPTX/XLSX) to PDF",
    )
    topdf_parser.add_argument(
        "input_file",
        help="Office file to convert (DOCX, PPTX, or XLSX)",
    )
    topdf_parser.add_argument(
        "output_pdf",
        nargs="?",
        default=None,
        help="Output PDF (default: same name with .pdf extension)",
    )
    topdf_parser.set_defaults(func=topdf_command)
    
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        return 0
    
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
