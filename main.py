#!/usr/bin/env python3
"""
PDF Layout-Preserving Translation Engine.

Supports multiple translation pipelines:
- direct (default): Direct PDF text replacement using PyMuPDF
- office/docx: PDF → Office (auto-detect DOCX/PPTX/XLSX) → Translate → PDF
- xliff: Generate XLIFF format for professional CAT tools

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
    OfficeFormat,
)
from pdf_layout.source_detector import (
    detect_source_format,
    get_recommended_pipeline,
    SourceFormat,
)
from pdf_layout.rebuilder_unicode import RenderMethod


# Map CLI names to pipeline types
PIPELINE_MAP = {
    "direct": PipelineType.DIRECT_PDF,
    "1": PipelineType.DIRECT_PDF,
    "office": PipelineType.DOCX_ROUNDTRIP,
    "docx": PipelineType.DOCX_ROUNDTRIP,
    "2": PipelineType.DOCX_ROUNDTRIP,
    "xliff": PipelineType.XLIFF,
    "3": PipelineType.XLIFF,
}

RENDER_METHOD_MAP = {
    "1": RenderMethod.LINE_BY_LINE,
    "line-by-line": RenderMethod.LINE_BY_LINE,
    # Future methods:
    # "2": RenderMethod.WORD_WRAP,
    # "word-wrap": RenderMethod.WORD_WRAP,
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
        print(f"  → Use: python main.py extract {input_path} --pipeline office")
    elif info.format == SourceFormat.POWERPOINT:
        print(f"  → Use: python main.py extract {input_path} --pipeline office --office-format pptx")
    elif info.format == SourceFormat.EXCEL:
        print(f"  → Use: python main.py extract {input_path} --pipeline office --office-format xlsx")
    else:
        print(f"  → Use: python main.py extract {input_path} --pipeline direct")
    
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
                keep_intermediate=getattr(args, 'keep_intermediate', False),
            )
        elif pipeline_type == PipelineType.XLIFF:
            pipeline = create_xliff_pipeline(
                target_language=target_lang,
                source_language=getattr(args, 'source_language', 'en'),
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
        print(f"Then: python main.py merge {input_path} output.pdf --pipeline {args.pipeline}")
        
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
    output_path = Path(args.output_pdf)
    
    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}", file=sys.stderr)
        return 1
    
    # Parse pipeline type
    pipeline_type = PIPELINE_MAP.get(args.pipeline.lower(), PipelineType.DIRECT_PDF)
    target_lang = getattr(args, 'language', 'Hindi')
    
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
                keep_intermediate=getattr(args, 'keep_intermediate', False),
            )
        elif pipeline_type == PipelineType.XLIFF:
            pipeline = create_xliff_pipeline(target_language=target_lang)
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


def main() -> int:
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="PDF Layout-Preserving Translation Engine",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Pipelines:
  direct (1)   Direct PDF text replacement (default, fastest)
  office (2)   PDF → Office (auto-detect DOCX/PPTX/XLSX) → Translate → PDF
  xliff (3)    Generate XLIFF format for CAT tools

Workflow:
  1. python main.py info input.pdf
     Shows detected source format and recommended pipeline
  
  2. python main.py extract input.pdf --pipeline direct -l Hindi
     Creates: input.pdf_layout.json, input.pdf_translate.txt, input.pdf_translated.txt
  
  3. Translate input.pdf_translate.txt -> input.pdf_translated.txt
     (Use LLM or CAT tool with the provided prompt)
  
  4. python main.py merge input.pdf output.pdf --pipeline direct
     Creates: output.pdf with translated text

Examples:
  # Check source format
  python main.py info document.pdf
  
  # Direct PDF (default - fastest)
  python main.py extract document.pdf -l Spanish
  python main.py merge document.pdf translated.pdf
  
  # Office roundtrip (auto-detects Word/PowerPoint/Excel)
  python main.py extract document.pdf --pipeline office -l French
  python main.py merge document.pdf translated.pdf --pipeline office
  
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
    pipeline_help = "Translation pipeline: direct/1 (default), office/docx/2, xliff/3"
    office_format_help = "Office format: auto (default), docx, pptx, xlsx"
    
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
        default="direct",
        choices=["direct", "1", "office", "docx", "2", "xliff", "3"],
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
        "--source-language",
        type=str,
        default="en",
        help="Source language code for XLIFF (default: en)",
    )
    extract_parser.add_argument(
        "--keep-intermediate",
        action="store_true",
        help="Keep intermediate Office files",
    )
    extract_parser.set_defaults(func=extract_command)
    
    # Merge command
    merge_parser = subparsers.add_parser(
        "merge",
        help="Merge translations and rebuild PDF",
    )
    merge_parser.add_argument("input_pdf", help="Original input PDF file")
    merge_parser.add_argument("output_pdf", help="Output PDF file path")
    merge_parser.add_argument(
        "--pipeline", "-p",
        type=str,
        default="direct",
        choices=["direct", "1", "office", "docx", "2", "xliff", "3"],
        help=pipeline_help,
    )
    merge_parser.add_argument(
        "--office-format",
        type=str,
        default="auto",
        choices=["auto", "docx", "pptx", "xlsx"],
        help=office_format_help,
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
        "--keep-intermediate",
        action="store_true",
        help="Keep intermediate Office files",
    )
    merge_parser.set_defaults(func=merge_command)
    
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        return 0
    
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
