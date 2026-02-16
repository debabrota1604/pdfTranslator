#!/usr/bin/env python3
"""
PDF Layout-Preserving Translation Engine.

Two-stage workflow:
  1. extract input.pdf  -> generates layout + translation file
  2. merge input.pdf output.pdf  -> rebuilds PDF with translations
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pdf_layout.extractor import extract_pdf_layout
from pdf_layout.rebuilder_unicode import PDFRebuilder, RebuildConfig, rebuild_pdf
from pdf_layout.translation_io import (
    generate_translate_file,
    generate_translated_template,
    parse_translated_file,
    get_translation_prompt,
)


def derive_paths(input_pdf: Path) -> dict[str, Path]:
    """Derive all intermediate file paths from input PDF."""
    base = input_pdf.parent / input_pdf.name
    return {
        "layout": Path(f"{base}_layout.json"),
        "translate": Path(f"{base}_translate.txt"),
        "translated": Path(f"{base}_translated.txt"),
        "translations": Path(f"{base}_translations.json"),
    }


def extract_command(args: argparse.Namespace) -> int:
    """Extract layout and generate translation file."""
    input_pdf = Path(args.input_pdf)
    
    if not input_pdf.exists():
        print(f"Error: Input PDF not found: {input_pdf}", file=sys.stderr)
        return 1
    
    paths = derive_paths(input_pdf)
    
    try:
        # Extract layout
        print(f"Extracting: {input_pdf}")
        document = extract_pdf_layout(input_pdf, paths["layout"])
        total_blocks = sum(len(page.blocks) for page in document.pages)
        print(f"  Layout: {paths['layout']} ({total_blocks} blocks)")
        
        # Generate translation file
        generate_translate_file(paths["layout"], paths["translate"])
        print(f"  Translate: {paths['translate']}")
        
        # Generate empty template for translated output
        generate_translated_template(paths["layout"], paths["translated"])
        print(f"  Template: {paths['translated']}")
        
        # Output prompt
        print()
        print("=" * 60)
        print("LLM PROMPT (copy this):")
        print("=" * 60)
        target_lang = args.language if args.language else "target language"
        print(get_translation_prompt(total_blocks, target_lang))
        print("=" * 60)
        print()
        print(f"Next: Translate {paths['translate']} -> {paths['translated']}")
        print(f"Then: python main.py merge {input_pdf} output.pdf")
        
        return 0
        
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


def merge_command(args: argparse.Namespace) -> int:
    """Merge translations and rebuild PDF."""
    input_pdf = Path(args.input_pdf)
    output_pdf = Path(args.output_pdf)
    
    if not input_pdf.exists():
        print(f"Error: Input PDF not found: {input_pdf}", file=sys.stderr)
        return 1
    
    paths = derive_paths(input_pdf)
    
    # Check required files
    if not paths["layout"].exists():
        print(f"Error: Layout not found: {paths['layout']}", file=sys.stderr)
        print(f"  Run: python main.py extract {input_pdf}")
        return 1
    
    if not paths["translated"].exists():
        print(f"Error: Translated file not found: {paths['translated']}", file=sys.stderr)
        return 1
    
    try:
        # Parse translations
        print(f"Parsing: {paths['translated']}")
        translations = parse_translated_file(
            paths["translated"],
            paths["layout"],
            paths["translations"]
        )
        print(f"  Parsed {len(translations)} blocks")
        
        # Rebuild PDF
        print(f"Rebuilding: {output_pdf}")
        config = RebuildConfig(
            min_font_size=args.min_font_size,
            font_step=args.font_step,
        )
        
        rebuild_pdf(
            pdf_path=input_pdf,
            layout_path=paths["layout"],
            translations=translations,
            output_path=output_pdf,
            config=config,
        )
        
        print(f"Done: {output_pdf}")
        return 0
        
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


def main() -> int:
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="PDF Layout-Preserving Translation Engine",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Workflow:
  1. python main.py extract input.pdf [-l LANGUAGE]
     Creates: input.pdf_layout.json, input.pdf_translate.txt, input.pdf_translated.txt
  
  2. Translate input.pdf_translate.txt -> input.pdf_translated.txt
     (Use LLM with the provided prompt)
  
  3. python main.py merge input.pdf output.pdf
     Creates: output.pdf with translated text
""",
    )
    
    subparsers = parser.add_subparsers(dest="command", help="Commands")
    
    # Extract command
    extract_parser = subparsers.add_parser(
        "extract",
        help="Extract layout and generate translation file",
    )
    extract_parser.add_argument("input_pdf", help="Input PDF file")
    extract_parser.add_argument(
        "-l", "--language",
        help="Target language for prompt (e.g., Hindi, Spanish)",
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
    merge_parser.set_defaults(func=merge_command)
    
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        return 0
    
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
