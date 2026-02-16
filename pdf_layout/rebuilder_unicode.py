"""
PDF Rebuilder Module with Unicode support using PyMuPDF TextWriter.

Rebuilds PDF with translated text while preserving layout.
Uses embedded TrueType fonts for proper multilingual text rendering.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional, Union

import fitz  # PyMuPDF

from .extractor import DocumentData, PageData, BlockData


# Default Unicode font paths for different platforms
UNICODE_FONTS = [
    Path("C:/Windows/Fonts/Nirmala.ttc"),  # Windows - Hindi/Devanagari
    Path("C:/Windows/Fonts/arial.ttf"),     # Windows fallback
    Path("C:/Windows/Fonts/seguiemj.ttf"),  # Windows Emoji
    Path("/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf"),  # Linux
    Path("/System/Library/Fonts/Supplemental/Arial Unicode.ttf"),  # macOS
]


def _find_unicode_font() -> Optional[Path]:
    """Find an available Unicode font on the system."""
    for font_path in UNICODE_FONTS:
        if font_path.exists():
            return font_path
    return None


def _contains_non_latin(text: str) -> bool:
    """Check if text contains non-Latin characters."""
    for ch in text:
        if ord(ch) > 127:
            return True
    return False


@dataclass
class RebuildConfig:
    """Configuration for PDF rebuilding."""
    
    min_font_size: float = 6.0
    font_step: float = 0.5
    overlay_color: tuple[float, float, float] = (1.0, 1.0, 1.0)  # White
    unicode_font_path: Optional[Path] = field(default_factory=_find_unicode_font)
    fallback_font: str = "helv"


class PDFRebuilder:
    """
    Rebuilds PDF with translated text using PyMuPDF TextWriter.
    
    Uses Font objects with embedded TrueType fonts for Unicode support.
    """
    
    def __init__(self, config: Optional[RebuildConfig] = None):
        self.config = config or RebuildConfig()
        self._unicode_font: Optional[fitz.Font] = None
        self._load_unicode_font()
    
    def _load_unicode_font(self) -> None:
        """Load the Unicode font for text rendering."""
        if self.config.unicode_font_path and self.config.unicode_font_path.exists():
            try:
                self._unicode_font = fitz.Font(
                    fontfile=str(self.config.unicode_font_path)
                )
                print(f"Loaded Unicode font: {self._unicode_font.name}")
            except Exception as e:
                print(f"Warning: Could not load Unicode font: {e}")
                self._unicode_font = None
    
    def rebuild(
        self,
        pdf_path: Union[str, Path],
        layout_data: Union[DocumentData, dict[str, Any], str, Path],
        translations: dict[str, str],
        output_path: Union[str, Path],
    ) -> None:
        """Rebuild PDF with translations."""
        pdf_path = Path(pdf_path)
        output_path = Path(output_path)
        
        if not pdf_path.exists():
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")
        
        # Load layout data
        document = self._load_layout(layout_data)
        
        # Create page lookup
        page_lookup = {page.page_number: page for page in document.pages}
        
        # Open PDF
        doc = fitz.open(str(pdf_path))
        
        try:
            for page_num in range(len(doc)):
                page = doc[page_num]
                page_data = page_lookup.get(page_num + 1)
                
                if page_data:
                    self._process_page(page, page_data, translations)
            
            # Save output
            output_path.parent.mkdir(parents=True, exist_ok=True)
            doc.save(
                str(output_path),
                garbage=4,
                deflate=True,
                clean=True,
            )
        finally:
            doc.close()
    
    def _process_page(
        self,
        page: fitz.Page,
        page_data: PageData,
        translations: dict[str, str],
    ) -> None:
        """Process a single page."""
        # First pass: add redaction annotations to cover original text
        for block in page_data.blocks:
            rect = fitz.Rect(block.bbox)
            page.add_redact_annot(rect, fill=(1, 1, 1))  # White fill
        
        # Apply all redactions - this actually removes the text
        page.apply_redactions()
        
        # Second pass: insert translated text
        for block in page_data.blocks:
            translated_text = translations.get(block.block_id, block.text)
            self._insert_text(page, block, translated_text)
    
    def _cover_rect(self, page: fitz.Page, rect: fitz.Rect) -> None:
        """Cover a rectangle with white fill."""
        shape = page.new_shape()
        shape.draw_rect(rect)
        shape.finish(
            color=self.config.overlay_color,
            fill=self.config.overlay_color,
        )
        shape.commit()
    
    def _insert_text(
        self,
        page: fitz.Page,
        block: BlockData,
        text: str,
    ) -> None:
        """Insert text into a block area using TextWriter for Unicode support."""
        rect = fitz.Rect(block.bbox)
        use_unicode = _contains_non_latin(text) and self._unicode_font is not None
        
        # Calculate fitting font size
        font_size = self._calculate_font_size(page, text, rect, block.font_size, use_unicode)
        
        # Get color
        color = self._hex_to_rgb(block.color)
        
        if use_unicode:
            # Use TextWriter with Unicode font
            self._insert_with_textwriter(page, rect, text, font_size, color)
        else:
            # Use standard insert_textbox for Latin text
            self._insert_with_textbox(page, rect, text, font_size, color, block.font_name)
    
    def _insert_with_textwriter(
        self,
        page: fitz.Page,
        rect: fitz.Rect,
        text: str,
        font_size: float,
        color: tuple[float, float, float],
    ) -> None:
        """Insert text using TextWriter with Unicode font."""
        tw = fitz.TextWriter(page.rect)
        
        # Split text into lines and render each
        lines = text.replace('\\n', '\n').split('\n')
        y_pos = rect.y0 + font_size  # Start position (baseline)
        line_height = font_size * 1.2
        
        for line in lines:
            if y_pos > rect.y1:
                break  # Stop if we've exceeded the rectangle
            
            # Truncate line if too wide (simple approach)
            try:
                # Append text to TextWriter
                tw.append(
                    pos=(rect.x0, y_pos),
                    text=line,
                    font=self._unicode_font,
                    fontsize=font_size,
                )
            except Exception as e:
                print(f"Warning: TextWriter failed for line: {e}")
            
            y_pos += line_height
        
        # Write to page
        try:
            tw.write_text(page, color=color)
        except Exception as e:
            print(f"Warning: Could not write text: {e}")
    
    def _insert_with_textbox(
        self,
        page: fitz.Page,
        rect: fitz.Rect,
        text: str,
        font_size: float,
        color: tuple[float, float, float],
        font_name: str,
    ) -> None:
        """Insert text using standard textbox for Latin text."""
        pymupdf_font = self._get_pymupdf_font(font_name)
        
        try:
            page.insert_textbox(
                rect,
                text.replace('\\n', '\n'),
                fontsize=font_size,
                fontname=pymupdf_font,
                color=color,
                align=fitz.TEXT_ALIGN_LEFT,
            )
        except Exception:
            try:
                page.insert_textbox(
                    rect,
                    text.replace('\\n', '\n'),
                    fontsize=font_size,
                    fontname=self.config.fallback_font,
                    color=color,
                    align=fitz.TEXT_ALIGN_LEFT,
                )
            except Exception as e:
                print(f"Warning: Failed to insert text: {e}")
    
    def _calculate_font_size(
        self,
        page: fitz.Page,
        text: str,
        rect: fitz.Rect,
        initial_size: float,
        use_unicode: bool,
    ) -> float:
        """Calculate font size that fits in the rectangle."""
        font_size = initial_size
        min_size = self.config.min_font_size
        
        # For Unicode text with TextWriter, estimate based on text length
        if use_unicode and self._unicode_font:
            while font_size >= min_size:
                # Estimate text dimensions
                lines = text.replace('\\n', '\n').split('\n')
                line_height = font_size * 1.2
                total_height = len(lines) * line_height
                
                if total_height <= rect.height:
                    return font_size
                
                font_size -= self.config.font_step
            return min_size
        
        # For Latin text, use insert_textbox measurement
        while font_size >= min_size:
            try:
                rc = page.insert_textbox(
                    rect,
                    text.replace('\\n', '\n'),
                    fontsize=font_size,
                    fontname=self.config.fallback_font,
                    render_mode=3,  # Invisible - just measure
                )
                if rc >= 0:
                    return font_size
            except Exception:
                pass
            font_size -= self.config.font_step
        
        return min_size
    
    def _get_pymupdf_font(self, font_name: str) -> str:
        """Map font name to PyMuPDF built-in font."""
        builtin_fonts = {
            "helv": "helv",
            "helvetica": "helv",
            "arial": "helv",
            "times": "tiro",
            "times-roman": "tiro",
            "courier": "cour",
        }
        font_lower = font_name.lower()
        for key, value in builtin_fonts.items():
            if key in font_lower:
                return value
        return self.config.fallback_font
    
    def _hex_to_rgb(self, hex_color: str) -> tuple[float, float, float]:
        """Convert hex color to RGB tuple (0.0-1.0)."""
        hex_color = hex_color.lstrip("#")
        if len(hex_color) != 6:
            return (0.0, 0.0, 0.0)
        try:
            r = int(hex_color[0:2], 16) / 255.0
            g = int(hex_color[2:4], 16) / 255.0
            b = int(hex_color[4:6], 16) / 255.0
            return (round(r, 3), round(g, 3), round(b, 3))
        except ValueError:
            return (0.0, 0.0, 0.0)
    
    def _load_layout(
        self,
        data: Union[DocumentData, dict[str, Any], str, Path]
    ) -> DocumentData:
        """Load layout data from various formats."""
        if isinstance(data, DocumentData):
            return data
        
        if isinstance(data, (str, Path)):
            path = Path(data)
            if path.exists():
                data = json.loads(path.read_text(encoding="utf-8"))
            elif isinstance(data, str) and data.strip().startswith("{"):
                data = json.loads(data)
            else:
                raise ValueError(f"Invalid layout data: {data}")
        
        return self._dict_to_document(data)
    
    def _dict_to_document(self, data: dict[str, Any]) -> DocumentData:
        """Convert dictionary to DocumentData."""
        pages = []
        for page_dict in data.get("pages", []):
            blocks = []
            for block_dict in page_dict.get("blocks", []):
                block = BlockData(
                    block_id=block_dict["block_id"],
                    bbox=tuple(block_dict["bbox"]),
                    text=block_dict["text"],
                    font_name=block_dict["font_name"],
                    font_size=block_dict["font_size"],
                    color=block_dict["color"],
                    writing_direction=block_dict.get("writing_direction", "ltr"),
                    line_height=block_dict.get("line_height"),
                )
                blocks.append(block)
            
            page = PageData(
                page_number=page_dict["page_number"],
                width=page_dict["width"],
                height=page_dict["height"],
                rotation=page_dict.get("rotation", 0),
                blocks=blocks,
            )
            pages.append(page)
        
        return DocumentData(
            source_file=data.get("source_file", ""),
            pages=pages,
        )


def rebuild_pdf(
    pdf_path: Union[str, Path],
    layout_path: Union[str, Path],
    translations: dict[str, str],
    output_path: Union[str, Path],
    config: Optional[RebuildConfig] = None,
) -> None:
    """Convenience function to rebuild PDF."""
    rebuilder = PDFRebuilder(config)
    rebuilder.rebuild(pdf_path, layout_path, translations, output_path)
