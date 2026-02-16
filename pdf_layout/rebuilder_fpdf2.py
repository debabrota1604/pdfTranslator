"""
PDF Rebuilder Module using fpdf2 for Unicode text support.

Rebuilds PDF with translated text while preserving layout.
Uses fpdf2 for proper Unicode/multilingual text rendering.
"""

from __future__ import annotations

import io
import json
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional, Union

import fitz  # PyMuPDF
from fpdf import FPDF

from .extractor import DocumentData, PageData, BlockData


# Default Unicode font paths for different platforms
UNICODE_FONTS = [
    Path("C:/Windows/Fonts/Nirmala.ttc"),  # Windows - Hindi/Devanagari
    Path("C:/Windows/Fonts/arial.ttf"),     # Windows fallback
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


class FPDF2TextRenderer:
    """
    Uses fpdf2 for reliable Unicode text rendering.
    Creates a text overlay PDF that is merged with the original.
    """
    
    def __init__(self, config: RebuildConfig):
        self.config = config
        self._font_registered = False
        self._font_name = "UniFont"
    
    def create_text_overlay(
        self,
        page_width: float,
        page_height: float,
        blocks: list[tuple[BlockData, str]],  # (block, translated_text)
    ) -> bytes:
        """
        Create a PDF page with only text at specified positions.
        
        Args:
            page_width: Page width in points.
            page_height: Page height in points.
            blocks: List of (block_data, translated_text) pairs.
            
        Returns:
            PDF bytes for the text overlay.
        """
        # Convert points to mm (fpdf2 uses mm by default)
        # 1 point = 0.352778 mm
        pt_to_mm = 0.352778
        width_mm = page_width * pt_to_mm
        height_mm = page_height * pt_to_mm
        
        pdf = FPDF(unit="mm", format=(width_mm, height_mm))
        pdf.set_auto_page_break(False)
        pdf.add_page()
        
        # Register Unicode font
        if self.config.unicode_font_path and self.config.unicode_font_path.exists():
            try:
                pdf.add_font(self._font_name, "", str(self.config.unicode_font_path))
            except Exception as e:
                print(f"Warning: Could not load font {self.config.unicode_font_path}: {e}")
                # Use built-in font as fallback
                self._font_name = "Helvetica"
        else:
            self._font_name = "Helvetica"
        
        for block, text in blocks:
            self._render_block(pdf, block, text, pt_to_mm)
        
        return pdf.output()
    
    def _render_block(
        self,
        pdf: FPDF,
        block: BlockData,
        text: str,
        pt_to_mm: float,
    ) -> None:
        """Render a single text block."""
        # Convert bbox from points to mm
        x0, y0, x1, y1 = block.bbox
        x_mm = x0 * pt_to_mm
        y_mm = y0 * pt_to_mm
        w_mm = (x1 - x0) * pt_to_mm
        h_mm = (y1 - y0) * pt_to_mm
        
        # Ensure minimum width
        if w_mm < 5:
            w_mm = 50  # Fallback width
        
        # Font size: points to fpdf2 points (fpdf2 uses actual points for font)
        font_size = block.font_size
        
        # Fit text to box by reducing font size if needed
        font_size = self._fit_font_size(pdf, text, w_mm, h_mm, font_size)
        
        # Set font and color
        pdf.set_font(self._font_name, size=font_size)
        
        # Parse hex color
        color = self._hex_to_rgb(block.color)
        pdf.set_text_color(color[0], color[1], color[2])
        
        # Position and render text
        pdf.set_xy(x_mm, y_mm)
        
        # Use multi_cell for text with word wrap, but catch errors
        try:
            # Line height based on font size
            line_h = font_size * 0.4
            if line_h < 1:
                line_h = 3
            pdf.multi_cell(w=w_mm, h=line_h, text=text, align="L")
        except Exception:
            # Fallback: just place text without wrapping
            try:
                pdf.text(x_mm, y_mm + font_size * pt_to_mm, text[:100])
            except Exception:
                pass  # Skip if text rendering fails
    
    def _fit_font_size(
        self,
        pdf: FPDF,
        text: str,
        width_mm: float,
        height_mm: float,
        initial_size: float,
    ) -> float:
        """Calculate font size that fits text in the given area."""
        font_size = initial_size
        min_size = self.config.min_font_size
        
        while font_size >= min_size:
            pdf.set_font(self._font_name, size=font_size)
            
            # Estimate text height
            lines = text.split('\n')
            total_height = 0
            line_h = font_size * 0.4  # Approximate line height
            
            for line in lines:
                try:
                    line_width = pdf.get_string_width(line)
                    num_wraps = max(1, int(line_width / max(width_mm, 1)) + 1)
                    total_height += num_wraps * line_h
                except Exception:
                    total_height += line_h
            
            if total_height <= height_mm:
                return font_size
            
            font_size -= self.config.font_step
        
        return max(min_size, font_size)
    
    def _hex_to_rgb(self, hex_color: str) -> tuple[int, int, int]:
        """Convert hex color to RGB tuple (0-255)."""
        hex_color = hex_color.lstrip("#")
        if len(hex_color) != 6:
            return (0, 0, 0)
        try:
            r = int(hex_color[0:2], 16)
            g = int(hex_color[2:4], 16)
            b = int(hex_color[4:6], 16)
            return (r, g, b)
        except ValueError:
            return (0, 0, 0)


class PDFRebuilder:
    """
    Rebuilds PDF with translated text using fpdf2 for Unicode support.
    
    Process:
    1. Open original PDF with PyMuPDF
    2. Cover original text blocks with white rectangles
    3. Create text overlay PDF using fpdf2 (proper Unicode rendering)
    4. Merge overlay onto original page
    5. Save output
    """
    
    def __init__(self, config: Optional[RebuildConfig] = None):
        self.config = config or RebuildConfig()
        self.text_renderer = FPDF2TextRenderer(self.config)
    
    def rebuild(
        self,
        pdf_path: Union[str, Path],
        layout_data: Union[DocumentData, dict[str, Any], str, Path],
        translations: dict[str, str],
        output_path: Union[str, Path],
    ) -> None:
        """
        Rebuild PDF with translations.
        
        Args:
            pdf_path: Path to original PDF.
            layout_data: Extracted layout data.
            translations: Dictionary mapping block_id to translated text.
            output_path: Path for output PDF.
        """
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
        # Collect blocks with translations
        blocks_to_render = []
        
        for block in page_data.blocks:
            translated_text = translations.get(block.block_id, block.text)
            
            # Cover original text with white rectangle
            self._cover_rect(page, fitz.Rect(block.bbox))
            
            blocks_to_render.append((block, translated_text))
        
        if not blocks_to_render:
            return
        
        # Create text overlay using fpdf2
        overlay_bytes = self.text_renderer.create_text_overlay(
            page_width=page_data.width,
            page_height=page_data.height,
            blocks=blocks_to_render,
        )
        
        # Merge overlay onto page
        overlay_doc = fitz.open(stream=overlay_bytes, filetype="pdf")
        try:
            overlay_page = overlay_doc[0]
            # Show the overlay PDF on top of the current page
            page.show_pdf_page(
                page.rect,
                overlay_doc,
                0,
                overlay=(True, True, True),
            )
        finally:
            overlay_doc.close()
    
    def _cover_rect(self, page: fitz.Page, rect: fitz.Rect) -> None:
        """Cover a rectangle with white fill."""
        shape = page.new_shape()
        shape.draw_rect(rect)
        shape.finish(
            color=self.config.overlay_color,
            fill=self.config.overlay_color,
        )
        shape.commit()
    
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
    """
    Convenience function to rebuild PDF.
    
    Args:
        pdf_path: Path to original PDF.
        layout_path: Path to layout JSON.
        translations: Translation dictionary.
        output_path: Output PDF path.
        config: Optional rebuild configuration.
    """
    rebuilder = PDFRebuilder(config)
    rebuilder.rebuild(pdf_path, layout_path, translations, output_path)
