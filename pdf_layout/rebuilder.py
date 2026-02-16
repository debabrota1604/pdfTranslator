"""
PDF Rebuilder Module.

Rebuilds PDF with translated text while preserving layout.
Handles font scaling and text insertion within bounding boxes.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional, Union

import fitz  # PyMuPDF

from .extractor import DocumentData, PageData, BlockData


# Unicode font path for non-Latin scripts (Hindi, etc.)
# Windows: Nirmala UI supports Devanagari
UNICODE_FONT_PATH = Path("C:/Windows/Fonts/Nirmala.ttc")


def _contains_non_latin(text: str) -> bool:
    """Check if text contains non-Latin characters (e.g., Devanagari)."""
    for ch in text:
        if ord(ch) > 127:  # Non-ASCII
            return True
    return False


@dataclass
class RebuildConfig:
    """Configuration for PDF rebuilding."""
    
    min_font_size: float = 6.0  # Minimum font size threshold
    font_step: float = 0.5  # Font size reduction step
    fallback_font: str = "helv"  # PyMuPDF built-in font name
    overlay_color: tuple[float, float, float] = (1.0, 1.0, 1.0)  # White
    text_align: int = fitz.TEXT_ALIGN_LEFT
    preserve_line_breaks: bool = True
    unicode_font_path: Path = UNICODE_FONT_PATH  # Font for non-Latin text


class FontScaler:
    """
    Handles deterministic font scaling to fit text within bounding boxes.
    """
    
    def __init__(self, config: RebuildConfig):
        """
        Initialize font scaler.
        
        Args:
            config: Rebuild configuration.
        """
        self.config = config
    
    def calculate_fitting_font_size(
        self,
        page: fitz.Page,
        text: str,
        bbox: tuple[float, float, float, float],
        initial_font_size: float,
        font_name: str,
        use_unicode_font: bool = False,
    ) -> float:
        """
        Calculate the font size that fits text within the bounding box.
        
        Uses binary-like search with fixed steps for determinism.
        
        Args:
            page: PyMuPDF page object.
            text: Text to fit.
            bbox: Bounding box (x0, y0, x1, y1).
            initial_font_size: Starting font size.
            font_name: Font to use.
            use_unicode_font: Whether to use Unicode font for non-Latin text.
            
        Returns:
            Font size that fits, or minimum if cannot fit.
        """
        rect = fitz.Rect(bbox)
        font_size = initial_font_size
        
        # For Unicode text, use the Unicode font file
        if use_unicode_font and self.config.unicode_font_path.exists():
            while font_size >= self.config.min_font_size:
                try:
                    rc = page.insert_textbox(
                        rect,
                        text,
                        fontsize=font_size,
                        fontfile=str(self.config.unicode_font_path),
                        align=self.config.text_align,
                        render_mode=3,  # Invisible - just measure
                    )
                    if rc >= 0:
                        return round(font_size, 3)
                except Exception:
                    pass
                font_size -= self.config.font_step
            return self.config.min_font_size
        
        # Get appropriate fontname for PyMuPDF
        pymupdf_font = self._get_pymupdf_font(font_name)
        
        while font_size >= self.config.min_font_size:
            # Test if text fits using insert_textbox with render_mode=3 (invisible)
            # This doesn't actually render but returns the overflow
            try:
                rc = page.insert_textbox(
                    rect,
                    text,
                    fontsize=font_size,
                    fontname=pymupdf_font,
                    align=self.config.text_align,
                    render_mode=3,  # Invisible - just measure
                )
                
                # rc < 0 means overflow, rc >= 0 means text fits
                if rc >= 0:
                    return round(font_size, 3)
                    
            except Exception:
                # If font fails, try fallback
                try:
                    rc = page.insert_textbox(
                        rect,
                        text,
                        fontsize=font_size,
                        fontname=self.config.fallback_font,
                        align=self.config.text_align,
                        render_mode=3,
                    )
                    if rc >= 0:
                        return round(font_size, 3)
                except Exception:
                    pass
            
            # Reduce font size
            font_size -= self.config.font_step
        
        # Return minimum font size if nothing fits
        return self.config.min_font_size
    
    def _get_pymupdf_font(self, font_name: str) -> str:
        """
        Map font name to PyMuPDF built-in font.
        
        Args:
            font_name: Original font name.
            
        Returns:
            PyMuPDF font identifier.
        """
        # PyMuPDF built-in fonts
        builtin_fonts = {
            "helv": "helv",
            "helvetica": "helv",
            "times": "tiro",
            "times-roman": "tiro",
            "courier": "cour",
            "symbol": "symb",
            "zapfdingbats": "zadb",
        }
        
        font_lower = font_name.lower()
        
        # Check for exact match or substring match
        for key, value in builtin_fonts.items():
            if key in font_lower:
                return value
        
        # Default to Helvetica
        return self.config.fallback_font


class PDFRebuilder:
    """
    Rebuilds PDF with translated text while preserving layout.
    
    Process:
    1. Open original PDF
    2. For each block with translation:
       - Cover original text with white rectangle
       - Insert new text with appropriate font size
    3. Save output PDF
    """
    
    def __init__(
        self,
        config: Optional[RebuildConfig] = None,
    ):
        """
        Initialize rebuilder.
        
        Args:
            config: Rebuild configuration (uses defaults if None).
        """
        self.config = config or RebuildConfig()
        self.font_scaler = FontScaler(self.config)
    
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
            
            # Save output with deterministic settings
            output_path.parent.mkdir(parents=True, exist_ok=True)
            doc.save(
                str(output_path),
                garbage=4,  # Maximum garbage collection
                deflate=True,  # Compress
                clean=True,  # Clean content streams
            )
        finally:
            doc.close()
    
    def _load_layout(
        self,
        data: Union[DocumentData, dict[str, Any], str, Path]
    ) -> DocumentData:
        """
        Load layout data from various formats.
        
        Args:
            data: Layout data.
            
        Returns:
            DocumentData object.
        """
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
        from .extractor import DocumentData, PageData, BlockData
        
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
    
    def _process_page(
        self,
        page: fitz.Page,
        page_data: PageData,
        translations: dict[str, str],
    ) -> None:
        """
        Process a single page, replacing text in blocks.
        
        Args:
            page: PyMuPDF page object.
            page_data: Extracted page data.
            translations: Translation dictionary.
        """
        for block in page_data.blocks:
            # Get translated text (use original if not translated)
            translated_text = translations.get(block.block_id, block.text)
            
            # Process this block
            self._replace_block_text(page, block, translated_text)
    
    def _replace_block_text(
        self,
        page: fitz.Page,
        block: BlockData,
        new_text: str,
    ) -> None:
        """
        Replace text in a single block.
        
        Args:
            page: PyMuPDF page object.
            block: Block data.
            new_text: New text to insert.
        """
        rect = fitz.Rect(block.bbox)
        
        # Step 1: Cover original text with white rectangle
        self._cover_rect(page, rect)
        
        # Check if text contains non-Latin characters (Hindi, etc.)
        use_unicode_font = _contains_non_latin(new_text)
        
        # Step 2: Calculate fitting font size
        font_size = self.font_scaler.calculate_fitting_font_size(
            page=page,
            text=new_text,
            bbox=block.bbox,
            initial_font_size=block.font_size,
            font_name=block.font_name,
            use_unicode_font=use_unicode_font,
        )
        
        # Step 3: Get color
        color = self._hex_to_rgb(block.color)
        
        # Step 4: Determine text alignment based on writing direction
        align = self._get_alignment(block.writing_direction)
        
        # Step 5: Insert new text
        if use_unicode_font and self.config.unicode_font_path.exists():
            # Use Unicode font file for non-Latin text
            try:
                page.insert_textbox(
                    rect,
                    new_text,
                    fontsize=font_size,
                    fontfile=str(self.config.unicode_font_path),
                    color=color,
                    align=align,
                )
                return
            except Exception as e:
                print(f"Warning: Unicode font failed for block {block.block_id}: {e}")
        
        # Use built-in font for Latin text
        font_name = self.font_scaler._get_pymupdf_font(block.font_name)
        try:
            page.insert_textbox(
                rect,
                new_text,
                fontsize=font_size,
                fontname=font_name,
                color=color,
                align=align,
            )
        except Exception:
            # Fallback: try with default font
            try:
                page.insert_textbox(
                    rect,
                    new_text,
                    fontsize=font_size,
                    fontname=self.config.fallback_font,
                    color=color,
                    align=align,
                )
            except Exception as e:
                # Log error but continue
                print(f"Warning: Failed to insert text for block {block.block_id}: {e}")
    
    def _cover_rect(self, page: fitz.Page, rect: fitz.Rect) -> None:
        """
        Cover a rectangle with white fill.
        
        Args:
            page: PyMuPDF page object.
            rect: Rectangle to cover.
        """
        shape = page.new_shape()
        shape.draw_rect(rect)
        shape.finish(
            color=self.config.overlay_color,
            fill=self.config.overlay_color,
        )
        shape.commit()
    
    def _hex_to_rgb(self, hex_color: str) -> tuple[float, float, float]:
        """
        Convert hex color to RGB tuple.
        
        Args:
            hex_color: Hex color string like "#000000".
            
        Returns:
            RGB tuple with values 0.0-1.0.
        """
        hex_color = hex_color.lstrip("#")
        
        if len(hex_color) != 6:
            return (0.0, 0.0, 0.0)  # Default to black
        
        try:
            r = int(hex_color[0:2], 16) / 255.0
            g = int(hex_color[2:4], 16) / 255.0
            b = int(hex_color[4:6], 16) / 255.0
            return (round(r, 3), round(g, 3), round(b, 3))
        except ValueError:
            return (0.0, 0.0, 0.0)
    
    def _get_alignment(self, writing_direction: str) -> int:
        """
        Get text alignment based on writing direction.
        
        Args:
            writing_direction: "ltr", "rtl", or "ttb".
            
        Returns:
            PyMuPDF alignment constant.
        """
        if writing_direction == "rtl":
            return fitz.TEXT_ALIGN_RIGHT
        elif writing_direction == "center":
            return fitz.TEXT_ALIGN_CENTER
        else:
            return fitz.TEXT_ALIGN_LEFT


def rebuild_pdf(
    pdf_path: Union[str, Path],
    layout_path: Union[str, Path],
    translations_path: Union[str, Path],
    output_path: Union[str, Path],
    config: Optional[RebuildConfig] = None,
) -> None:
    """
    Convenience function to rebuild PDF with translations.
    
    Args:
        pdf_path: Path to original PDF.
        layout_path: Path to layout JSON.
        translations_path: Path to translations JSON.
        output_path: Path for output PDF.
        config: Optional rebuild configuration.
    """
    # Load translations
    translations_path = Path(translations_path)
    translations = json.loads(translations_path.read_text(encoding="utf-8"))
    
    # Load layout
    layout_path = Path(layout_path)
    layout_data = json.loads(layout_path.read_text(encoding="utf-8"))
    
    # Rebuild
    rebuilder = PDFRebuilder(config=config)
    rebuilder.rebuild(
        pdf_path=pdf_path,
        layout_data=layout_data,
        translations=translations,
        output_path=output_path,
    )
