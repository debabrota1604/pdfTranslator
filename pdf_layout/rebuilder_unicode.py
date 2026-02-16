"""
PDF Rebuilder Module with Unicode support using PyMuPDF TextWriter.

Rebuilds PDF with translated text while preserving layout.
Uses embedded TrueType fonts for proper multilingual text rendering.

Render Methods:
- Method 1 (LINE_BY_LINE): Replace text line by line, preserving original line breaks.
  Trade-off: Good for similar-length translations, may overflow for longer translations.
- Future methods will be added here.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Optional, Union

import fitz  # PyMuPDF

from .extractor import DocumentData, PageData, BlockData


class RenderMethod(Enum):
    """Text rendering method for PDF rebuilding."""
    LINE_BY_LINE = 1  # Preserve line breaks, scale font to fit
    WORD_WRAP = 2     # Reflow text within box, ignore original line breaks
    # Future methods:
    # HYBRID = 3          # Try line-by-line first, fall back to word wrap
    # PARAGRAPH_REFLOW = 4  # Paragraph-level intelligent reflow


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
    
    # Render method selection
    render_method: RenderMethod = RenderMethod.LINE_BY_LINE
    
    # Font settings
    min_font_size: float = 6.0
    font_step: float = 0.5
    unicode_font_path: Optional[Path] = field(default_factory=_find_unicode_font)
    fallback_font: str = "helv"
    
    # Visual settings
    overlay_color: tuple[float, float, float] = (1.0, 1.0, 1.0)  # White


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
        # Add small padding to ensure complete coverage
        for block in page_data.blocks:
            rect = fitz.Rect(block.bbox)
            # Expand rect slightly to ensure complete text removal
            padded_rect = rect + (-1, -1, 1, 1)
            page.add_redact_annot(padded_rect, fill=(1, 1, 1))  # White fill
        
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
        """Insert text into a block area.
        
        Dispatches to appropriate render method based on config.
        """
        if self.config.render_method == RenderMethod.LINE_BY_LINE:
            self._insert_text_line_by_line(page, block, text)
        elif self.config.render_method == RenderMethod.WORD_WRAP:
            self._insert_text_word_wrap(page, block, text)
        else:
            # Default to line-by-line
            self._insert_text_line_by_line(page, block, text)
    
    def _insert_text_line_by_line(
        self,
        page: fitz.Page,
        block: BlockData,
        text: str,
    ) -> None:
        """Method 1: Line-by-line replacement.
        
        Preserves original line breaks from source text.
        Scales font to fit within bounding box.
        Trade-off: Good for similar-length translations, may overflow for longer text.
        """
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
        line_height = font_size * 1.3  # Match the calculation method
        y_pos = rect.y0 + font_size  # Start position (baseline)
        
        for line in lines:
            if y_pos > rect.y1:
                break  # Stop if we've exceeded the rectangle
            
            # Truncate line if too wide using font metrics
            truncated_line = self._truncate_line_to_fit(line, rect.width, font_size)
            
            try:
                # Append text to TextWriter
                tw.append(
                    pos=(rect.x0, y_pos),
                    text=truncated_line,
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
    
    def _truncate_line_to_fit(
        self,
        line: str,
        max_width: float,
        font_size: float,
    ) -> str:
        """Truncate line to fit within max_width using font metrics."""
        if not self._unicode_font or not line:
            return line
        
        # Use font's text_length for accurate measurement
        try:
            text_width = self._unicode_font.text_length(line, fontsize=font_size)
            if text_width <= max_width:
                return line
            
            # Binary search for fitting length
            left, right = 0, len(line)
            while left < right:
                mid = (left + right + 1) // 2
                test_text = line[:mid]
                width = self._unicode_font.text_length(test_text, fontsize=font_size)
                if width <= max_width:
                    left = mid
                else:
                    right = mid - 1
            
            # Truncate with ellipsis if significantly shortened
            if left < len(line) - 3:
                return line[:max(0, left - 2)] + ".."
            return line[:left]
        except Exception:
            # Fallback: simple character-based truncation
            avg_char_width = font_size * 0.55
            max_chars = int(max_width / avg_char_width)
            if len(line) > max_chars:
                return line[:max(0, max_chars - 2)] + ".."
            return line
    
    def _insert_text_word_wrap(
        self,
        page: fitz.Page,
        block: BlockData,
        text: str,
    ) -> None:
        """Method 2: Word-wrap replacement.
        
        Reflows text to fit within bounding box, ignoring original line breaks.
        Better for translations that are significantly longer than source.
        Trade-off: Loses original paragraph structure.
        """
        rect = fitz.Rect(block.bbox)
        use_unicode = _contains_non_latin(text) and self._unicode_font is not None
        
        # Flatten text - replace line breaks with spaces
        flat_text = text.replace('\\n', ' ').replace('\n', ' ')
        # Clean up multiple spaces
        while '  ' in flat_text:
            flat_text = flat_text.replace('  ', ' ')
        flat_text = flat_text.strip()
        
        # Calculate fitting font size with word wrap
        font_size = self._calculate_font_size_wordwrap(flat_text, rect, block.font_size, use_unicode)
        
        # Get color
        color = self._hex_to_rgb(block.color)
        
        if use_unicode:
            # Use TextWriter with word wrapping
            self._insert_wordwrap_unicode(page, rect, flat_text, font_size, color)
        else:
            # Use standard insert_textbox which does word wrapping automatically
            self._insert_with_textbox(page, rect, flat_text, font_size, color, block.font_name)
    
    def _calculate_font_size_wordwrap(
        self,
        text: str,
        rect: fitz.Rect,
        initial_size: float,
        use_unicode: bool,
    ) -> float:
        """Calculate font size for word-wrapped text."""
        font_size = initial_size
        min_size = self.config.min_font_size
        
        words = text.split()
        
        while font_size >= min_size:
            # Estimate how many lines we need with word wrap
            if use_unicode and self._unicode_font:
                lines_needed = self._estimate_lines_unicode(words, rect.width, font_size)
            else:
                # Rough estimate for Latin
                avg_char_width = font_size * 0.5
                chars_per_line = int(rect.width / avg_char_width)
                total_chars = len(text)
                lines_needed = (total_chars // chars_per_line) + 1
            
            line_height = font_size * 1.3
            total_height = lines_needed * line_height
            
            if total_height <= rect.height:
                return font_size
            
            font_size -= self.config.font_step
        
        return min_size
    
    def _estimate_lines_unicode(
        self,
        words: list[str],
        max_width: float,
        font_size: float,
    ) -> int:
        """Estimate number of lines needed for word-wrapped Unicode text."""
        if not self._unicode_font:
            return len(words)  # Rough fallback
        
        lines = 1
        current_width = 0.0
        space_width = self._unicode_font.text_length(' ', fontsize=font_size)
        
        for word in words:
            word_width = self._unicode_font.text_length(word, fontsize=font_size)
            
            if current_width + word_width > max_width:
                lines += 1
                current_width = word_width + space_width
            else:
                current_width += word_width + space_width
        
        return lines
    
    def _insert_wordwrap_unicode(
        self,
        page: fitz.Page,
        rect: fitz.Rect,
        text: str,
        font_size: float,
        color: tuple[float, float, float],
    ) -> None:
        """Insert word-wrapped Unicode text."""
        tw = fitz.TextWriter(page.rect)
        
        words = text.split()
        line_height = font_size * 1.3
        y_pos = rect.y0 + font_size
        current_line = []
        current_width = 0.0
        space_width = self._unicode_font.text_length(' ', fontsize=font_size) if self._unicode_font else font_size * 0.3
        
        for word in words:
            if self._unicode_font:
                word_width = self._unicode_font.text_length(word, fontsize=font_size)
            else:
                word_width = len(word) * font_size * 0.5
            
            # Check if word fits on current line
            if current_width + word_width > rect.width and current_line:
                # Render current line
                line_text = ' '.join(current_line)
                if y_pos <= rect.y1:
                    try:
                        tw.append(
                            pos=(rect.x0, y_pos),
                            text=line_text,
                            font=self._unicode_font,
                            fontsize=font_size,
                        )
                    except Exception:
                        pass
                
                y_pos += line_height
                current_line = [word]
                current_width = word_width + space_width
            else:
                current_line.append(word)
                current_width += word_width + space_width
        
        # Render last line
        if current_line and y_pos <= rect.y1:
            line_text = ' '.join(current_line)
            try:
                tw.append(
                    pos=(rect.x0, y_pos),
                    text=line_text,
                    font=self._unicode_font,
                    fontsize=font_size,
                )
            except Exception:
                pass
        
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
        """Calculate font size that fits in the rectangle (both width and height)."""
        font_size = initial_size
        min_size = self.config.min_font_size
        
        lines = text.replace('\\n', '\n').split('\n')
        max_line_len = max((len(line) for line in lines), default=1)
        
        # For Unicode text with TextWriter, check both width and height
        if use_unicode and self._unicode_font:
            while font_size >= min_size:
                line_height = font_size * 1.3  # Slightly more generous line height
                total_height = len(lines) * line_height
                
                # Estimate width using font metrics
                # Average char width for Unicode is roughly 0.6 * font_size
                avg_char_width = font_size * 0.55
                max_line_width = max_line_len * avg_char_width
                
                # Check both dimensions fit
                if total_height <= rect.height and max_line_width <= rect.width:
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
