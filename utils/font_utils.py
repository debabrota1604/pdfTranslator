"""
Font Utilities Module.

Provides font mapping, metrics calculation, and font fallback logic.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

import fitz  # PyMuPDF


@dataclass
class FontMetrics:
    """Font metrics for text measurement."""
    
    font_name: str
    ascender: float
    descender: float
    line_height: float
    avg_char_width: float
    
    @property
    def em_height(self) -> float:
        """Get the em height (ascender - descender)."""
        return self.ascender - self.descender


class FontMapper:
    """
    Maps font names to PyMuPDF built-in fonts and provides fallbacks.
    
    PyMuPDF Built-in fonts:
    - helv: Helvetica
    - tiro: Times Roman
    - cour: Courier
    - symb: Symbol
    - zadb: ZapfDingbats
    """
    
    # Mapping of common font family patterns to PyMuPDF built-ins
    FONT_MAPPINGS: dict[str, str] = {
        # Helvetica family
        "helvetica": "helv",
        "arial": "helv",
        "sans": "helv",
        "gothic": "helv",
        "verdana": "helv",
        "tahoma": "helv",
        "calibri": "helv",
        "segoe": "helv",
        
        # Times family
        "times": "tiro",
        "roman": "tiro",
        "serif": "tiro",
        "georgia": "tiro",
        "cambria": "tiro",
        "garamond": "tiro",
        "palatino": "tiro",
        
        # Courier family
        "courier": "cour",
        "mono": "cour",
        "consolas": "cour",
        "monaco": "cour",
        "menlo": "cour",
        "lucida console": "cour",
        
        # Symbol fonts
        "symbol": "symb",
        "wingding": "zadb",
        "dingbat": "zadb",
    }
    
    # Bold and italic variants (PyMuPDF naming)
    STYLE_SUFFIXES: dict[str, str] = {
        "bold": "bo",
        "italic": "it",
        "oblique": "it",
        "bold-italic": "bi",
        "bolditalic": "bi",
        "bold-oblique": "bi",
    }
    
    DEFAULT_FONT = "helv"
    
    def __init__(self, fallback_font: Optional[str] = None):
        """
        Initialize font mapper.
        
        Args:
            fallback_font: Default fallback font (default: Helvetica).
        """
        self.fallback_font = fallback_font or self.DEFAULT_FONT
    
    def map_font(self, font_name: str) -> str:
        """
        Map a font name to PyMuPDF built-in font.
        
        Args:
            font_name: Original font name from PDF.
            
        Returns:
            PyMuPDF built-in font identifier.
        """
        if not font_name:
            return self.fallback_font
        
        font_lower = font_name.lower()
        
        # Check for direct match with built-in names
        if font_lower in ("helv", "tiro", "cour", "symb", "zadb"):
            return font_lower
        
        # Check for font family patterns
        base_font = self._find_base_font(font_lower)
        
        # Check for style suffix
        style_suffix = self._find_style_suffix(font_lower)
        
        if style_suffix:
            return f"{base_font}{style_suffix}"
        
        return base_font
    
    def _find_base_font(self, font_lower: str) -> str:
        """
        Find base font from font name.
        
        Args:
            font_lower: Lowercase font name.
            
        Returns:
            Base PyMuPDF font name.
        """
        for pattern, builtin in self.FONT_MAPPINGS.items():
            if pattern in font_lower:
                return builtin
        
        return self.fallback_font
    
    def _find_style_suffix(self, font_lower: str) -> str:
        """
        Find style suffix from font name.
        
        Args:
            font_lower: Lowercase font name.
            
        Returns:
            Style suffix or empty string.
        """
        for pattern, suffix in self.STYLE_SUFFIXES.items():
            if pattern in font_lower:
                return suffix
        
        return ""
    
    def is_monospace(self, font_name: str) -> bool:
        """
        Check if font is monospace.
        
        Args:
            font_name: Font name to check.
            
        Returns:
            True if monospace font.
        """
        font_lower = font_name.lower()
        monospace_patterns = ["courier", "mono", "consolas", "monaco", "menlo"]
        return any(p in font_lower for p in monospace_patterns)
    
    def is_serif(self, font_name: str) -> bool:
        """
        Check if font is serif.
        
        Args:
            font_name: Font name to check.
            
        Returns:
            True if serif font.
        """
        font_lower = font_name.lower()
        serif_patterns = ["times", "roman", "serif", "georgia", "garamond", "palatino"]
        sans_patterns = ["sans", "helvetica", "arial"]
        
        # Check if explicitly sans-serif first
        if any(p in font_lower for p in sans_patterns):
            return False
        
        return any(p in font_lower for p in serif_patterns)


def get_font_metrics(
    font_name: str,
    font_size: float = 12.0,
) -> FontMetrics:
    """
    Get font metrics for a given font.
    
    Args:
        font_name: Font name.
        font_size: Font size in points.
        
    Returns:
        FontMetrics object.
    """
    mapper = FontMapper()
    pymupdf_font = mapper.map_font(font_name)
    
    # Use fitz.Font to get metrics
    try:
        font = fitz.Font(pymupdf_font)
        ascender = font.ascender * font_size
        descender = font.descender * font_size
        
        # Estimate average character width (roughly 0.5 * font_size for proportional)
        if mapper.is_monospace(font_name):
            avg_char_width = font_size * 0.6
        else:
            avg_char_width = font_size * 0.5
        
        line_height = (ascender - descender) * 1.2  # Standard 120% line height
        
        return FontMetrics(
            font_name=pymupdf_font,
            ascender=round(ascender, 3),
            descender=round(descender, 3),
            line_height=round(line_height, 3),
            avg_char_width=round(avg_char_width, 3),
        )
    except Exception:
        # Fallback to estimates
        ascender = font_size * 0.8
        descender = font_size * -0.2
        line_height = font_size * 1.2
        avg_char_width = font_size * 0.5
        
        return FontMetrics(
            font_name=mapper.fallback_font,
            ascender=round(ascender, 3),
            descender=round(descender, 3),
            line_height=round(line_height, 3),
            avg_char_width=round(avg_char_width, 3),
        )


def map_font_name(font_name: str) -> str:
    """
    Convenience function to map font name to PyMuPDF built-in.
    
    Args:
        font_name: Original font name.
        
    Returns:
        PyMuPDF built-in font name.
    """
    mapper = FontMapper()
    return mapper.map_font(font_name)


def estimate_text_width(
    text: str,
    font_name: str,
    font_size: float,
) -> float:
    """
    Estimate the width of text in points.
    
    Args:
        text: Text to measure.
        font_name: Font name.
        font_size: Font size in points.
        
    Returns:
        Estimated width in points.
    """
    metrics = get_font_metrics(font_name, font_size)
    
    # Simple estimation based on average character width
    # This is approximate; actual rendering may differ
    return len(text) * metrics.avg_char_width


def estimate_text_height(
    text: str,
    font_name: str,
    font_size: float,
    bbox_width: float,
) -> float:
    """
    Estimate the height needed for text to fit in a given width.
    
    Args:
        text: Text to measure.
        font_name: Font name.
        font_size: Font size in points.
        bbox_width: Available width for text.
        
    Returns:
        Estimated height in points.
    """
    metrics = get_font_metrics(font_name, font_size)
    
    # Estimate number of lines
    text_width = estimate_text_width(text, font_name, font_size)
    
    if bbox_width <= 0:
        return metrics.line_height
    
    # Account for explicit line breaks
    lines_from_breaks = text.count('\n') + 1
    
    # Account for word wrapping
    chars_per_line = max(1, int(bbox_width / metrics.avg_char_width))
    wrapped_lines = max(1, len(text) // chars_per_line)
    
    total_lines = max(lines_from_breaks, wrapped_lines)
    
    return total_lines * metrics.line_height
