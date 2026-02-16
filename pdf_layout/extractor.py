"""
PDF Layout Extractor Module.

Extracts text blocks with bounding boxes, font information, and layout metadata
from PDF files using PyMuPDF in a deterministic manner.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any, Optional, Union

import fitz  # PyMuPDF


@dataclass
class BlockData:
    """Represents a single text block extracted from a PDF page."""
    
    block_id: str
    bbox: tuple[float, float, float, float]  # (x0, y0, x1, y1)
    text: str
    font_name: str
    font_size: float
    color: str  # Hex color string
    writing_direction: str = "ltr"  # ltr, rtl, ttb
    line_height: Optional[float] = None
    
    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        return {
            "block_id": self.block_id,
            "bbox": [round(v, 3) for v in self.bbox],
            "text": self.text,
            "font_name": self.font_name,
            "font_size": round(self.font_size, 3),
            "color": self.color,
            "writing_direction": self.writing_direction,
            "line_height": round(self.line_height, 3) if self.line_height else None,
        }


@dataclass
class PageData:
    """Represents a single page with its extracted blocks."""
    
    page_number: int
    width: float
    height: float
    rotation: int = 0
    blocks: list[BlockData] = field(default_factory=list)
    
    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        return {
            "page_number": self.page_number,
            "width": round(self.width, 3),
            "height": round(self.height, 3),
            "rotation": self.rotation,
            "blocks": [block.to_dict() for block in self.blocks],
        }


@dataclass
class DocumentData:
    """Represents the entire extracted document structure."""
    
    source_file: str
    pages: list[PageData] = field(default_factory=list)
    
    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        return {
            "source_file": self.source_file,
            "pages": [page.to_dict() for page in self.pages],
        }
    
    def to_json(self, indent: int = 2) -> str:
        """Convert to JSON string."""
        return json.dumps(self.to_dict(), indent=indent, ensure_ascii=False)


class PDFExtractor:
    """
    Extracts text blocks with layout information from PDF files.
    
    Uses PyMuPDF's get_text("dict") for structured extraction.
    All operations are deterministic and reproducible.
    """
    
    def __init__(self, pdf_path: Union[str, Path]):
        """
        Initialize extractor with PDF file path.
        
        Args:
            pdf_path: Path to the PDF file to extract from.
        """
        self.pdf_path = Path(pdf_path)
        if not self.pdf_path.exists():
            raise FileNotFoundError(f"PDF file not found: {self.pdf_path}")
        
        self._doc: Optional[fitz.Document] = None
    
    def __enter__(self) -> PDFExtractor:
        """Context manager entry."""
        self.open()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        """Context manager exit."""
        self.close()
    
    def open(self) -> None:
        """Open the PDF document."""
        self._doc = fitz.open(str(self.pdf_path))
    
    def close(self) -> None:
        """Close the PDF document."""
        if self._doc:
            self._doc.close()
            self._doc = None
    
    @property
    def doc(self) -> fitz.Document:
        """Get the opened document, opening if necessary."""
        if self._doc is None:
            self.open()
        return self._doc
    
    def extract(self) -> DocumentData:
        """
        Extract all text blocks from the PDF.
        
        Returns:
            DocumentData containing all pages and blocks.
        """
        document = DocumentData(source_file=str(self.pdf_path))
        
        for page_num in range(len(self.doc)):
            page = self.doc[page_num]
            page_data = self._extract_page(page, page_num)
            document.pages.append(page_data)
        
        return document
    
    def _extract_page(self, page: fitz.Page, page_num: int) -> PageData:
        """
        Extract blocks from a single page.
        
        Args:
            page: The PyMuPDF page object.
            page_num: Zero-indexed page number.
            
        Returns:
            PageData with extracted blocks.
        """
        # Get page dimensions
        rect = page.rect
        page_data = PageData(
            page_number=page_num + 1,  # 1-indexed for output
            width=rect.width,
            height=rect.height,
            rotation=page.rotation,
        )
        
        # Extract using dict mode for structured data
        page_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
        
        blocks = []
        block_counter = 0
        
        for block in page_dict.get("blocks", []):
            # Skip image blocks (type 1)
            if block.get("type") != 0:
                continue
            
            block_data = self._extract_block(block, page_num, block_counter)
            if block_data and block_data.text.strip():  # Skip empty blocks
                blocks.append(block_data)
                block_counter += 1
        
        # Sort blocks deterministically: top-to-bottom, then left-to-right
        blocks = self._sort_blocks(blocks)
        
        # Reassign block_ids after sorting to ensure stability
        for idx, block in enumerate(blocks):
            block.block_id = f"p{page_num + 1}_b{idx}"
        
        page_data.blocks = blocks
        return page_data
    
    def _extract_block(
        self, 
        block: dict[str, Any], 
        page_num: int, 
        block_idx: int
    ) -> Optional[BlockData]:
        """
        Extract data from a single block.
        
        Args:
            block: The block dictionary from PyMuPDF.
            page_num: Zero-indexed page number.
            block_idx: Block index within page.
            
        Returns:
            BlockData or None if block is empty/invalid.
        """
        bbox = tuple(block.get("bbox", (0, 0, 0, 0)))
        
        # Collect text and font info from all lines and spans
        text_parts = []
        font_info = self._collect_font_info(block)
        line_heights = []
        
        for line in block.get("lines", []):
            line_bbox = line.get("bbox", (0, 0, 0, 0))
            line_height = line_bbox[3] - line_bbox[1]
            if line_height > 0:
                line_heights.append(line_height)
            
            line_text_parts = []
            for span in line.get("spans", []):
                span_text = span.get("text", "")
                line_text_parts.append(span_text)
            
            # Join spans within a line (no separator)
            text_parts.append("".join(line_text_parts))
        
        # Join lines with newline to preserve structure
        text = "\n".join(text_parts)
        
        if not text:
            return None
        
        # Calculate average line height
        avg_line_height = (
            sum(line_heights) / len(line_heights) 
            if line_heights else None
        )
        
        # Determine writing direction from the first line
        writing_dir = self._get_writing_direction(block)
        
        return BlockData(
            block_id=f"p{page_num + 1}_b{block_idx}",  # Will be reassigned after sorting
            bbox=bbox,
            text=text,
            font_name=font_info["font_name"],
            font_size=font_info["font_size"],
            color=font_info["color"],
            writing_direction=writing_dir,
            line_height=avg_line_height,
        )
    
    def _collect_font_info(self, block: dict[str, Any]) -> dict[str, Any]:
        """
        Collect predominant font information from block spans.
        
        Uses the most common font/size/color across all spans.
        
        Args:
            block: The block dictionary.
            
        Returns:
            Dictionary with font_name, font_size, color.
        """
        font_names: dict[str, int] = {}
        font_sizes: dict[float, int] = {}
        colors: dict[str, int] = {}
        
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                # Count font names
                fname = span.get("font", "Helvetica")
                font_names[fname] = font_names.get(fname, 0) + len(span.get("text", ""))
                
                # Count font sizes (round to 3 decimals for consistency)
                fsize = round(span.get("size", 12.0), 3)
                font_sizes[fsize] = font_sizes.get(fsize, 0) + len(span.get("text", ""))
                
                # Count colors
                color_int = span.get("color", 0)
                color_hex = self._int_to_hex_color(color_int)
                colors[color_hex] = colors.get(color_hex, 0) + len(span.get("text", ""))
        
        # Get most common values
        font_name = max(font_names, key=font_names.get) if font_names else "Helvetica"
        font_size = max(font_sizes, key=font_sizes.get) if font_sizes else 12.0
        color = max(colors, key=colors.get) if colors else "#000000"
        
        return {
            "font_name": font_name,
            "font_size": font_size,
            "color": color,
        }
    
    def _int_to_hex_color(self, color_int: int) -> str:
        """
        Convert integer color to hex string.
        
        Args:
            color_int: Color as integer (RGB).
            
        Returns:
            Hex color string like "#000000".
        """
        # Ensure non-negative
        color_int = max(0, color_int)
        
        # Extract RGB components
        r = (color_int >> 16) & 0xFF
        g = (color_int >> 8) & 0xFF
        b = color_int & 0xFF
        
        return f"#{r:02x}{g:02x}{b:02x}"
    
    def _get_writing_direction(self, block: dict[str, Any]) -> str:
        """
        Determine writing direction from block structure.
        
        Args:
            block: The block dictionary.
            
        Returns:
            Writing direction: "ltr", "rtl", or "ttb".
        """
        lines = block.get("lines", [])
        if not lines:
            return "ltr"
        
        # Check first line's direction
        first_line = lines[0]
        wdir = first_line.get("dir", (1, 0))
        
        if isinstance(wdir, (list, tuple)) and len(wdir) >= 2:
            # (1, 0) = left-to-right
            # (-1, 0) = right-to-left  
            # (0, 1) = top-to-bottom
            if wdir[0] < 0:
                return "rtl"
            elif wdir[0] == 0 and wdir[1] != 0:
                return "ttb"
        
        return "ltr"
    
    def _sort_blocks(self, blocks: list[BlockData]) -> list[BlockData]:
        """
        Sort blocks deterministically: top-to-bottom, then left-to-right.
        
        Args:
            blocks: List of BlockData to sort.
            
        Returns:
            Sorted list of BlockData.
        """
        def sort_key(block: BlockData) -> tuple[float, float]:
            # Round to 3 decimals for determinism
            # Primary: y0 (top position)
            # Secondary: x0 (left position)
            return (round(block.bbox[1], 3), round(block.bbox[0], 3))
        
        return sorted(blocks, key=sort_key)
    
    def save_json(self, output_path: Union[str, Path]) -> None:
        """
        Extract and save to JSON file.
        
        Args:
            output_path: Path for output JSON file.
        """
        document = self.extract()
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(document.to_json(), encoding="utf-8")


def extract_pdf_layout(
    pdf_path: Union[str, Path], 
    output_path: Optional[Union[str, Path]] = None
) -> DocumentData:
    """
    Convenience function to extract layout from a PDF.
    
    Args:
        pdf_path: Path to input PDF.
        output_path: Optional path to save JSON output.
        
    Returns:
        DocumentData with extracted layout.
    """
    with PDFExtractor(pdf_path) as extractor:
        document = extractor.extract()
        
        if output_path:
            output_path = Path(output_path)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_text(document.to_json(), encoding="utf-8")
        
        return document
