"""
Source Format Detection Utility.

Detects the original Office format (DOCX/PPTX/XLSX) from PDF metadata
and layout characteristics.
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Optional

import fitz  # PyMuPDF


class SourceFormat(Enum):
    """Detected source format of PDF."""
    UNKNOWN = "unknown"
    WORD = "docx"       # Microsoft Word
    POWERPOINT = "pptx"  # Microsoft PowerPoint
    EXCEL = "xlsx"       # Microsoft Excel
    LIBREOFFICE_WRITER = "odt"
    LIBREOFFICE_IMPRESS = "odp"
    LIBREOFFICE_CALC = "ods"
    LATEX = "tex"
    GENERIC_PDF = "pdf"


@dataclass
class SourceInfo:
    """Information about the PDF source."""
    format: SourceFormat
    creator: Optional[str] = None
    producer: Optional[str] = None
    confidence: float = 0.0  # 0.0 to 1.0
    details: str = ""


# Keywords to detect source format from creator/producer strings
FORMAT_KEYWORDS = {
    SourceFormat.WORD: [
        "word", "microsoft word", "ms word",
        "libreoffice writer", "openoffice writer",
    ],
    SourceFormat.POWERPOINT: [
        "powerpoint", "microsoft powerpoint", "ms powerpoint",
        "libreoffice impress", "openoffice impress",
        "keynote",
    ],
    SourceFormat.EXCEL: [
        "excel", "microsoft excel", "ms excel",
        "libreoffice calc", "openoffice calc",
        "numbers",  # Apple Numbers
    ],
    SourceFormat.LATEX: [
        "latex", "pdflatex", "xelatex", "lualatex",
        "tex", "pdftex",
    ],
}


def detect_source_format(pdf_path: Path) -> SourceInfo:
    """
    Detect the source format of a PDF file.
    
    Uses multiple signals:
    1. PDF metadata (creator, producer)
    2. Page layout characteristics
    3. Content analysis
    
    Args:
        pdf_path: Path to PDF file.
        
    Returns:
        SourceInfo with detected format and confidence.
    """
    doc = fitz.open(str(pdf_path))
    
    try:
        metadata = doc.metadata
        creator = metadata.get("creator", "").lower()
        producer = metadata.get("producer", "").lower()
        
        # First try metadata detection
        format_from_metadata, confidence = _detect_from_metadata(creator, producer)
        
        if confidence >= 0.8:
            return SourceInfo(
                format=format_from_metadata,
                creator=metadata.get("creator"),
                producer=metadata.get("producer"),
                confidence=confidence,
                details=f"Detected from metadata: {metadata.get('creator', '')}",
            )
        
        # Fall back to layout analysis
        format_from_layout, layout_confidence = _detect_from_layout(doc)
        
        # Combine signals
        if format_from_metadata == format_from_layout:
            final_confidence = min(1.0, confidence + layout_confidence * 0.5)
        elif layout_confidence > confidence:
            format_from_metadata = format_from_layout
            final_confidence = layout_confidence
        else:
            final_confidence = confidence
        
        # Fallback to DOCX if format is still unknown
        final_format = format_from_metadata
        if final_format == SourceFormat.UNKNOWN:
            final_format = SourceFormat.WORD  # Default to DOCX as most common
            final_confidence = 0.3  # Low confidence for fallback
            details = "Fallback to DOCX (no metadata detected)"
        else:
            details = _build_details(metadata, format_from_layout)
        
        return SourceInfo(
            format=final_format,
            creator=metadata.get("creator"),
            producer=metadata.get("producer"),
            confidence=final_confidence,
            details=details,
        )
        
    finally:
        doc.close()


def _detect_from_metadata(creator: str, producer: str) -> tuple[SourceFormat, float]:
    """Detect format from metadata strings."""
    combined = f"{creator} {producer}".lower()
    
    for format_type, keywords in FORMAT_KEYWORDS.items():
        for keyword in keywords:
            if keyword in combined:
                # Higher confidence for more specific matches
                confidence = 0.9 if "microsoft" in combined else 0.7
                return format_type, confidence
    
    return SourceFormat.UNKNOWN, 0.0


def _detect_from_layout(doc: fitz.Document) -> tuple[SourceFormat, float]:
    """Detect format from layout characteristics."""
    if len(doc) == 0:
        return SourceFormat.UNKNOWN, 0.0
    
    # Analyze first few pages
    pages_to_analyze = min(5, len(doc))
    
    landscape_count = 0
    avg_blocks_per_page = 0
    has_tables = False
    avg_font_size = 0
    
    for i in range(pages_to_analyze):
        page = doc[i]
        rect = page.rect
        
        # Check orientation
        if rect.width > rect.height:
            landscape_count += 1
        
        # Count text blocks
        blocks = page.get_text("dict")["blocks"]
        text_blocks = [b for b in blocks if b.get("type") == 0]
        avg_blocks_per_page += len(text_blocks)
        
        # Check for tables (by looking for many small aligned blocks)
        if _has_table_structure(text_blocks):
            has_tables = True
        
        # Average font size
        for block in text_blocks:
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    avg_font_size += span.get("size", 12)
    
    avg_blocks_per_page /= pages_to_analyze
    landscape_ratio = landscape_count / pages_to_analyze
    
    # PowerPoint detection: landscape, few blocks, large fonts
    if landscape_ratio > 0.7 and avg_blocks_per_page < 15:
        return SourceFormat.POWERPOINT, 0.7
    
    # Excel detection: many small blocks in table structure
    if has_tables and avg_blocks_per_page > 50:
        return SourceFormat.EXCEL, 0.6
    
    # Word detection: portrait, moderate blocks
    if landscape_ratio < 0.3 and 5 < avg_blocks_per_page < 50:
        return SourceFormat.WORD, 0.5
    
    return SourceFormat.UNKNOWN, 0.0


def _has_table_structure(blocks: list) -> bool:
    """Check if blocks appear to be arranged in a table."""
    if len(blocks) < 10:
        return False
    
    # Check for vertical alignment (same x coordinates)
    x_positions = []
    for block in blocks:
        bbox = block.get("bbox", [0, 0, 0, 0])
        x_positions.append(round(bbox[0], 0))
    
    # If many blocks share x positions, likely a table
    from collections import Counter
    x_counts = Counter(x_positions)
    aligned_count = sum(1 for count in x_counts.values() if count > 2)
    
    return aligned_count >= 3


def _build_details(metadata: dict, layout_format: SourceFormat) -> str:
    """Build detailed detection explanation."""
    parts = []
    
    if metadata.get("creator"):
        parts.append(f"Creator: {metadata['creator']}")
    if metadata.get("producer"):
        parts.append(f"Producer: {metadata['producer']}")
    if layout_format != SourceFormat.UNKNOWN:
        parts.append(f"Layout suggests: {layout_format.value}")
    
    return "; ".join(parts)


def get_recommended_pipeline(source_info: SourceInfo) -> str:
    """
    Get recommended pipeline based on source format.
    
    Args:
        source_info: Detected source information.
        
    Returns:
        Recommended pipeline name.
    """
    if source_info.format in [SourceFormat.WORD, SourceFormat.LIBREOFFICE_WRITER]:
        return "docx"
    elif source_info.format in [SourceFormat.POWERPOINT, SourceFormat.LIBREOFFICE_IMPRESS]:
        return "pptx"
    elif source_info.format in [SourceFormat.EXCEL, SourceFormat.LIBREOFFICE_CALC]:
        return "xlsx"
    elif source_info.format == SourceFormat.UNKNOWN:
        return "docx"  # Fallback to DOCX as most common format
    else:
        return "direct"


def print_source_info(pdf_path: Path) -> None:
    """Print source format information for a PDF."""
    info = detect_source_format(pdf_path)
    
    print(f"Source Format Detection for: {pdf_path}")
    print(f"  Detected Format: {info.format.value}")
    print(f"  Confidence: {info.confidence:.0%}")
    print(f"  Creator: {info.creator or 'Unknown'}")
    print(f"  Producer: {info.producer or 'Unknown'}")
    print(f"  Details: {info.details}")
    print(f"  Recommended Pipeline: {get_recommended_pipeline(info)}")
