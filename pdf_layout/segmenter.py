"""
PDF Segmentation Module.

Handles segmentation logic for text blocks, preserving reading order
and atomic block structure.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterator, Optional, Union

from .extractor import DocumentData, PageData, BlockData


@dataclass
class SegmentedBlock:
    """Represents a segmented text block ready for translation."""
    
    block_id: str
    page_number: int
    bbox: tuple[float, float, float, float]
    text: str
    font_name: str
    font_size: float
    color: str
    writing_direction: str
    line_height: Optional[float]
    segment_index: int  # Global index in reading order
    
    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary for serialization."""
        return {
            "block_id": self.block_id,
            "page_number": self.page_number,
            "bbox": list(self.bbox),
            "text": self.text,
            "font_name": self.font_name,
            "font_size": self.font_size,
            "color": self.color,
            "writing_direction": self.writing_direction,
            "line_height": self.line_height,
            "segment_index": self.segment_index,
        }


@dataclass  
class SegmentedDocument:
    """Represents a segmented document with all blocks in reading order."""
    
    source_file: str
    total_pages: int
    segments: list[SegmentedBlock] = field(default_factory=list)
    
    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary for serialization."""
        return {
            "source_file": self.source_file,
            "total_pages": self.total_pages,
            "total_segments": len(self.segments),
            "segments": [seg.to_dict() for seg in self.segments],
        }
    
    def to_json(self, indent: int = 2) -> str:
        """Convert to JSON string."""
        return json.dumps(self.to_dict(), indent=indent, ensure_ascii=False)
    
    def iter_by_page(self) -> Iterator[tuple[int, list[SegmentedBlock]]]:
        """
        Iterate segments grouped by page.
        
        Yields:
            Tuples of (page_number, list of segments on that page).
        """
        current_page = None
        current_segments: list[SegmentedBlock] = []
        
        for segment in self.segments:
            if segment.page_number != current_page:
                if current_page is not None:
                    yield current_page, current_segments
                current_page = segment.page_number
                current_segments = []
            current_segments.append(segment)
        
        if current_page is not None:
            yield current_page, current_segments
    
    def get_block_ids(self) -> list[str]:
        """Get list of all block IDs in reading order."""
        return [seg.block_id for seg in self.segments]
    
    def get_texts(self) -> dict[str, str]:
        """Get mapping of block_id to text."""
        return {seg.block_id: seg.text for seg in self.segments}


class PDFSegmenter:
    """
    Segments extracted PDF layout into atomic translation units.
    
    Rules:
    - Each block is atomic (not split or merged)
    - Reading order is preserved (top-to-bottom, left-to-right)
    - Empty blocks are removed
    - Whitespace is preserved exactly
    """
    
    def __init__(self, document_data: Union[DocumentData, dict[str, Any], str, Path]):
        """
        Initialize segmenter with document data.
        
        Args:
            document_data: DocumentData object, dict, JSON string, or path to JSON file.
        """
        self.document = self._load_document(document_data)
    
    def _load_document(
        self, 
        data: Union[DocumentData, dict[str, Any], str, Path]
    ) -> DocumentData:
        """
        Load document data from various input types.
        
        Args:
            data: Document data in various formats.
            
        Returns:
            DocumentData object.
        """
        if isinstance(data, DocumentData):
            return data
        
        if isinstance(data, (str, Path)):
            path = Path(data)
            if path.exists() and path.suffix == ".json":
                # Load from file
                data = json.loads(path.read_text(encoding="utf-8"))
            elif isinstance(data, str) and data.strip().startswith("{"):
                # JSON string
                data = json.loads(data)
            else:
                raise ValueError(f"Invalid input: {data}")
        
        # Convert dict to DocumentData
        return self._dict_to_document(data)
    
    def _dict_to_document(self, data: dict[str, Any]) -> DocumentData:
        """
        Convert dictionary to DocumentData.
        
        Args:
            data: Dictionary representation.
            
        Returns:
            DocumentData object.
        """
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
    
    def segment(self) -> SegmentedDocument:
        """
        Segment the document into atomic translation units.
        
        Returns:
            SegmentedDocument with all blocks in reading order.
        """
        segmented = SegmentedDocument(
            source_file=self.document.source_file,
            total_pages=len(self.document.pages),
        )
        
        segment_index = 0
        
        for page in self.document.pages:
            for block in page.blocks:
                # Skip empty blocks
                if not self._is_valid_block(block):
                    continue
                
                segment = SegmentedBlock(
                    block_id=block.block_id,
                    page_number=page.page_number,
                    bbox=block.bbox,
                    text=block.text,
                    font_name=block.font_name,
                    font_size=block.font_size,
                    color=block.color,
                    writing_direction=block.writing_direction,
                    line_height=block.line_height,
                    segment_index=segment_index,
                )
                
                segmented.segments.append(segment)
                segment_index += 1
        
        return segmented
    
    def _is_valid_block(self, block: BlockData) -> bool:
        """
        Check if a block is valid for segmentation.
        
        Args:
            block: Block to validate.
            
        Returns:
            True if block should be included.
        """
        # Block must have non-empty text (but preserve whitespace-only if intentional)
        if not block.text:
            return False
        
        # Block must have valid bounding box
        bbox = block.bbox
        if len(bbox) != 4:
            return False
        
        # Bounding box must have positive dimensions
        width = bbox[2] - bbox[0]
        height = bbox[3] - bbox[1]
        if width <= 0 or height <= 0:
            return False
        
        return True
    
    def create_translation_template(self) -> dict[str, str]:
        """
        Create a template dictionary for translation.
        
        Returns:
            Dictionary mapping block_id to original text.
        """
        segmented = self.segment()
        return {seg.block_id: seg.text for seg in segmented.segments}
    
    def save_template(self, output_path: Union[str, Path]) -> None:
        """
        Save translation template to JSON file.
        
        Args:
            output_path: Path for output file.
        """
        template = self.create_translation_template()
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(
            json.dumps(template, indent=2, ensure_ascii=False),
            encoding="utf-8"
        )


def segment_document(
    document_data: Union[DocumentData, dict[str, Any], str, Path]
) -> SegmentedDocument:
    """
    Convenience function to segment a document.
    
    Args:
        document_data: Document data in various formats.
        
    Returns:
        SegmentedDocument with atomic blocks.
    """
    segmenter = PDFSegmenter(document_data)
    return segmenter.segment()


def create_translation_template(
    document_data: Union[DocumentData, dict[str, Any], str, Path],
    output_path: Optional[Union[str, Path]] = None
) -> dict[str, str]:
    """
    Create translation template from document data.
    
    Args:
        document_data: Document data in various formats.
        output_path: Optional path to save template.
        
    Returns:
        Dictionary mapping block_id to text.
    """
    segmenter = PDFSegmenter(document_data)
    template = segmenter.create_translation_template()
    
    if output_path:
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(
            json.dumps(template, indent=2, ensure_ascii=False),
            encoding="utf-8"
        )
    
    return template
