"""
Direct PDF Pipeline - Fast Redact+Insert Approach.

The fastest PDF translation approach - directly manipulates PDF without 
intermediate Office conversion. Uses PyMuPDF to redact original text and 
insert translations at exact positions.

Performance: ~2-3 seconds for 10-page PDF (vs ~30s for Office roundtrip)

Workflow:
1. PDF → Extract text with positions
2. Generate Moses/Tagged translation format
3. Redact original text blocks (white overlay)
4. Insert translated text at same positions
5. Embed Unicode fonts for target language
"""

from __future__ import annotations

import json
import re
import sys
import time
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Optional

import fitz  # PyMuPDF

from .base import (
    TranslationPipeline,
    PipelineConfig,
    PipelineType,
    ExtractResult,
    MergeResult,
)


class RenderMethod(Enum):
    """Text rendering method for PDF rebuilding."""
    REDACT_INSERT = "redact_insert"   # Fast: redact area + insert text (default)
    LINE_BY_LINE = "line_by_line"     # Legacy: line-by-line replacement
    TEXTBOX_REFLOW = "textbox_reflow" # Reflow text within bounding box


# Unicode font paths for different platforms
UNICODE_FONTS = [
    # Windows - Hindi/Devanagari
    Path("C:/Windows/Fonts/Nirmala.ttc"),
    Path("C:/Windows/Fonts/NirmalaS.ttc"),
    Path("C:/Windows/Fonts/mangal.ttf"),
    # Windows - General Unicode
    Path("C:/Windows/Fonts/arial.ttf"),
    Path("C:/Windows/Fonts/arialuni.ttf"),
    Path("C:/Windows/Fonts/seguiemj.ttf"),
    # Linux
    Path("/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf"),
    Path("/usr/share/fonts/truetype/freefont/FreeSans.ttf"),
    # macOS
    Path("/System/Library/Fonts/Supplemental/Arial Unicode.ttf"),
    Path("/Library/Fonts/Arial Unicode.ttf"),
]


def _find_unicode_font() -> Optional[Path]:
    """Find an available Unicode font on the system."""
    for font_path in UNICODE_FONTS:
        if font_path.exists():
            return font_path
    return None


def _contains_non_latin(text: str) -> bool:
    """Check if text contains non-Latin characters (requires Unicode font)."""
    return any(ord(ch) > 127 for ch in text)


@dataclass
class TextBlock:
    """Represents a text block with full position and formatting data."""
    block_id: str
    page_num: int
    bbox: tuple[float, float, float, float]  # (x0, y0, x1, y1)
    text: str
    font_name: str
    font_size: float
    font_flags: int  # Bold=16, Italic=2
    color: int  # RGB as integer
    lines: list[dict] = field(default_factory=list)  # Line-level data
    
    @property
    def is_bold(self) -> bool:
        return bool(self.font_flags & 16)
    
    @property
    def is_italic(self) -> bool:
        return bool(self.font_flags & 2)
    
    @property
    def width(self) -> float:
        return self.bbox[2] - self.bbox[0]
    
    @property
    def height(self) -> float:
        return self.bbox[3] - self.bbox[1]
    
    def to_dict(self) -> dict:
        return {
            "block_id": self.block_id,
            "page_num": self.page_num,
            "bbox": list(self.bbox),
            "text": self.text,
            "font_name": self.font_name,
            "font_size": self.font_size,
            "font_flags": self.font_flags,
            "color": self.color,
            "lines": self.lines,
        }


@dataclass
class DirectPDFConfig(PipelineConfig):
    """Configuration for Direct PDF pipeline."""
    
    pipeline_type: PipelineType = PipelineType.DIRECT_PDF
    
    # Render method
    render_method: RenderMethod = RenderMethod.REDACT_INSERT
    
    # Font settings
    unicode_font_path: Optional[Path] = field(default_factory=_find_unicode_font)
    fallback_font: str = "helv"
    min_font_size: float = 6.0
    font_step: float = 0.5
    
    # Redaction settings
    redact_color: tuple[float, float, float] = (1.0, 1.0, 1.0)  # White
    
    # Text fitting
    auto_scale: bool = True  # Auto-scale font for longer translations
    max_font_reduction: float = 0.5  # Maximum font size reduction (50%)
    
    # Output format
    encoding: str = "utf-8"
    keep_intermediate: bool = True


class DirectPDFPipeline(TranslationPipeline):
    """
    Fast Direct PDF manipulation pipeline using Redact+Insert.
    
    This is the fastest approach for PDF translation:
    1. Extracts text blocks with exact positions
    2. Redacts (removes) original text with white overlay
    3. Inserts translated text at same positions
    4. Embeds Unicode fonts for non-Latin scripts
    
    Performance:
    - 10-page PDF: ~2-3 seconds
    - No LibreOffice or Office conversion needed
    - Memory efficient (page-by-page processing)
    
    Best for:
    - Simple text PDFs
    - Documents where layout must be preserved exactly
    - High-volume translation workflows
    """
    
    def __init__(self, config: DirectPDFConfig):
        super().__init__(config)
        self.config: DirectPDFConfig = config
        self._unicode_font: Optional[fitz.Font] = None
        self._load_unicode_font()
    
    def _load_unicode_font(self) -> None:
        """Load Unicode font for non-Latin text rendering."""
        if self.config.unicode_font_path and self.config.unicode_font_path.exists():
            try:
                self._unicode_font = fitz.Font(
                    fontfile=str(self.config.unicode_font_path)
                )
            except Exception as e:
                print(f"  Warning: Could not load Unicode font: {e}")
    
    @property
    def name(self) -> str:
        return "Direct PDF (Redact+Insert)"
    
    @property
    def description(self) -> str:
        return "Fast PDF text replacement - redact original, insert translation"
    
    def extract(self, input_path: Path) -> ExtractResult:
        """
        Extract text blocks with full position metadata.
        
        Extracts:
        - Text content with exact bounding boxes
        - Font name, size, flags (bold/italic)
        - Color information
        - Line-level breakdown for precise positioning
        
        Generates:
        - Layout JSON with complete position data
        - Moses format (source.txt, target.txt)
        - Tagged format for fallback
        """
        paths = self.derive_paths(input_path)
        
        print(f"  Opening PDF: {input_path.name}")
        start_time = time.time()
        
        # Extract all text blocks with positions
        blocks = self._extract_text_blocks(input_path)
        
        print(f"  Extracted {len(blocks)} text blocks")
        
        # Build layout JSON
        layout = {
            "source_file": str(input_path),
            "pipeline": "direct_pdf",
            "render_method": self.config.render_method.value,
            "target_language": self.config.target_language,
            "encoding": self.config.encoding,
            "blocks": [b.to_dict() for b in blocks],
            "block_order": [b.block_id for b in blocks],
        }
        
        paths["layout"].write_text(
            json.dumps(layout, indent=2, ensure_ascii=False),
            encoding="utf-8"
        )
        
        # Generate Moses format
        source_path = input_path.parent / f"{input_path.name}_source.txt"
        target_path = input_path.parent / f"{input_path.name}_target.txt"
        
        source_lines = []
        target_lines = []
        for block in blocks:
            # Normalize: replace newlines with <br> marker
            text = block.text.replace('\n', ' <br> ')
            text = ' '.join(text.split())  # Clean multiple spaces
            source_lines.append(text)
            target_lines.append(text)  # Template for translation
        
        source_path.write_text('\n'.join(source_lines), encoding=self.config.encoding)
        target_path.write_text('\n'.join(target_lines), encoding=self.config.encoding)
        
        # Also generate tagged format
        tagged_lines = []
        for i, block in enumerate(blocks):
            text = block.text.replace('\n', '\\n')
            tagged_lines.append(f"<{i}>{text}</{i}>")
        
        paths["translate"].write_text('\n'.join(tagged_lines), encoding=self.config.encoding)
        paths["translated"].write_text('\n'.join(tagged_lines), encoding=self.config.encoding)
        
        elapsed = time.time() - start_time
        print(f"  Extraction complete ({elapsed:.1f}s)")
        print(f"  Moses format: {source_path.name}, {target_path.name}")
        
        return ExtractResult(
            layout_path=paths["layout"],
            translate_path=paths["translate"],
            translated_template_path=paths["translated"],
            extra_files={
                "source_txt": source_path,
                "target_txt": target_path,
            },
        )
    
    def _extract_text_blocks(self, pdf_path: Path) -> list[TextBlock]:
        """Extract text blocks with full metadata from PDF."""
        blocks = []
        doc = fitz.open(str(pdf_path))
        
        try:
            for page_num, page in enumerate(doc):
                page_dict = page.get_text("dict")
                
                for block_idx, block in enumerate(page_dict.get("blocks", [])):
                    if block.get("type") != 0:  # Skip non-text blocks
                        continue
                    
                    # Collect text and formatting from lines/spans
                    text_parts = []
                    lines_data = []
                    dominant_font = "Helvetica"
                    dominant_size = 12.0
                    dominant_flags = 0
                    dominant_color = 0
                    
                    for line in block.get("lines", []):
                        line_text = ""
                        line_info = {
                            "bbox": line.get("bbox", [0, 0, 0, 0]),
                            "spans": [],
                        }
                        
                        for span in line.get("spans", []):
                            span_text = span.get("text", "")
                            line_text += span_text
                            
                            # Track dominant formatting
                            if len(span_text.strip()) > 0:
                                dominant_font = span.get("font", dominant_font)
                                dominant_size = span.get("size", dominant_size)
                                dominant_flags = span.get("flags", dominant_flags)
                                dominant_color = span.get("color", dominant_color)
                            
                            line_info["spans"].append({
                                "text": span_text,
                                "font": span.get("font", ""),
                                "size": span.get("size", 12),
                                "flags": span.get("flags", 0),
                                "color": span.get("color", 0),
                            })
                        
                        if line_text.strip():
                            text_parts.append(line_text)
                            lines_data.append(line_info)
                    
                    full_text = '\n'.join(text_parts)
                    if not full_text.strip():
                        continue
                    
                    blocks.append(TextBlock(
                        block_id=f"p{page_num}_b{block_idx}",
                        page_num=page_num,
                        bbox=tuple(block.get("bbox", [0, 0, 0, 0])),
                        text=full_text,
                        font_name=dominant_font,
                        font_size=dominant_size,
                        font_flags=dominant_flags,
                        color=dominant_color,
                        lines=lines_data,
                    ))
        finally:
            doc.close()
        
        return blocks
    
    def merge(
        self,
        input_path: Path,
        output_path: Path,
        translated_path: Path,
        layout_path: Path,
    ) -> MergeResult:
        """
        Apply translations using fast Redact+Insert method.
        
        Steps:
        1. Load layout with position data
        2. Parse translations (Moses or tagged format)
        3. Open PDF for modification
        4. For each page:
           a. Redact original text blocks (white overlay)
           b. Apply all redactions at once
           c. Insert translated text at same positions
        5. Save modified PDF
        """
        print(f"  Loading layout...")
        layout = json.loads(layout_path.read_text(encoding="utf-8"))
        
        # Parse translations
        translations = self._parse_translations(layout_path, translated_path, layout)
        print(f"  Parsed {len(translations)} translations")
        
        start_time = time.time()
        
        # Open PDF
        doc = fitz.open(str(input_path))
        total_pages = len(doc)
        
        try:
            # Group blocks by page
            blocks_by_page: dict[int, list[dict]] = {}
            for block_data in layout.get("blocks", []):
                page_num = block_data["page_num"]
                if page_num not in blocks_by_page:
                    blocks_by_page[page_num] = []
                
                block_id = block_data["block_id"]
                if block_id in translations:
                    block_data["translation"] = translations[block_id]
                    blocks_by_page[page_num].append(block_data)
            
            # Process each page
            for page_num in sorted(blocks_by_page.keys()):
                page_blocks = blocks_by_page[page_num]
                if not page_blocks:
                    continue
                
                sys.stdout.write(f"\r  Processing page {page_num + 1}/{total_pages}...")
                sys.stdout.flush()
                
                page = doc[page_num]
                
                # Step 1: Add redaction annotations for all blocks
                for block in page_blocks:
                    bbox = block["bbox"]
                    rect = fitz.Rect(bbox[0], bbox[1], bbox[2], bbox[3])
                    
                    # Add padding to ensure complete coverage
                    rect.x0 -= 1
                    rect.y0 -= 1
                    rect.x1 += 1
                    rect.y1 += 1
                    
                    annot = page.add_redact_annot(rect)
                    annot.set_colors(fill=self.config.redact_color)
                
                # Step 2: Apply all redactions at once (efficient)
                page.apply_redactions()
                
                # Step 3: Insert translated text
                for block in page_blocks:
                    self._insert_text_block(page, block)
            
            # Save modified PDF
            sys.stdout.write(f"\r  Saving PDF...                    \n")
            sys.stdout.flush()
            
            doc.save(str(output_path), garbage=4, deflate=True)
            
        finally:
            doc.close()
        
        elapsed = time.time() - start_time
        output_size = output_path.stat().st_size / 1024
        print(f"  Output: {output_path.name} ({output_size:.1f} KB)")
        print(f"  Completed in {elapsed:.1f}s")
        
        return MergeResult(
            output_path=output_path,
            blocks_processed=len(translations),
        )
    
    def _parse_translations(
        self,
        layout_path: Path,
        translated_path: Path,
        layout: dict,
    ) -> dict[str, str]:
        """Parse translations from Moses or tagged format.
        
        Priority:
        1. Moses format (target.txt) - if it differs from source.txt
        2. Tagged format (translated.txt) - fallback
        
        This ensures we don't use untranslated Moses templates.
        """
        block_order = layout.get("block_order", [])
        
        # Check Moses format files
        base_name = layout_path.stem.replace('_layout', '')
        source_path = layout_path.parent / f"{base_name}_source.txt"
        target_path = layout_path.parent / f"{base_name}_target.txt"
        
        # Use Moses format only if target differs from source (i.e., was translated)
        if target_path.exists() and source_path.exists():
            source_content = source_path.read_text(encoding=self.config.encoding)
            target_content = target_path.read_text(encoding=self.config.encoding)
            
            if source_content != target_content:
                # Target was modified, use Moses format
                return self._parse_moses_translations(target_path, block_order)
        elif target_path.exists() and not source_path.exists():
            # Only target exists, use it
            return self._parse_moses_translations(target_path, block_order)
        
        # Fall back to tagged format
        return self._parse_tagged_translations(translated_path, block_order)
    
    def _parse_moses_translations(
        self,
        target_path: Path,
        block_order: list[str],
    ) -> dict[str, str]:
        """Parse Moses format translations (one per line)."""
        target_lines = target_path.read_text(encoding=self.config.encoding).split('\n')
        
        translations = {}
        for i, text in enumerate(target_lines):
            if i < len(block_order):
                # Convert Moses format back
                text = text.replace(' <br> ', '\n')
                translations[block_order[i]] = text
        
        return translations
    
    def _parse_tagged_translations(
        self,
        translated_path: Path,
        block_order: list[str],
    ) -> dict[str, str]:
        """Parse tagged format translations."""
        content = translated_path.read_text(encoding=self.config.encoding)
        translations = {}
        
        pattern = r'<(\d+)>(.*?)</\1>'
        for match in re.finditer(pattern, content, re.DOTALL):
            idx = int(match.group(1))
            text = match.group(2).replace('\\n', '\n')
            
            if idx < len(block_order):
                translations[block_order[idx]] = text
        
        return translations
    
    def _insert_text_block(self, page: fitz.Page, block: dict) -> None:
        """Insert translated text at block position with formatting.
        
        Uses TextWriter for Unicode text to ensure proper character rendering.
        Falls back to insert_textbox for Latin-only text.
        """
        bbox = block["bbox"]
        text = block.get("translation", block.get("text", ""))
        font_size = block.get("font_size", 12)
        color = block.get("color", 0)
        
        if not text.strip():
            return
        
        # Create rect for text insertion
        rect = fitz.Rect(bbox[0], bbox[1], bbox[2], bbox[3])
        
        # Check if Unicode font needed
        use_unicode = _contains_non_latin(text) and self._unicode_font is not None
        
        # Calculate fitting font size
        if self.config.auto_scale:
            font_size = self._calculate_fitting_font_size(
                text, rect, font_size, use_unicode
            )
        
        # Convert color to RGB tuple
        if isinstance(color, int) and color != 0:
            r = ((color >> 16) & 0xFF) / 255
            g = ((color >> 8) & 0xFF) / 255
            b = (color & 0xFF) / 255
            text_color = (r, g, b)
        else:
            text_color = (0, 0, 0)  # Black default
        
        # Insert text
        if use_unicode:
            # Use TextWriter for proper Unicode rendering
            self._insert_unicode_text(page, rect, text, font_size, text_color)
        else:
            # Use built-in font with insert_textbox
            page.insert_textbox(
                rect,
                text,
                fontname=self.config.fallback_font,
                fontsize=font_size,
                color=text_color,
                align=fitz.TEXT_ALIGN_LEFT,
            )
    
    def _insert_unicode_text(
        self,
        page: fitz.Page,
        rect: fitz.Rect,
        text: str,
        font_size: float,
        color: tuple,
    ) -> None:
        """Insert Unicode text using TextWriter for proper rendering.
        
        TextWriter properly embeds Unicode fonts and preserves character data,
        unlike insert_textbox which may lose Unicode information.
        """
        if self._unicode_font is None:
            return
        
        # Create TextWriter
        tw = fitz.TextWriter(page.rect, color=color)
        
        # Calculate line metrics
        line_height = font_size * 1.2
        max_width = rect.width
        
        # Start position (top-left of rect, adjusted for baseline)
        x_start = rect.x0
        y_pos = rect.y0 + font_size  # Start below top edge by font height
        
        # Split text into lines and wrap
        lines = text.split('\n')
        
        for line in lines:
            if y_pos > rect.y1:  # Stop if we exceed the box
                break
            
            # Wrap line if needed
            wrapped_lines = self._wrap_line(line, max_width, font_size)
            
            for wrapped in wrapped_lines:
                if y_pos > rect.y1:
                    break
                
                # Append text at position
                try:
                    tw.append(
                        fitz.Point(x_start, y_pos),
                        wrapped,
                        font=self._unicode_font,
                        fontsize=font_size,
                    )
                except Exception:
                    pass  # Skip if text can't be rendered
                
                y_pos += line_height
        
        # Write all text to page
        tw.write_text(page)
    
    def _wrap_line(self, text: str, max_width: float, font_size: float) -> list[str]:
        """Wrap a single line of text to fit within max_width.
        
        Uses approximate character width calculation.
        """
        if not text:
            return [""]
        
        # Approximate average character width (varies by font, this is rough)
        avg_char_width = font_size * 0.5
        max_chars = int(max_width / avg_char_width) if avg_char_width > 0 else len(text)
        
        if max_chars <= 0:
            max_chars = 1
        
        if len(text) <= max_chars:
            return [text]
        
        # Word wrap
        words = text.split(' ')
        lines = []
        current_line = ""
        
        for word in words:
            test_line = f"{current_line} {word}".strip() if current_line else word
            if len(test_line) <= max_chars:
                current_line = test_line
            else:
                if current_line:
                    lines.append(current_line)
                # If single word is too long, split it
                if len(word) > max_chars:
                    while len(word) > max_chars:
                        lines.append(word[:max_chars])
                        word = word[max_chars:]
                    current_line = word
                else:
                    current_line = word
        
        if current_line:
            lines.append(current_line)
        
        return lines if lines else [text]
    
    def _calculate_fitting_font_size(
        self,
        text: str,
        rect: fitz.Rect,
        original_size: float,
        use_unicode: bool,
    ) -> float:
        """Calculate font size that fits text in bounding box."""
        min_size = max(self.config.min_font_size, 
                       original_size * self.config.max_font_reduction)
        
        # Estimate text dimensions
        # Average character width is roughly 0.5 * font_size
        # Line height is roughly 1.2 * font_size
        
        lines = text.split('\n')
        max_line_length = max(len(line) for line in lines) if lines else 1
        num_lines = len(lines)
        
        # Calculate required size based on width
        char_width_factor = 0.55  # Approximate
        required_width = max_line_length * original_size * char_width_factor
        width_ratio = rect.width / required_width if required_width > 0 else 1
        
        # Calculate required size based on height
        line_height_factor = 1.3
        required_height = num_lines * original_size * line_height_factor
        height_ratio = rect.height / required_height if required_height > 0 else 1
        
        # Use smaller ratio to ensure fit
        scale = min(width_ratio, height_ratio, 1.0)
        new_size = original_size * scale
        
        return max(min_size, new_size)
    
    def get_translation_prompt(self, block_count: int) -> str:
        """Get translation prompt for Moses format."""
        return f"""
Direct PDF Translation - {block_count} text blocks

Files:
- *_source.txt: Source text (one block per line)
- *_target.txt: Translate this file

Instructions:
1. Open *_target.txt in a text editor
2. Translate each line to {self.config.target_language}
3. Keep the SAME number of lines (do NOT add or remove lines)
4. Preserve <br> markers (they represent line breaks within blocks)
5. Save as {self.config.encoding} encoding

Example:
Source: Hello world <br> Welcome to our service
Target: नमस्ते दुनिया <br> हमारी सेवा में आपका स्वागत है

After translation, run: python main.py merge input.pdf
"""


def create_direct_pdf_pipeline(
    target_language: str = "Hindi",
    render_method: RenderMethod = RenderMethod.REDACT_INSERT,
    min_font_size: float = 6.0,
    font_step: float = 0.5,
    auto_scale: bool = True,
) -> DirectPDFPipeline:
    """
    Factory function to create DirectPDF pipeline.
    
    Args:
        target_language: Target language for translation (default: Hindi)
        render_method: Text rendering method (default: REDACT_INSERT - fastest)
        min_font_size: Minimum font size after scaling (default: 6.0)
        font_step: Font size reduction step (default: 0.5)
        auto_scale: Auto-scale font for longer translations (default: True)
    """
    config = DirectPDFConfig(
        target_language=target_language,
        render_method=render_method,
        min_font_size=min_font_size,
        font_step=font_step,
        auto_scale=auto_scale,
    )
    return DirectPDFPipeline(config)
