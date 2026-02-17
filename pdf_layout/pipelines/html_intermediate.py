"""
HTML Intermediate Pipeline - Layout-preserving PDF translation via HTML.

Converts PDF to HTML with CSS absolute positioning, enables translation,
then renders back to PDF. Preserves exact layout through CSS coordinates.

Workflow:
1. PDF → HTML (absolute positioned divs with CSS)
2. Extract text segments for translation
3. Translate text
4. Update HTML with translations
5. HTML → PDF (via WeasyPrint or browser)

Performance: ~2-5 seconds for typical PDF

Dependencies:
- PyMuPDF>=1.23.0 - PDF text extraction
- weasyprint>=60.0 (optional) - HTML to PDF rendering
- Or use browser print-to-PDF for HTML output
"""

from __future__ import annotations

import html
import json
import re
import sys
import time
from dataclasses import dataclass, field
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


@dataclass
class TextSegment:
    """A text segment with position information for HTML rendering."""
    seg_id: str
    page_num: int
    text: str
    x: float  # Left position (pt)
    y: float  # Top position (pt)
    width: float
    height: float
    font_name: str
    font_size: float
    color: str  # Hex color
    is_bold: bool = False
    is_italic: bool = False
    line_height: float = 1.2
    
    def to_dict(self) -> dict:
        return {
            "seg_id": self.seg_id,
            "page_num": self.page_num,
            "text": self.text,
            "x": self.x,
            "y": self.y,
            "width": self.width,
            "height": self.height,
            "font_name": self.font_name,
            "font_size": self.font_size,
            "color": self.color,
            "is_bold": self.is_bold,
            "is_italic": self.is_italic,
            "line_height": self.line_height,
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> "TextSegment":
        return cls(**data)


@dataclass
class PageInfo:
    """Page dimensions and content."""
    page_num: int
    width: float
    height: float
    segments: list[TextSegment] = field(default_factory=list)
    
    def to_dict(self) -> dict:
        return {
            "page_num": self.page_num,
            "width": self.width,
            "height": self.height,
            "segments": [s.to_dict() for s in self.segments],
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> "PageInfo":
        page = cls(
            page_num=data["page_num"],
            width=data["width"],
            height=data["height"],
        )
        page.segments = [TextSegment.from_dict(s) for s in data.get("segments", [])]
        return page


@dataclass
class HTMLConfig(PipelineConfig):
    """Configuration for HTML intermediate pipeline."""
    
    pipeline_type: PipelineType = PipelineType.HTML_INTERMEDIATE
    
    # HTML output settings
    embed_fonts: bool = True
    use_web_fonts: bool = True  # Use Google Fonts for Unicode
    
    # Font fallbacks for different scripts
    font_fallbacks: dict[str, str] = field(default_factory=lambda: {
        "default": "Noto Sans, Arial, sans-serif",
        "hindi": "Noto Sans Devanagari, Mangal, sans-serif",
        "bengali": "Noto Sans Bengali, Vrinda, sans-serif",
        "arabic": "Noto Sans Arabic, sans-serif",
        "chinese": "Noto Sans SC, SimSun, sans-serif",
        "japanese": "Noto Sans JP, MS Gothic, sans-serif",
        "korean": "Noto Sans KR, Malgun Gothic, sans-serif",
    })
    
    # PDF rendering method
    render_method: str = "weasyprint"  # "weasyprint" or "browser"
    
    # Output settings
    keep_html: bool = True  # Keep intermediate HTML files


class HTMLIntermediatePipeline(TranslationPipeline):
    """
    HTML Intermediate pipeline for layout-preserving translation.
    
    This pipeline converts PDF to HTML with CSS absolute positioning,
    allows translation of text segments, then renders back to PDF.
    
    Advantages:
    - Excellent Unicode support (via web fonts)
    - Easy to preview/edit in browser
    - CSS handles complex layouts well
    - Can manually adjust positioning if needed
    
    Workflow:
    1. Extract text blocks with exact positions from PDF
    2. Generate HTML with absolute-positioned divs
    3. Generate translation file (Moses format)
    4. After translation, update HTML
    5. Render HTML to PDF
    
    Best for:
    - Documents requiring visual preview before final PDF
    - Unicode-heavy translations
    - When manual adjustment might be needed
    """
    
    def __init__(self, config: HTMLConfig):
        super().__init__(config)
        self.config: HTMLConfig = config
    
    @property
    def name(self) -> str:
        return "HTML Intermediate"
    
    @property
    def description(self) -> str:
        return "PDF → HTML (CSS positioned) → Translate → PDF"
    
    def extract(self, input_path: Path) -> ExtractResult:
        """
        Extract text from PDF and generate HTML with CSS positioning.
        """
        paths = self.derive_paths(input_path)
        
        print(f"  Opening PDF: {input_path.name}")
        start_time = time.time()
        
        doc = fitz.open(str(input_path))
        pages: list[PageInfo] = []
        all_segments: list[TextSegment] = []
        seg_count = 0
        
        try:
            for page_num in range(len(doc)):
                page = doc[page_num]
                page_rect = page.rect
                
                page_info = PageInfo(
                    page_num=page_num,
                    width=page_rect.width,
                    height=page_rect.height,
                )
                
                # Extract text blocks
                blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
                
                for block in blocks:
                    if block.get("type") != 0:  # Text blocks only
                        continue
                    
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            text = span.get("text", "").strip()
                            if not text:
                                continue
                            
                            bbox = span.get("bbox", (0, 0, 0, 0))
                            font = span.get("font", "Arial")
                            size = span.get("size", 12)
                            color = span.get("color", 0)
                            flags = span.get("flags", 0)
                            
                            # Convert color to hex
                            if isinstance(color, int):
                                hex_color = f"#{color:06x}"
                            else:
                                hex_color = "#000000"
                            
                            # Detect bold/italic from flags
                            is_bold = bool(flags & 2**4)  # Bold flag
                            is_italic = bool(flags & 2**1)  # Italic flag
                            
                            segment = TextSegment(
                                seg_id=f"seg_{seg_count}",
                                page_num=page_num,
                                text=text,
                                x=bbox[0],
                                y=bbox[1],
                                width=bbox[2] - bbox[0],
                                height=bbox[3] - bbox[1],
                                font_name=font,
                                font_size=size,
                                color=hex_color,
                                is_bold=is_bold,
                                is_italic=is_italic,
                            )
                            
                            page_info.segments.append(segment)
                            all_segments.append(segment)
                            seg_count += 1
                
                pages.append(page_info)
            
        finally:
            doc.close()
        
        print(f"  Extracted {seg_count} text segments from {len(pages)} pages")
        
        # Save layout JSON
        layout = {
            "source_file": str(input_path),
            "pipeline": "html_intermediate",
            "target_language": self.config.target_language,
            "pages": [p.to_dict() for p in pages],
            "segment_order": [s.seg_id for s in all_segments],
        }
        
        paths["layout"].write_text(
            json.dumps(layout, indent=2, ensure_ascii=False),
            encoding="utf-8"
        )
        
        # Generate HTML file
        html_path = input_path.parent / f"{input_path.name}_preview.html"
        html_content = self._generate_html(pages, all_segments)
        html_path.write_text(html_content, encoding="utf-8")
        print(f"  Generated HTML: {html_path.name}")
        
        # Generate Moses format files
        source_path = input_path.parent / f"{input_path.name}_source.txt"
        target_path = input_path.parent / f"{input_path.name}_target.txt"
        
        source_lines = []
        for seg in all_segments:
            text = seg.text.replace('\n', ' <br> ')
            source_lines.append(text)
        
        source_path.write_text('\n'.join(source_lines), encoding="utf-8")
        target_path.write_text('\n'.join(source_lines), encoding="utf-8")  # Copy for translation
        
        # Generate tagged format
        tagged_lines = []
        for i, seg in enumerate(all_segments):
            text = seg.text.replace('\n', '\\n')
            tagged_lines.append(f"<{i}>{text}</{i}>")
        
        paths["translate"].write_text('\n'.join(tagged_lines), encoding="utf-8")
        paths["translated"].write_text('\n'.join(tagged_lines), encoding="utf-8")
        
        elapsed = time.time() - start_time
        print(f"  Extraction complete ({elapsed:.2f}s)")
        
        return ExtractResult(
            layout_path=paths["layout"],
            translate_path=paths["translate"],
            translated_template_path=paths["translated"],
            extra_files={
                "html_preview": html_path,
                "source_txt": source_path,
                "target_txt": target_path,
            },
        )
    
    def _generate_html(
        self,
        pages: list[PageInfo],
        segments: list[TextSegment],
    ) -> str:
        """Generate HTML with CSS absolute positioning."""
        
        # Determine font family based on target language
        lang_lower = self.config.target_language.lower()
        if "hindi" in lang_lower or "devanagari" in lang_lower:
            font_family = self.config.font_fallbacks.get("hindi", self.config.font_fallbacks["default"])
        elif "bengali" in lang_lower or "bangla" in lang_lower:
            font_family = self.config.font_fallbacks.get("bengali", self.config.font_fallbacks["default"])
        elif "arabic" in lang_lower:
            font_family = self.config.font_fallbacks.get("arabic", self.config.font_fallbacks["default"])
        elif "chinese" in lang_lower:
            font_family = self.config.font_fallbacks.get("chinese", self.config.font_fallbacks["default"])
        elif "japanese" in lang_lower:
            font_family = self.config.font_fallbacks.get("japanese", self.config.font_fallbacks["default"])
        elif "korean" in lang_lower:
            font_family = self.config.font_fallbacks.get("korean", self.config.font_fallbacks["default"])
        else:
            font_family = self.config.font_fallbacks["default"]
        
        # Build HTML
        html_parts = [
            '<!DOCTYPE html>',
            '<html lang="en">',
            '<head>',
            '  <meta charset="UTF-8">',
            '  <meta name="viewport" content="width=device-width, initial-scale=1.0">',
            f'  <title>PDF Translation - {self.config.target_language}</title>',
        ]
        
        # Add Google Fonts for Unicode support
        if self.config.use_web_fonts:
            html_parts.extend([
                '  <link rel="preconnect" href="https://fonts.googleapis.com">',
                '  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>',
                '  <link href="https://fonts.googleapis.com/css2?family=Noto+Sans:wght@400;700&family=Noto+Sans+Devanagari:wght@400;700&family=Noto+Sans+Bengali:wght@400;700&family=Noto+Sans+Arabic:wght@400;700&display=swap" rel="stylesheet">',
            ])
        
        # CSS styles
        html_parts.extend([
            '  <style>',
            '    @page { margin: 0; }',
            '    * { box-sizing: border-box; margin: 0; padding: 0; }',
            '    body { background: #f0f0f0; }',
            '    .page-container { margin: 20px auto; box-shadow: 0 2px 10px rgba(0,0,0,0.2); }',
            '    .page { position: relative; background: white; overflow: hidden; }',
            '    .text-segment {',
            '      position: absolute;',
            '      white-space: pre-wrap;',
            '      overflow: visible;',
            f'      font-family: {font_family};',
            '    }',
            '    .text-segment:hover { background: rgba(255, 255, 0, 0.3); }',
            '    @media print {',
            '      body { background: white; }',
            '      .page-container { margin: 0; box-shadow: none; }',
            '      .page { page-break-after: always; }',
            '      .text-segment:hover { background: transparent; }',
            '    }',
            '  </style>',
            '</head>',
            '<body>',
        ])
        
        # Generate pages
        for page_info in pages:
            html_parts.extend([
                f'  <div class="page-container" data-page="{page_info.page_num}">',
                f'    <div class="page" style="width: {page_info.width}pt; height: {page_info.height}pt;">',
            ])
            
            for seg in page_info.segments:
                # Build inline styles
                styles = [
                    f'left: {seg.x}pt',
                    f'top: {seg.y}pt',
                    f'font-size: {seg.font_size}pt',
                    f'color: {seg.color}',
                ]
                if seg.is_bold:
                    styles.append('font-weight: bold')
                if seg.is_italic:
                    styles.append('font-style: italic')
                
                style_str = '; '.join(styles)
                escaped_text = html.escape(seg.text)
                
                html_parts.append(
                    f'      <div class="text-segment" id="{seg.seg_id}" '
                    f'data-seg-id="{seg.seg_id}" style="{style_str}">{escaped_text}</div>'
                )
            
            html_parts.extend([
                '    </div>',
                '  </div>',
            ])
        
        html_parts.extend([
            '</body>',
            '</html>',
        ])
        
        return '\n'.join(html_parts)
    
    def merge(
        self,
        input_path: Path,
        output_path: Path,
        translated_path: Path,
        layout_path: Path,
    ) -> MergeResult:
        """
        Apply translations and generate PDF from HTML.
        """
        print(f"  Loading layout...")
        layout = json.loads(layout_path.read_text(encoding="utf-8"))
        
        # Parse translations
        translations = self._parse_translations(layout_path, translated_path, layout)
        print(f"  Parsed {len(translations)} translations")
        
        start_time = time.time()
        
        # Reconstruct pages from layout
        pages = [PageInfo.from_dict(p) for p in layout["pages"]]
        
        # Apply translations to segments
        for page_info in pages:
            for seg in page_info.segments:
                if seg.seg_id in translations:
                    seg.text = translations[seg.seg_id]
        
        # Flatten segments for HTML generation
        all_segments = []
        for page_info in pages:
            all_segments.extend(page_info.segments)
        
        # Generate translated HTML
        html_content = self._generate_html(pages, all_segments)
        html_path = input_path.parent / f"{input_path.name}_translated.html"
        html_path.write_text(html_content, encoding="utf-8")
        print(f"  Generated translated HTML: {html_path.name}")
        
        # Convert HTML to PDF
        pdf_generated = False
        
        if self.config.render_method == "weasyprint":
            try:
                pdf_generated = self._render_with_weasyprint(html_path, output_path, pages)
            except ImportError:
                print("  WeasyPrint not installed, trying PyMuPDF rendering...")
        
        if not pdf_generated:
            # Fallback: use PyMuPDF to create PDF with positioned text
            pdf_generated = self._render_with_pymupdf(pages, translations, input_path, output_path)
        
        if not pdf_generated:
            print(f"  Warning: Could not generate PDF automatically.")
            print(f"  Open {html_path} in a browser and print to PDF.")
            # Create empty placeholder
            output_path = html_path
        
        elapsed = time.time() - start_time
        
        if output_path.exists() and output_path.suffix == '.pdf':
            output_size = output_path.stat().st_size / 1024
            print(f"  Output: {output_path.name} ({output_size:.1f} KB)")
        
        print(f"  Completed in {elapsed:.2f}s")
        
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
        """Parse translations from Moses or tagged format."""
        segment_order = layout.get("segment_order", [])
        
        # Check Moses format first
        base_name = layout_path.stem.replace('_layout', '')
        source_path = layout_path.parent / f"{base_name}_source.txt"
        target_path = layout_path.parent / f"{base_name}_target.txt"
        
        if target_path.exists() and source_path.exists():
            source_content = source_path.read_text(encoding="utf-8")
            target_content = target_path.read_text(encoding="utf-8")
            
            # Only use Moses if content differs (has been translated)
            if source_content != target_content:
                return self._parse_moses_translations(target_path, segment_order)
        
        # Fall back to tagged format
        return self._parse_tagged_translations(translated_path, segment_order)
    
    def _parse_moses_translations(
        self,
        target_path: Path,
        segment_order: list[str],
    ) -> dict[str, str]:
        """Parse Moses format translations."""
        target_lines = target_path.read_text(encoding="utf-8").split('\n')
        
        translations = {}
        for i, text in enumerate(target_lines):
            if i < len(segment_order):
                text = text.replace(' <br> ', '\n')
                translations[segment_order[i]] = text
        
        return translations
    
    def _parse_tagged_translations(
        self,
        translated_path: Path,
        segment_order: list[str],
    ) -> dict[str, str]:
        """Parse tagged format translations."""
        content = translated_path.read_text(encoding="utf-8")
        translations = {}
        
        pattern = r'<(\d+)>(.*?)</\1>'
        for match in re.finditer(pattern, content, re.DOTALL):
            idx = int(match.group(1))
            text = match.group(2).replace('\\n', '\n')
            
            if idx < len(segment_order):
                translations[segment_order[idx]] = text
        
        return translations
    
    def _render_with_weasyprint(
        self,
        html_path: Path,
        output_path: Path,
        pages: list[PageInfo],
    ) -> bool:
        """Render HTML to PDF using WeasyPrint."""
        try:
            from weasyprint import HTML, CSS
            
            print("  Rendering PDF with WeasyPrint...")
            
            # Create custom CSS for page sizes
            page_css_parts = []
            for i, page_info in enumerate(pages):
                if i == 0:
                    page_css_parts.append(
                        f'@page {{ size: {page_info.width}pt {page_info.height}pt; margin: 0; }}'
                    )
            
            css_str = '\n'.join(page_css_parts)
            
            html_doc = HTML(filename=str(html_path))
            css = CSS(string=css_str)
            html_doc.write_pdf(str(output_path), stylesheets=[css])
            
            return True
            
        except ImportError:
            return False
        except Exception as e:
            print(f"  WeasyPrint error: {e}")
            return False
    
    def _render_with_pymupdf(
        self,
        pages: list[PageInfo],
        translations: dict[str, str],
        input_path: Path,
        output_path: Path,
    ) -> bool:
        """Render PDF using PyMuPDF with positioned text."""
        try:
            print("  Rendering PDF with PyMuPDF...")
            
            # Create new PDF
            doc = fitz.open()
            
            # Load Unicode font
            font_path = self._find_unicode_font()
            unicode_font = None
            if font_path:
                try:
                    unicode_font = fitz.Font(fontfile=str(font_path))
                    print(f"    Using font: {font_path.name}")
                except Exception as e:
                    print(f"    Warning: Could not load font: {e}")
            
            for page_info in pages:
                # Create page with same dimensions
                page = doc.new_page(
                    width=page_info.width,
                    height=page_info.height,
                )
                
                for seg in page_info.segments:
                    text = translations.get(seg.seg_id, seg.text)
                    
                    # Parse color from hex
                    try:
                        color_hex = seg.color.lstrip('#')
                        r = int(color_hex[0:2], 16) / 255.0
                        g = int(color_hex[2:4], 16) / 255.0
                        b = int(color_hex[4:6], 16) / 255.0
                        text_color = (r, g, b)
                    except:
                        text_color = (0, 0, 0)
                    
                    # Use TextWriter for proper Unicode support
                    try:
                        tw = fitz.TextWriter(page.rect, color=text_color)
                        
                        # Insert text at position (baseline)
                        if unicode_font:
                            tw.append(
                                fitz.Point(seg.x, seg.y + seg.font_size),
                                text,
                                font=unicode_font,
                                fontsize=seg.font_size,
                            )
                        else:
                            tw.append(
                                fitz.Point(seg.x, seg.y + seg.font_size),
                                text,
                                font=fitz.Font("helv"),
                                fontsize=seg.font_size,
                            )
                        
                        tw.write_text(page)
                        
                    except Exception as e:
                        # Fallback to simple text insertion
                        page.insert_text(
                            (seg.x, seg.y + seg.font_size),
                            text,
                            fontsize=seg.font_size,
                            color=text_color,
                        )
            
            doc.save(str(output_path))
            doc.close()
            
            return True
            
        except Exception as e:
            print(f"  PyMuPDF rendering error: {e}")
            return False
    
    def _find_unicode_font(self) -> Optional[Path]:
        """Find a Unicode font for rendering."""
        # Common font paths (in order of preference)
        font_candidates = [
            # Windows - Hindi/Bengali/Devanagari
            Path("C:/Windows/Fonts/Nirmala.ttc"),
            Path("C:/Windows/Fonts/NirmalaS.ttc"),
            Path("C:/Windows/Fonts/mangal.ttf"),
            # Windows - General Unicode
            Path("C:/Windows/Fonts/arial.ttf"),
            Path("C:/Windows/Fonts/arialuni.ttf"),
            Path("C:/Windows/Fonts/NotoSans-Regular.ttf"),
            # Linux
            Path("/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf"),
            Path("/usr/share/fonts/opentype/noto/NotoSansBengali-Regular.ttf"),
            Path("/usr/share/fonts/truetype/freefont/FreeSans.ttf"),
            # macOS
            Path("/System/Library/Fonts/Supplemental/Arial Unicode.ttf"),
            Path("/Library/Fonts/Arial Unicode.ttf"),
        ]
        
        for font_path in font_candidates:
            if font_path.exists():
                return font_path
        
        return None
    
    def get_translation_prompt(self, block_count: int) -> str:
        """Get translation prompt for HTML pipeline."""
        return f"""
HTML Intermediate Translation - {block_count} text segments

Files:
- *_source.txt: Source text (one segment per line)
- *_target.txt: Translate this file
- *_preview.html: Visual preview (open in browser)

Instructions:
1. Open *_preview.html in browser to see layout
2. Edit *_target.txt - translate each line
3. Keep the SAME number of lines
4. Preserve <br> markers (line breaks)
5. Save as UTF-8 encoding

After translation, run: python main.py merge input.pdf

The merge will generate:
- *_translated.html (preview with translations)
- *_translated.pdf (final output)
"""


def create_html_pipeline(
    target_language: str = "Hindi",
    use_web_fonts: bool = True,
    render_method: str = "weasyprint",
) -> HTMLIntermediatePipeline:
    """
    Factory function to create HTML Intermediate pipeline.
    
    Args:
        target_language: Target language for translation
        use_web_fonts: Use Google Fonts for Unicode support
        render_method: "weasyprint" or "browser"
    """
    config = HTMLConfig(
        target_language=target_language,
        use_web_fonts=use_web_fonts,
        render_method=render_method,
    )
    return HTMLIntermediatePipeline(config)
