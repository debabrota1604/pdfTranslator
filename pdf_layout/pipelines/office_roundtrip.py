"""
Office Roundtrip Pipeline - PDF → Office Format → Translate → Office → PDF.

Converts PDF to the original Office format (DOCX/PPTX/XLSX), translates the
XML content, then converts back to PDF.

Supports:
- DOCX (Word documents)
- PPTX (PowerPoint presentations)
- XLSX (Excel spreadsheets)

Required dependencies (install separately):
- pdf2docx: pip install pdf2docx (for PDF → DOCX)
- python-pptx: pip install python-pptx (for PPTX handling)
- openpyxl: pip install openpyxl (for XLSX handling)
- LibreOffice: https://www.libreoffice.org/ (for Office → PDF conversion)
"""

from __future__ import annotations

import json
import re
import subprocess
import zipfile
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

from .base import (
    TranslationPipeline,
    PipelineConfig,
    PipelineType,
    ExtractResult,
    MergeResult,
)
from ..source_detector import detect_source_format, SourceFormat, SourceInfo


class OfficeFormat(Enum):
    """Office format types."""
    DOCX = "docx"
    PPTX = "pptx"
    XLSX = "xlsx"
    AUTO = "auto"  # Auto-detect from PDF metadata


# XML namespaces for different Office formats
WORD_NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

PPTX_NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

XLSX_NS = {
    'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}


@dataclass
class OfficeRoundtripConfig(PipelineConfig):
    """Configuration for Office roundtrip pipeline."""
    
    pipeline_type: PipelineType = PipelineType.DOCX_ROUNDTRIP
    
    # Office format (auto-detect from PDF if AUTO)
    office_format: OfficeFormat = OfficeFormat.AUTO
    
    # LibreOffice path
    libreoffice_path: Optional[Path] = None
    
    # Keep intermediate files
    keep_intermediate: bool = False


class OfficeRoundtripPipeline(TranslationPipeline):
    """
    Office Roundtrip Pipeline.
    
    Auto-detects the original Office format from PDF metadata and uses
    appropriate conversion tools.
    
    Workflow:
    1. Detect source format from PDF metadata
    2. Convert PDF to Office format (DOCX/PPTX/XLSX)
    3. Extract text from Office XML
    4. Generate translation file
    5. Parse translations and update XML
    6. Repack Office file
    7. Convert Office → PDF via LibreOffice
    """
    
    def __init__(self, config: OfficeRoundtripConfig):
        super().__init__(config)
        self.config: OfficeRoundtripConfig = config
        self._detected_format: Optional[SourceInfo] = None
        self._check_dependencies()
    
    def _check_dependencies(self) -> None:
        """Check available dependencies."""
        self._has_pdf2docx = False
        self._has_pptx = False
        self._has_openpyxl = False
        self._has_libreoffice = False
        
        try:
            import pdf2docx
            self._has_pdf2docx = True
        except ImportError:
            pass
        
        try:
            import pptx
            self._has_pptx = True
        except ImportError:
            pass
        
        try:
            import openpyxl
            self._has_openpyxl = True
        except ImportError:
            pass
        
        # Check LibreOffice
        lo_paths = [
            self.config.libreoffice_path,
            Path("C:/Program Files/LibreOffice/program/soffice.exe"),
            Path("/usr/bin/libreoffice"),
            Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
        ]
        for path in lo_paths:
            if path and path.exists():
                self.config.libreoffice_path = path
                self._has_libreoffice = True
                break
    
    @property
    def name(self) -> str:
        return "Office Roundtrip"
    
    @property
    def description(self) -> str:
        return "PDF → Office (DOCX/PPTX/XLSX) → Translate → Office → PDF"
    
    def derive_paths(self, input_path: Path) -> dict[str, Path]:
        """Derive paths including Office intermediates."""
        base = super().derive_paths(input_path)
        
        # Will be updated with actual extension after format detection
        base.update({
            "office": input_path.with_suffix(".office"),  # Placeholder
            "office_translated": input_path.parent / f"{input_path.stem}_translated.office",
        })
        return base
    
    def _get_office_extension(self) -> str:
        """Get Office file extension based on detected/configured format."""
        if self.config.office_format == OfficeFormat.PPTX:
            return ".pptx"
        elif self.config.office_format == OfficeFormat.XLSX:
            return ".xlsx"
        else:
            return ".docx"
    
    def extract(self, input_path: Path) -> ExtractResult:
        """Extract text from PDF via Office format conversion."""
        paths = self.derive_paths(input_path)
        
        # Step 1: Detect source format if AUTO
        if self.config.office_format == OfficeFormat.AUTO:
            self._detected_format = detect_source_format(input_path)
            print(f"  Detected source: {self._detected_format.format.value} "
                  f"({self._detected_format.confidence:.0%} confidence)")
            
            # Map source format to office format
            if self._detected_format.format == SourceFormat.POWERPOINT:
                self.config.office_format = OfficeFormat.PPTX
            elif self._detected_format.format == SourceFormat.EXCEL:
                self.config.office_format = OfficeFormat.XLSX
            else:
                self.config.office_format = OfficeFormat.DOCX
        
        # Update paths with correct extension
        ext = self._get_office_extension()
        paths["office"] = input_path.with_suffix(ext)
        paths["office_translated"] = input_path.parent / f"{input_path.stem}_translated{ext}"
        
        # Step 2: Convert PDF to Office format
        if self.config.office_format == OfficeFormat.PPTX:
            text_blocks = self._extract_pptx(input_path, paths["office"])
        elif self.config.office_format == OfficeFormat.XLSX:
            text_blocks = self._extract_xlsx(input_path, paths["office"])
        else:
            text_blocks = self._extract_docx(input_path, paths["office"])
        
        # Step 3: Generate layout JSON
        layout = {
            "source_file": str(input_path),
            "pipeline": "office_roundtrip",
            "office_format": self.config.office_format.value,
            "office_path": str(paths["office"]),
            "blocks": text_blocks,
            "block_order": [b["block_id"] for b in text_blocks],
        }
        paths["layout"].write_text(
            json.dumps(layout, indent=2, ensure_ascii=False),
            encoding="utf-8"
        )
        
        # Step 4: Generate translation file
        lines = []
        for i, block in enumerate(text_blocks):
            text = block["text"].replace('\n', '\\n')
            lines.append(f"<{i}>{text}</{i}>")
        
        paths["translate"].write_text('\n'.join(lines), encoding="utf-8")
        paths["translated"].write_text('\n'.join(lines), encoding="utf-8")
        
        return ExtractResult(
            layout_path=paths["layout"],
            translate_path=paths["translate"],
            translated_template_path=paths["translated"],
            extra_files={
                "office": paths["office"],
                "format": Path(self.config.office_format.value),  # Just for display
            },
        )
    
    def _extract_docx(self, pdf_path: Path, docx_path: Path) -> list[dict]:
        """Extract text via DOCX conversion."""
        if not self._has_pdf2docx:
            raise ImportError(
                "pdf2docx is required for DOCX conversion.\n"
                "Install with: pip install pdf2docx"
            )
        
        from pdf2docx import Converter
        
        cv = Converter(str(pdf_path))
        cv.convert(str(docx_path), start=0, end=None)
        cv.close()
        
        return self._extract_docx_text(docx_path)
    
    def _extract_docx_text(self, docx_path: Path) -> list[dict]:
        """Extract text blocks from DOCX XML."""
        text_blocks = []
        
        with zipfile.ZipFile(docx_path, 'r') as zf:
            with zf.open('word/document.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()
        
        for i, para in enumerate(root.findall('.//w:p', WORD_NS)):
            text_parts = []
            for text_elem in para.findall('.//w:t', WORD_NS):
                if text_elem.text:
                    text_parts.append(text_elem.text)
            
            full_text = ''.join(text_parts).strip()
            if full_text:
                text_blocks.append({
                    "block_id": f"docx_p{i}",
                    "text": full_text,
                    "type": "paragraph",
                })
        
        return text_blocks
    
    def _extract_pptx(self, pdf_path: Path, pptx_path: Path) -> list[dict]:
        """Extract text via PPTX (currently uses DOCX as fallback)."""
        # Note: Direct PDF → PPTX conversion is complex
        # For now, we extract text differently for PPTX-like layouts
        print("  Note: PPTX extraction uses PDF text extraction (no direct PDF→PPTX converter)")
        
        # Use PyMuPDF to extract text with presentation-style handling
        import fitz
        
        text_blocks = []
        doc = fitz.open(str(pdf_path))
        
        try:
            for page_num, page in enumerate(doc):
                blocks = page.get_text("dict")["blocks"]
                
                for block_num, block in enumerate(blocks):
                    if block.get("type") == 0:  # Text block
                        text_parts = []
                        for line in block.get("lines", []):
                            for span in line.get("spans", []):
                                if span.get("text", "").strip():
                                    text_parts.append(span["text"])
                        
                        full_text = ' '.join(text_parts).strip()
                        if full_text:
                            text_blocks.append({
                                "block_id": f"slide{page_num + 1}_b{block_num}",
                                "text": full_text,
                                "type": "slide_text",
                                "slide": page_num + 1,
                            })
        finally:
            doc.close()
        
        # Create a placeholder PPTX file
        self._create_placeholder_pptx(pptx_path, text_blocks)
        
        return text_blocks
    
    def _create_placeholder_pptx(self, pptx_path: Path, blocks: list[dict]) -> None:
        """Create a placeholder PPTX file structure."""
        if not self._has_pptx:
            # Create minimal ZIP structure
            with zipfile.ZipFile(pptx_path, 'w') as zf:
                zf.writestr('[Content_Types].xml', '<?xml version="1.0"?><Types/>')
            return
        
        from pptx import Presentation
        from pptx.util import Inches, Pt
        
        prs = Presentation()
        
        # Group blocks by slide
        slides_data = {}
        for block in blocks:
            slide_num = block.get("slide", 1)
            if slide_num not in slides_data:
                slides_data[slide_num] = []
            slides_data[slide_num].append(block["text"])
        
        # Create slides
        blank_layout = prs.slide_layouts[6]  # Blank layout
        for slide_num in sorted(slides_data.keys()):
            slide = prs.slides.add_slide(blank_layout)
            
            # Add text boxes for each text block
            y_pos = Inches(1)
            for text in slides_data[slide_num]:
                txBox = slide.shapes.add_textbox(Inches(0.5), y_pos, Inches(9), Inches(0.5))
                tf = txBox.text_frame
                tf.text = text
                y_pos += Inches(0.6)
        
        prs.save(str(pptx_path))
    
    def _extract_xlsx(self, pdf_path: Path, xlsx_path: Path) -> list[dict]:
        """Extract text via XLSX (currently uses PDF text extraction)."""
        print("  Note: XLSX extraction uses PDF text extraction (no direct PDF→XLSX converter)")
        
        import fitz
        
        text_blocks = []
        doc = fitz.open(str(pdf_path))
        
        try:
            for page_num, page in enumerate(doc):
                blocks = page.get_text("dict")["blocks"]
                
                for block_num, block in enumerate(blocks):
                    if block.get("type") == 0:
                        text_parts = []
                        for line in block.get("lines", []):
                            line_text = []
                            for span in line.get("spans", []):
                                if span.get("text", "").strip():
                                    line_text.append(span["text"])
                            if line_text:
                                text_parts.append(' '.join(line_text))
                        
                        full_text = '\n'.join(text_parts).strip()
                        if full_text:
                            text_blocks.append({
                                "block_id": f"sheet{page_num + 1}_cell{block_num}",
                                "text": full_text,
                                "type": "cell",
                                "sheet": page_num + 1,
                            })
        finally:
            doc.close()
        
        # Create placeholder XLSX
        self._create_placeholder_xlsx(xlsx_path, text_blocks)
        
        return text_blocks
    
    def _create_placeholder_xlsx(self, xlsx_path: Path, blocks: list[dict]) -> None:
        """Create placeholder XLSX file."""
        if not self._has_openpyxl:
            with zipfile.ZipFile(xlsx_path, 'w') as zf:
                zf.writestr('[Content_Types].xml', '<?xml version="1.0"?><Types/>')
            return
        
        from openpyxl import Workbook
        
        wb = Workbook()
        ws = wb.active
        
        for i, block in enumerate(blocks):
            ws.cell(row=i + 1, column=1, value=block["text"])
        
        wb.save(str(xlsx_path))
    
    def merge(
        self,
        input_path: Path,
        output_path: Path,
        translated_path: Path,
        layout_path: Path,
    ) -> MergeResult:
        """Apply translations and convert back to PDF."""
        if not self._has_libreoffice:
            raise RuntimeError(
                "LibreOffice is required for Office → PDF conversion.\n"
                "Install from: https://www.libreoffice.org/"
            )
        
        # Load layout
        layout = json.loads(layout_path.read_text(encoding="utf-8"))
        office_format = layout.get("office_format", "docx")
        office_path = Path(layout.get("office_path", ""))
        block_order = layout.get("block_order", [])
        
        # Parse translations
        translations = self._parse_translations(translated_path, block_order)
        
        # Create translated Office file
        ext = f".{office_format}"
        translated_office = input_path.parent / f"{input_path.stem}_translated{ext}"
        
        if office_format == "docx":
            self._update_docx(office_path, translated_office, translations)
        elif office_format == "pptx":
            self._update_pptx(office_path, translated_office, translations, layout)
        elif office_format == "xlsx":
            self._update_xlsx(office_path, translated_office, translations, layout)
        
        # Convert to PDF
        self._office_to_pdf(translated_office, output_path)
        
        # Cleanup
        if not self.config.keep_intermediate:
            office_path.unlink(missing_ok=True)
            translated_office.unlink(missing_ok=True)
        
        return MergeResult(
            output_path=output_path,
            blocks_processed=len(translations),
        )
    
    def _parse_translations(
        self,
        translated_path: Path,
        block_order: list[str],
    ) -> dict[str, str]:
        """Parse translated file."""
        content = translated_path.read_text(encoding="utf-8")
        translations = {}
        
        pattern = r'<(\d+)>(.*?)</\1>'
        for match in re.finditer(pattern, content, re.DOTALL):
            idx = int(match.group(1))
            text = match.group(2).replace('\\n', '\n')
            
            if idx < len(block_order):
                translations[block_order[idx]] = text
        
        return translations
    
    def _update_docx(
        self,
        source: Path,
        output: Path,
        translations: dict[str, str],
    ) -> None:
        """Update DOCX with translations."""
        import shutil
        shutil.copy(source, output)
        
        # Read and modify XML
        with zipfile.ZipFile(output, 'r') as zf:
            with zf.open('word/document.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()
        
        for i, para in enumerate(root.findall('.//w:p', WORD_NS)):
            block_id = f"docx_p{i}"
            if block_id in translations:
                text_elems = para.findall('.//w:t', WORD_NS)
                if text_elems:
                    text_elems[0].text = translations[block_id]
                    for elem in text_elems[1:]:
                        elem.text = ""
        
        for prefix, uri in WORD_NS.items():
            ET.register_namespace(prefix, uri)
        
        with zipfile.ZipFile(output, 'a') as zf:
            zf.writestr('word/document.xml', ET.tostring(root, encoding='unicode'))
    
    def _update_pptx(
        self,
        source: Path,
        output: Path,
        translations: dict[str, str],
        layout: dict,
    ) -> None:
        """Update PPTX with translations."""
        if not self._has_pptx:
            raise ImportError("python-pptx required: pip install python-pptx")
        
        from pptx import Presentation
        
        prs = Presentation(str(source))
        
        # Build slide/block index
        block_idx = 0
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    block_id = layout.get("blocks", [{}])[block_idx].get("block_id", "")
                    if block_id in translations:
                        for para in shape.text_frame.paragraphs:
                            if para.runs:
                                para.runs[0].text = translations[block_id]
                                for run in para.runs[1:]:
                                    run.text = ""
                    block_idx += 1
        
        prs.save(str(output))
    
    def _update_xlsx(
        self,
        source: Path,
        output: Path,
        translations: dict[str, str],
        layout: dict,
    ) -> None:
        """Update XLSX with translations."""
        if not self._has_openpyxl:
            raise ImportError("openpyxl required: pip install openpyxl")
        
        from openpyxl import load_workbook
        
        wb = load_workbook(str(source))
        ws = wb.active
        
        for i, block in enumerate(layout.get("blocks", [])):
            block_id = block.get("block_id", "")
            if block_id in translations:
                ws.cell(row=i + 1, column=1, value=translations[block_id])
        
        wb.save(str(output))
    
    def _office_to_pdf(self, office_path: Path, pdf_path: Path) -> None:
        """Convert Office file to PDF via LibreOffice."""
        cmd = [
            str(self.config.libreoffice_path),
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(pdf_path.parent),
            str(office_path),
        ]
        
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")
        
        generated = office_path.with_suffix(".pdf")
        if generated != pdf_path and generated.exists():
            generated.rename(pdf_path)
    
    def get_translation_prompt(self, block_count: int) -> str:
        """Get translation prompt."""
        from ..translation_io import get_translation_prompt
        return get_translation_prompt(block_count, self.config.target_language)


def create_office_roundtrip_pipeline(
    target_language: str = "Hindi",
    office_format: OfficeFormat = OfficeFormat.AUTO,
    libreoffice_path: Optional[Path] = None,
    keep_intermediate: bool = False,
) -> OfficeRoundtripPipeline:
    """Factory function to create Office roundtrip pipeline."""
    config = OfficeRoundtripConfig(
        target_language=target_language,
        office_format=office_format,
        libreoffice_path=libreoffice_path,
        keep_intermediate=keep_intermediate,
    )
    return OfficeRoundtripPipeline(config)
