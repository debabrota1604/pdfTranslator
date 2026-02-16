"""
DOCX Roundtrip Pipeline - PDF → DOCX → Translate → DOCX → PDF.

Converts PDF to DOCX, translates the XML content, then converts back to PDF.
This approach leverages Word's text flow and layout engine.

Required dependencies (install separately):
- pdf2docx: pip install pdf2docx
- python-docx: pip install python-docx
- LibreOffice (for DOCX → PDF conversion): https://www.libreoffice.org/

Status: STUB - Not yet implemented
"""

from __future__ import annotations

import json
import subprocess
import zipfile
from dataclasses import dataclass
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


# DOCX XML namespaces
WORD_NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}


@dataclass
class DocxRoundtripConfig(PipelineConfig):
    """Configuration specific to DOCX roundtrip pipeline."""
    
    pipeline_type: PipelineType = PipelineType.DOCX_ROUNDTRIP
    
    # LibreOffice path (for DOCX → PDF conversion)
    libreoffice_path: Optional[Path] = None
    
    # Whether to keep intermediate DOCX files
    keep_intermediate: bool = True


class DocxRoundtripPipeline(TranslationPipeline):
    """
    DOCX Roundtrip Pipeline.
    
    Workflow:
    1. Convert PDF to DOCX using pdf2docx
    2. Extract XML from DOCX (it's a ZIP archive)
    3. Parse and extract translatable text from document.xml
    4. Generate translation file
    5. Parse translations and update XML
    6. Repack DOCX
    7. Convert DOCX to PDF using LibreOffice
    
    Pros:
    - Better text reflow (Word handles layout)
    - Preserves more formatting
    - Industry-standard intermediate format
    
    Cons:
    - Requires LibreOffice for final PDF conversion
    - PDF → DOCX conversion is lossy for complex layouts
    - Slower than direct PDF manipulation
    """
    
    def __init__(self, config: DocxRoundtripConfig):
        super().__init__(config)
        self.config: DocxRoundtripConfig = config
        self._check_dependencies()
    
    def _check_dependencies(self) -> None:
        """Check if required dependencies are available."""
        self._has_pdf2docx = False
        self._has_libreoffice = False
        
        try:
            import pdf2docx
            self._has_pdf2docx = True
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
        return "DOCX Roundtrip"
    
    @property
    def description(self) -> str:
        return "PDF → DOCX → Translate XML → DOCX → PDF (requires LibreOffice)"
    
    def derive_paths(self, input_path: Path) -> dict[str, Path]:
        """Derive paths including DOCX intermediates."""
        base = super().derive_paths(input_path)
        base.update({
            "docx": input_path.with_suffix(".docx"),
            "docx_translated": input_path.parent / f"{input_path.stem}_translated.docx",
        })
        return base
    
    def extract(self, input_path: Path) -> ExtractResult:
        """Convert PDF to DOCX and extract translatable text."""
        if not self._has_pdf2docx:
            raise ImportError(
                "pdf2docx is required for DOCX roundtrip pipeline.\n"
                "Install with: pip install pdf2docx"
            )
        
        paths = self.derive_paths(input_path)
        
        # Step 1: Convert PDF to DOCX
        from pdf2docx import Converter
        
        cv = Converter(str(input_path))
        cv.convert(str(paths["docx"]), start=0, end=None)
        cv.close()
        
        # Step 2: Extract text from DOCX XML
        text_blocks = self._extract_docx_text(paths["docx"])
        
        # Step 3: Generate layout JSON (compatible format)
        layout = {
            "source_file": str(input_path),
            "pipeline": "docx_roundtrip",
            "docx_path": str(paths["docx"]),
            "blocks": text_blocks,
            "block_order": [b["block_id"] for b in text_blocks],
        }
        paths["layout"].write_text(
            json.dumps(layout, indent=2, ensure_ascii=False),
            encoding="utf-8"
        )
        
        # Step 4: Generate translation file (same format as direct PDF)
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
            extra_files={"docx": paths["docx"]},
        )
    
    def _extract_docx_text(self, docx_path: Path) -> list[dict]:
        """Extract text blocks from DOCX XML."""
        text_blocks = []
        
        with zipfile.ZipFile(docx_path, 'r') as zf:
            with zf.open('word/document.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()
        
        # Find all paragraph elements
        for i, para in enumerate(root.findall('.//w:p', WORD_NS)):
            text_parts = []
            for text_elem in para.findall('.//w:t', WORD_NS):
                if text_elem.text:
                    text_parts.append(text_elem.text)
            
            full_text = ''.join(text_parts).strip()
            if full_text:
                text_blocks.append({
                    "block_id": f"p{i}",
                    "text": full_text,
                    "xpath": f".//w:p[{i+1}]",  # For later replacement
                })
        
        return text_blocks
    
    def merge(
        self,
        input_path: Path,
        output_path: Path,
        translated_path: Path,
        layout_path: Path,
    ) -> MergeResult:
        """Apply translations to DOCX and convert to PDF."""
        if not self._has_libreoffice:
            raise RuntimeError(
                "LibreOffice is required for DOCX → PDF conversion.\n"
                "Install from: https://www.libreoffice.org/"
            )
        
        paths = self.derive_paths(input_path)
        
        # Load layout
        layout = json.loads(layout_path.read_text(encoding="utf-8"))
        block_order = layout.get("block_order", [])
        
        # Parse translations
        translations = self._parse_translations(translated_path, block_order)
        
        # Update DOCX with translations
        self._update_docx(paths["docx"], paths["docx_translated"], translations)
        
        # Convert DOCX to PDF using LibreOffice
        self._docx_to_pdf(paths["docx_translated"], output_path)
        
        # Keep intermediate files
        print(f"  Kept: {paths['docx'].name}")
        print(f"  Kept: {paths['docx_translated'].name}")
        
        return MergeResult(
            output_path=output_path,
            blocks_processed=len(translations),
        )
    
    def _parse_translations(
        self,
        translated_path: Path,
        block_order: list[str],
    ) -> dict[str, str]:
        """Parse translated file into block_id -> text mapping."""
        import re
        
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
        source_docx: Path,
        output_docx: Path,
        translations: dict[str, str],
    ) -> None:
        """Update DOCX XML with translations."""
        import shutil
        
        # Copy original DOCX
        shutil.copy(source_docx, output_docx)
        
        # Modify the XML inside
        with zipfile.ZipFile(output_docx, 'r') as zf_in:
            with zf_in.open('word/document.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()
        
        # Update paragraphs with translations
        for i, para in enumerate(root.findall('.//w:p', WORD_NS)):
            block_id = f"p{i}"
            if block_id in translations:
                # Find all text elements and update
                text_elems = para.findall('.//w:t', WORD_NS)
                if text_elems:
                    # Put all text in first element, clear others
                    text_elems[0].text = translations[block_id]
                    for elem in text_elems[1:]:
                        elem.text = ""
        
        # Write back to DOCX
        # Register namespaces to avoid ns0 prefixes
        for prefix, uri in WORD_NS.items():
            ET.register_namespace(prefix, uri)
        
        with zipfile.ZipFile(output_docx, 'a') as zf:
            zf.writestr('word/document.xml', ET.tostring(root, encoding='unicode'))
    
    def _docx_to_pdf(self, docx_path: Path, pdf_path: Path) -> None:
        """Convert DOCX to PDF using LibreOffice."""
        cmd = [
            str(self.config.libreoffice_path),
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(pdf_path.parent),
            str(docx_path),
        ]
        
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")
        
        # LibreOffice creates file with same name but .pdf extension
        generated_pdf = docx_path.with_suffix(".pdf")
        if generated_pdf != pdf_path:
            generated_pdf.rename(pdf_path)
    
    def get_translation_prompt(self, block_count: int) -> str:
        """Get LLM prompt for translation."""
        from ..translation_io import get_translation_prompt
        return get_translation_prompt(block_count, self.config.target_language)


def create_docx_roundtrip_pipeline(
    target_language: str = "Hindi",
    libreoffice_path: Optional[Path] = None,
    keep_intermediate: bool = True,
) -> DocxRoundtripPipeline:
    """Factory function to create DOCX roundtrip pipeline."""
    config = DocxRoundtripConfig(
        target_language=target_language,
        libreoffice_path=libreoffice_path,
        keep_intermediate=keep_intermediate,
    )
    return DocxRoundtripPipeline(config)
