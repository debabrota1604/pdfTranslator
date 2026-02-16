"""
XLIFF Pipeline - Generate industry-standard XLIFF format.

XLIFF (XML Localization Interchange File Format) is an XML-based format
for exchanging localization data between tools.

This pipeline generates XLIFF 1.2 or 2.0 format, compatible with:
- SDL Trados
- memoQ
- Memsource
- OmegaT
- Other CAT (Computer-Assisted Translation) tools
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET
from xml.dom import minidom

from .base import (
    TranslationPipeline,
    PipelineConfig,
    PipelineType,
    ExtractResult,
    MergeResult,
)
from ..extractor import extract_pdf_layout
from ..rebuilder_unicode import PDFRebuilder, RebuildConfig, RenderMethod


class XLIFFVersion:
    """XLIFF version constants."""
    V1_2 = "1.2"
    V2_0 = "2.0"


# XLIFF namespaces
XLIFF_NS_1_2 = "urn:oasis:names:tc:xliff:document:1.2"
XLIFF_NS_2_0 = "urn:oasis:names:tc:xliff:document:2.0"


@dataclass
class XLIFFConfig(PipelineConfig):
    """Configuration specific to XLIFF pipeline."""
    
    pipeline_type: PipelineType = PipelineType.XLIFF
    
    # XLIFF settings
    xliff_version: str = XLIFFVersion.V1_2
    source_language: str = "en"
    
    # Render settings for merge (same as direct PDF)
    render_method: RenderMethod = RenderMethod.LINE_BY_LINE


class XLIFFPipeline(TranslationPipeline):
    """
    XLIFF Format Pipeline.
    
    Generates industry-standard XLIFF files for professional translation workflows.
    
    Workflow:
    1. Extract text blocks from PDF
    2. Generate XLIFF file with source text
    3. Translator/CAT tool fills in <target> elements
    4. Parse translated XLIFF
    5. Rebuild PDF with translations
    
    Pros:
    - Industry standard, works with all CAT tools
    - Supports translation memory
    - Professional workflow integration
    
    Cons:
    - More complex file format
    - Still requires PDF rebuild step
    """
    
    def __init__(self, config: XLIFFConfig):
        super().__init__(config)
        self.config: XLIFFConfig = config
    
    @property
    def name(self) -> str:
        return "XLIFF Format"
    
    @property
    def description(self) -> str:
        return "Generate XLIFF format for professional CAT tools"
    
    def derive_paths(self, input_path: Path) -> dict[str, Path]:
        """Derive paths including XLIFF files."""
        base = super().derive_paths(input_path)
        base.update({
            "xliff": input_path.with_suffix(".xlf"),
            "xliff_translated": input_path.parent / f"{input_path.stem}_translated.xlf",
        })
        return base
    
    def extract(self, input_path: Path) -> ExtractResult:
        """Extract text and generate XLIFF file."""
        paths = self.derive_paths(input_path)
        
        # Extract PDF layout
        document = extract_pdf_layout(input_path, paths["layout"])
        
        # Count blocks
        total_blocks = sum(len(page.blocks) for page in document.pages)
        
        # Generate XLIFF
        self._generate_xliff(paths["layout"], paths["xliff"])
        
        # Also generate simple text format for quick LLM translation
        self._generate_simple_format(paths["layout"], paths["translate"])
        
        # Create template (copy of translate file)
        paths["translated"].write_text(
            paths["translate"].read_text(encoding="utf-8"),
            encoding="utf-8"
        )
        
        return ExtractResult(
            layout_path=paths["layout"],
            translate_path=paths["translate"],
            translated_template_path=paths["translated"],
            extra_files={"xliff": paths["xliff"]},
        )
    
    def _generate_xliff(self, layout_path: Path, xliff_path: Path) -> None:
        """Generate XLIFF file from layout."""
        layout = json.loads(layout_path.read_text(encoding="utf-8"))
        
        if self.config.xliff_version == XLIFFVersion.V1_2:
            xliff_content = self._generate_xliff_1_2(layout)
        else:
            xliff_content = self._generate_xliff_2_0(layout)
        
        xliff_path.write_text(xliff_content, encoding="utf-8")
    
    def _generate_xliff_1_2(self, layout: dict) -> str:
        """Generate XLIFF 1.2 format."""
        root = ET.Element("xliff")
        root.set("version", "1.2")
        root.set("xmlns", XLIFF_NS_1_2)
        
        file_elem = ET.SubElement(root, "file")
        file_elem.set("original", layout.get("source_file", "unknown"))
        file_elem.set("source-language", self.config.source_language)
        file_elem.set("target-language", self.config.target_language.lower()[:2])
        file_elem.set("datatype", "plaintext")
        
        body = ET.SubElement(file_elem, "body")
        
        for page in layout.get("pages", []):
            for block in page.get("blocks", []):
                trans_unit = ET.SubElement(body, "trans-unit")
                trans_unit.set("id", block["block_id"])
                
                source = ET.SubElement(trans_unit, "source")
                source.text = block["text"]
                
                # Empty target for translator to fill
                target = ET.SubElement(trans_unit, "target")
                target.set("state", "new")
                
                # Add note with context
                note = ET.SubElement(trans_unit, "note")
                note.text = f"Page {page['page_number']}, Font: {block.get('font_name', 'unknown')}"
        
        return self._prettify_xml(root)
    
    def _generate_xliff_2_0(self, layout: dict) -> str:
        """Generate XLIFF 2.0 format."""
        root = ET.Element("xliff")
        root.set("version", "2.0")
        root.set("xmlns", XLIFF_NS_2_0)
        root.set("srcLang", self.config.source_language)
        root.set("trgLang", self.config.target_language.lower()[:2])
        
        file_elem = ET.SubElement(root, "file")
        file_elem.set("id", "f1")
        file_elem.set("original", layout.get("source_file", "unknown"))
        
        for page in layout.get("pages", []):
            for block in page.get("blocks", []):
                unit = ET.SubElement(file_elem, "unit")
                unit.set("id", block["block_id"])
                
                segment = ET.SubElement(unit, "segment")
                
                source = ET.SubElement(segment, "source")
                source.text = block["text"]
                
                target = ET.SubElement(segment, "target")
                # Empty for translator
        
        return self._prettify_xml(root)
    
    def _generate_simple_format(self, layout_path: Path, output_path: Path) -> None:
        """Generate simple tagged format (same as DirectPDF)."""
        layout = json.loads(layout_path.read_text(encoding="utf-8"))
        
        lines = []
        idx = 0
        for page in layout.get("pages", []):
            for block in page.get("blocks", []):
                text = block["text"].replace('\n', '\\n')
                lines.append(f"<{idx}>{text}</{idx}>")
                idx += 1
        
        output_path.write_text('\n'.join(lines), encoding="utf-8")
    
    def _prettify_xml(self, elem: ET.Element) -> str:
        """Return pretty-printed XML string."""
        rough_string = ET.tostring(elem, encoding='unicode')
        reparsed = minidom.parseString(rough_string)
        return reparsed.toprettyxml(indent="  ")
    
    def merge(
        self,
        input_path: Path,
        output_path: Path,
        translated_path: Path,
        layout_path: Path,
    ) -> MergeResult:
        """Parse XLIFF or simple format and rebuild PDF."""
        paths = self.derive_paths(input_path)
        
        # Detect format and parse translations
        if translated_path.suffix.lower() in ['.xlf', '.xliff']:
            translations = self._parse_xliff(translated_path)
        else:
            translations = self._parse_simple_format(translated_path, layout_path)
        
        # Rebuild PDF using same method as DirectPDF
        rebuild_config = RebuildConfig(
            render_method=self.config.render_method,
            min_font_size=self.config.min_font_size,
            font_step=self.config.font_step,
        )
        
        rebuilder = PDFRebuilder(rebuild_config)
        rebuilder.rebuild(
            pdf_path=input_path,
            layout_data=layout_path,
            translations=translations,
            output_path=output_path,
        )
        
        return MergeResult(
            output_path=output_path,
            blocks_processed=len(translations),
        )
    
    def _parse_xliff(self, xliff_path: Path) -> dict[str, str]:
        """Parse translated XLIFF file."""
        translations = {}
        
        tree = ET.parse(xliff_path)
        root = tree.getroot()
        
        # Handle both XLIFF 1.2 and 2.0
        # XLIFF 1.2: trans-unit/target
        for trans_unit in root.findall('.//{urn:oasis:names:tc:xliff:document:1.2}trans-unit'):
            unit_id = trans_unit.get('id')
            target = trans_unit.find('{urn:oasis:names:tc:xliff:document:1.2}target')
            if target is not None and target.text:
                translations[unit_id] = target.text
        
        # XLIFF 2.0: unit/segment/target
        for unit in root.findall('.//{urn:oasis:names:tc:xliff:document:2.0}unit'):
            unit_id = unit.get('id')
            target = unit.find('.//{urn:oasis:names:tc:xliff:document:2.0}target')
            if target is not None and target.text:
                translations[unit_id] = target.text
        
        return translations
    
    def _parse_simple_format(
        self,
        translated_path: Path,
        layout_path: Path,
    ) -> dict[str, str]:
        """Parse simple tagged format."""
        import re
        
        layout = json.loads(layout_path.read_text(encoding="utf-8"))
        block_order = layout.get("block_order", [])
        
        # Build block_order if not present
        if not block_order:
            for page in layout.get("pages", []):
                for block in page.get("blocks", []):
                    block_order.append(block["block_id"])
        
        content = translated_path.read_text(encoding="utf-8")
        translations = {}
        
        pattern = r'<(\d+)>(.*?)</\1>'
        for match in re.finditer(pattern, content, re.DOTALL):
            idx = int(match.group(1))
            text = match.group(2).replace('\\n', '\n')
            
            if idx < len(block_order):
                translations[block_order[idx]] = text
        
        return translations
    
    def get_translation_prompt(self, block_count: int) -> str:
        """Get LLM prompt for translation."""
        from ..translation_io import get_translation_prompt
        return get_translation_prompt(block_count, self.config.target_language)


def create_xliff_pipeline(
    target_language: str = "Hindi",
    source_language: str = "en",
    xliff_version: str = XLIFFVersion.V1_2,
) -> XLIFFPipeline:
    """Factory function to create XLIFF pipeline."""
    config = XLIFFConfig(
        target_language=target_language,
        source_language=source_language,
        xliff_version=xliff_version,
    )
    return XLIFFPipeline(config)
