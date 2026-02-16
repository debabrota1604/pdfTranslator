"""
Office XML Handler - Extract and update text in Office document XML.

Office documents (DOCX, PPTX, XLSX) are ZIP archives containing XML files.
This module provides utilities to:
1. Extract text from the XML structure
2. Replace text with translations while preserving formatting
3. Repack the modified XML back into the Office archive
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, Iterator
from xml.etree import ElementTree as ET
import copy


# XML namespaces for Office formats
WORD_NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
}

PPTX_NAMESPACES = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

XLSX_NAMESPACES = {
    'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}


@dataclass
class TextSegment:
    """A segment of text from an Office document."""
    block_id: str
    text: str
    xml_path: str  # Path within the ZIP archive
    xpath: str  # XPath to locate the element
    element_index: int  # Index if multiple elements match
    metadata: dict = field(default_factory=dict)


@dataclass
class ExtractionResult:
    """Result of text extraction from Office document."""
    segments: list[TextSegment]
    xml_files: dict[str, bytes]  # Path -> original XML content
    other_files: dict[str, bytes]  # Non-XML files to preserve


class OfficeXMLHandler:
    """Base class for Office XML handling."""
    
    def __init__(self, office_path: Path):
        self.office_path = office_path
        self._namespaces: dict[str, str] = {}
        self._register_namespaces()
    
    def _register_namespaces(self) -> None:
        """Register XML namespaces to preserve them in output."""
        for prefix, uri in self._namespaces.items():
            ET.register_namespace(prefix, uri)
    
    def extract(self) -> ExtractionResult:
        """Extract text segments from Office document."""
        raise NotImplementedError
    
    def update(
        self,
        output_path: Path,
        translations: dict[str, str],
        extraction: ExtractionResult,
    ) -> None:
        """Update Office document with translations."""
        raise NotImplementedError
    
    def _read_archive(self) -> tuple[dict[str, bytes], dict[str, bytes]]:
        """Read all files from ZIP archive."""
        xml_files = {}
        other_files = {}
        
        with zipfile.ZipFile(self.office_path, 'r') as zf:
            for name in zf.namelist():
                content = zf.read(name)
                if name.endswith('.xml') or name.endswith('.rels'):
                    xml_files[name] = content
                else:
                    other_files[name] = content
        
        return xml_files, other_files
    
    def _write_archive(
        self,
        output_path: Path,
        xml_files: dict[str, bytes],
        other_files: dict[str, bytes],
    ) -> None:
        """Write files to ZIP archive."""
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for name, content in xml_files.items():
                zf.writestr(name, content)
            for name, content in other_files.items():
                zf.writestr(name, content)


class DocxXMLHandler(OfficeXMLHandler):
    """Handler for DOCX (Word) documents."""
    
    def __init__(self, docx_path: Path):
        self._namespaces = WORD_NAMESPACES
        super().__init__(docx_path)
    
    def extract(self) -> ExtractionResult:
        """Extract text from DOCX document."""
        xml_files, other_files = self._read_archive()
        segments = []
        
        # Main document content
        if 'word/document.xml' in xml_files:
            segments.extend(
                self._extract_from_document(xml_files['word/document.xml'])
            )
        
        # Headers
        for name in xml_files:
            if name.startswith('word/header') and name.endswith('.xml'):
                segments.extend(
                    self._extract_from_header_footer(xml_files[name], name, 'header')
                )
        
        # Footers
        for name in xml_files:
            if name.startswith('word/footer') and name.endswith('.xml'):
                segments.extend(
                    self._extract_from_header_footer(xml_files[name], name, 'footer')
                )
        
        return ExtractionResult(
            segments=segments,
            xml_files=xml_files,
            other_files=other_files,
        )
    
    def _extract_from_document(self, content: bytes) -> list[TextSegment]:
        """Extract text segments from document.xml."""
        segments = []
        root = ET.fromstring(content)
        
        # Find all paragraphs
        para_idx = 0
        for para in root.findall('.//w:p', WORD_NAMESPACES):
            text = self._get_paragraph_text(para)
            if text.strip():
                segments.append(TextSegment(
                    block_id=f"p{para_idx}",
                    text=text,
                    xml_path='word/document.xml',
                    xpath=f'.//w:p[{para_idx + 1}]',
                    element_index=para_idx,
                    metadata={'type': 'paragraph'},
                ))
            para_idx += 1
        
        # Find all tables
        table_idx = 0
        for table in root.findall('.//w:tbl', WORD_NAMESPACES):
            row_idx = 0
            for row in table.findall('.//w:tr', WORD_NAMESPACES):
                cell_idx = 0
                for cell in row.findall('.//w:tc', WORD_NAMESPACES):
                    text = self._get_cell_text(cell)
                    if text.strip():
                        segments.append(TextSegment(
                            block_id=f"t{table_idx}_r{row_idx}_c{cell_idx}",
                            text=text,
                            xml_path='word/document.xml',
                            xpath=f'.//w:tbl[{table_idx + 1}]//w:tr[{row_idx + 1}]//w:tc[{cell_idx + 1}]',
                            element_index=table_idx * 1000 + row_idx * 100 + cell_idx,
                            metadata={'type': 'table_cell', 'table': table_idx, 'row': row_idx, 'col': cell_idx},
                        ))
                    cell_idx += 1
                row_idx += 1
            table_idx += 1
        
        return segments
    
    def _extract_from_header_footer(
        self,
        content: bytes,
        xml_path: str,
        hf_type: str,
    ) -> list[TextSegment]:
        """Extract text from header/footer XML."""
        segments = []
        root = ET.fromstring(content)
        
        # Extract number from filename (e.g., header1.xml -> 1)
        num = re.search(r'(\d+)', xml_path)
        prefix = f"{hf_type}{num.group(1) if num else ''}"
        
        para_idx = 0
        for para in root.findall('.//w:p', WORD_NAMESPACES):
            text = self._get_paragraph_text(para)
            if text.strip():
                segments.append(TextSegment(
                    block_id=f"{prefix}_p{para_idx}",
                    text=text,
                    xml_path=xml_path,
                    xpath=f'.//w:p[{para_idx + 1}]',
                    element_index=para_idx,
                    metadata={'type': hf_type},
                ))
            para_idx += 1
        
        return segments
    
    def _get_paragraph_text(self, para: ET.Element) -> str:
        """Get full text from a paragraph element."""
        texts = []
        for t in para.findall('.//w:t', WORD_NAMESPACES):
            if t.text:
                texts.append(t.text)
        return ''.join(texts)
    
    def _get_cell_text(self, cell: ET.Element) -> str:
        """Get full text from a table cell."""
        texts = []
        for para in cell.findall('.//w:p', WORD_NAMESPACES):
            para_text = self._get_paragraph_text(para)
            if para_text:
                texts.append(para_text)
        return '\n'.join(texts)
    
    def update(
        self,
        output_path: Path,
        translations: dict[str, str],
        extraction: ExtractionResult,
    ) -> None:
        """Update DOCX with translations."""
        # Group segments by XML file
        segments_by_file: dict[str, list[TextSegment]] = {}
        for seg in extraction.segments:
            if seg.xml_path not in segments_by_file:
                segments_by_file[seg.xml_path] = []
            segments_by_file[seg.xml_path].append(seg)
        
        # Update each XML file
        updated_xml = dict(extraction.xml_files)
        
        for xml_path, segments in segments_by_file.items():
            if xml_path in updated_xml:
                content = updated_xml[xml_path]
                updated_content = self._update_xml_content(
                    content, segments, translations
                )
                updated_xml[xml_path] = updated_content
        
        # Write output
        self._write_archive(output_path, updated_xml, extraction.other_files)
    
    def _update_xml_content(
        self,
        content: bytes,
        segments: list[TextSegment],
        translations: dict[str, str],
    ) -> bytes:
        """Update XML content with translations."""
        root = ET.fromstring(content)
        
        # Build index of segments to update
        block_translations = {
            seg.block_id: translations.get(seg.block_id, seg.text)
            for seg in segments
        }
        
        # Update paragraphs
        para_idx = 0
        for para in root.findall('.//w:p', WORD_NAMESPACES):
            block_id = f"p{para_idx}"
            if block_id in block_translations:
                self._update_paragraph(para, block_translations[block_id])
            para_idx += 1
        
        # Update table cells
        table_idx = 0
        for table in root.findall('.//w:tbl', WORD_NAMESPACES):
            row_idx = 0
            for row in table.findall('.//w:tr', WORD_NAMESPACES):
                cell_idx = 0
                for cell in row.findall('.//w:tc', WORD_NAMESPACES):
                    block_id = f"t{table_idx}_r{row_idx}_c{cell_idx}"
                    if block_id in block_translations:
                        self._update_cell(cell, block_translations[block_id])
                    cell_idx += 1
                row_idx += 1
            table_idx += 1
        
        # Update headers/footers
        for seg in segments:
            if seg.metadata.get('type') in ('header', 'footer'):
                para_idx = 0
                for para in root.findall('.//w:p', WORD_NAMESPACES):
                    if seg.element_index == para_idx and seg.block_id in block_translations:
                        self._update_paragraph(para, block_translations[seg.block_id])
                    para_idx += 1
        
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    
    def _update_paragraph(self, para: ET.Element, new_text: str) -> None:
        """Update text in a paragraph while preserving formatting."""
        text_elements = para.findall('.//w:t', WORD_NAMESPACES)
        if not text_elements:
            return
        
        # Put all text in first element, clear the rest
        text_elements[0].text = new_text
        for t in text_elements[1:]:
            t.text = ''
    
    def _update_cell(self, cell: ET.Element, new_text: str) -> None:
        """Update text in a table cell."""
        paragraphs = cell.findall('.//w:p', WORD_NAMESPACES)
        if not paragraphs:
            return
        
        # Split text by newlines and distribute across paragraphs
        lines = new_text.split('\n')
        
        for i, para in enumerate(paragraphs):
            if i < len(lines):
                self._update_paragraph(para, lines[i])
            else:
                self._update_paragraph(para, '')


class PptxXMLHandler(OfficeXMLHandler):
    """Handler for PPTX (PowerPoint) documents."""
    
    def __init__(self, pptx_path: Path):
        self._namespaces = PPTX_NAMESPACES
        super().__init__(pptx_path)
    
    def extract(self) -> ExtractionResult:
        """Extract text from PPTX document."""
        xml_files, other_files = self._read_archive()
        segments = []
        
        # Find all slide XML files
        slide_files = sorted([
            name for name in xml_files
            if re.match(r'ppt/slides/slide\d+\.xml', name)
        ])
        
        for slide_path in slide_files:
            slide_num = int(re.search(r'slide(\d+)', slide_path).group(1))
            segments.extend(
                self._extract_from_slide(xml_files[slide_path], slide_path, slide_num)
            )
        
        # Notes
        notes_files = sorted([
            name for name in xml_files
            if re.match(r'ppt/notesSlides/notesSlide\d+\.xml', name)
        ])
        
        for notes_path in notes_files:
            notes_num = int(re.search(r'notesSlide(\d+)', notes_path).group(1))
            segments.extend(
                self._extract_from_notes(xml_files[notes_path], notes_path, notes_num)
            )
        
        return ExtractionResult(
            segments=segments,
            xml_files=xml_files,
            other_files=other_files,
        )
    
    def _extract_from_slide(
        self,
        content: bytes,
        xml_path: str,
        slide_num: int,
    ) -> list[TextSegment]:
        """Extract text from slide XML."""
        segments = []
        root = ET.fromstring(content)
        
        shape_idx = 0
        for shape in root.findall('.//p:sp', PPTX_NAMESPACES):
            texts = self._get_shape_text(shape)
            if texts.strip():
                segments.append(TextSegment(
                    block_id=f"s{slide_num}_sh{shape_idx}",
                    text=texts,
                    xml_path=xml_path,
                    xpath=f'.//p:sp[{shape_idx + 1}]',
                    element_index=shape_idx,
                    metadata={'type': 'shape', 'slide': slide_num},
                ))
            shape_idx += 1
        
        return segments
    
    def _extract_from_notes(
        self,
        content: bytes,
        xml_path: str,
        notes_num: int,
    ) -> list[TextSegment]:
        """Extract text from slide notes."""
        segments = []
        root = ET.fromstring(content)
        
        para_idx = 0
        for para in root.findall('.//a:p', PPTX_NAMESPACES):
            text = self._get_paragraph_text(para)
            if text.strip():
                segments.append(TextSegment(
                    block_id=f"notes{notes_num}_p{para_idx}",
                    text=text,
                    xml_path=xml_path,
                    xpath=f'.//a:p[{para_idx + 1}]',
                    element_index=para_idx,
                    metadata={'type': 'notes', 'slide': notes_num},
                ))
            para_idx += 1
        
        return segments
    
    def _get_shape_text(self, shape: ET.Element) -> str:
        """Get all text from a shape."""
        texts = []
        for para in shape.findall('.//a:p', PPTX_NAMESPACES):
            para_text = self._get_paragraph_text(para)
            if para_text:
                texts.append(para_text)
        return '\n'.join(texts)
    
    def _get_paragraph_text(self, para: ET.Element) -> str:
        """Get text from a paragraph."""
        texts = []
        for t in para.findall('.//a:t', PPTX_NAMESPACES):
            if t.text:
                texts.append(t.text)
        return ''.join(texts)
    
    def update(
        self,
        output_path: Path,
        translations: dict[str, str],
        extraction: ExtractionResult,
    ) -> None:
        """Update PPTX with translations."""
        segments_by_file: dict[str, list[TextSegment]] = {}
        for seg in extraction.segments:
            if seg.xml_path not in segments_by_file:
                segments_by_file[seg.xml_path] = []
            segments_by_file[seg.xml_path].append(seg)
        
        updated_xml = dict(extraction.xml_files)
        
        for xml_path, segments in segments_by_file.items():
            if xml_path in updated_xml:
                content = updated_xml[xml_path]
                updated_content = self._update_xml_content(
                    content, segments, translations
                )
                updated_xml[xml_path] = updated_content
        
        self._write_archive(output_path, updated_xml, extraction.other_files)
    
    def _update_xml_content(
        self,
        content: bytes,
        segments: list[TextSegment],
        translations: dict[str, str],
    ) -> bytes:
        """Update XML content with translations."""
        root = ET.fromstring(content)
        
        block_translations = {
            seg.block_id: translations.get(seg.block_id, seg.text)
            for seg in segments
        }
        
        # Get slide number from first segment
        slide_num = segments[0].metadata.get('slide', 1) if segments else 1
        
        if segments and segments[0].metadata.get('type') == 'notes':
            # Update notes
            para_idx = 0
            for para in root.findall('.//a:p', PPTX_NAMESPACES):
                block_id = f"notes{slide_num}_p{para_idx}"
                if block_id in block_translations:
                    self._update_paragraph(para, block_translations[block_id])
                para_idx += 1
        else:
            # Update shapes
            shape_idx = 0
            for shape in root.findall('.//p:sp', PPTX_NAMESPACES):
                block_id = f"s{slide_num}_sh{shape_idx}"
                if block_id in block_translations:
                    self._update_shape(shape, block_translations[block_id])
                shape_idx += 1
        
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    
    def _update_paragraph(self, para: ET.Element, new_text: str, scale_factor: float = 1.0) -> None:
        """Update text in a paragraph with optional font scaling."""
        text_elements = para.findall('.//a:t', PPTX_NAMESPACES)
        if not text_elements:
            return
        
        text_elements[0].text = new_text
        for t in text_elements[1:]:
            t.text = ''
        
        # Scale font sizes if needed
        if scale_factor < 1.0:
            self._scale_paragraph_fonts(para, scale_factor)
    
    def _scale_paragraph_fonts(self, para: ET.Element, scale_factor: float) -> None:
        """Scale all font sizes in a paragraph by the given factor."""
        # Minimum scale to avoid unreadable text
        scale_factor = max(scale_factor, 0.5)
        
        # Scale run properties (a:rPr) font sizes
        for rPr in para.findall('.//a:rPr', PPTX_NAMESPACES):
            sz = rPr.get('sz')
            if sz:
                try:
                    new_sz = int(int(sz) * scale_factor)
                    rPr.set('sz', str(max(new_sz, 600)))  # Min 6pt
                except ValueError:
                    pass
        
        # Scale default run properties (a:defRPr) in paragraph properties
        for defRPr in para.findall('.//a:defRPr', PPTX_NAMESPACES):
            sz = defRPr.get('sz')
            if sz:
                try:
                    new_sz = int(int(sz) * scale_factor)
                    defRPr.set('sz', str(max(new_sz, 600)))  # Min 6pt
                except ValueError:
                    pass
    
    def _update_shape(self, shape: ET.Element, new_text: str) -> None:
        """Update text in a shape with auto-fit enabled."""
        paragraphs = shape.findall('.//a:p', PPTX_NAMESPACES)
        if not paragraphs:
            return
        
        # Get original text length for scaling calculation
        original_text = self._get_shape_text(shape)
        original_len = len(original_text) if original_text else 1
        new_len = len(new_text) if new_text else 1
        
        # Calculate scale factor (if translation is longer, scale down fonts)
        scale_factor = min(1.0, original_len / new_len) if new_len > original_len else 1.0
        
        lines = new_text.split('\n')
        
        for i, para in enumerate(paragraphs):
            if i < len(lines):
                self._update_paragraph(para, lines[i], scale_factor)
            else:
                self._update_paragraph(para, '', scale_factor)
        
        # Enable auto-fit to shrink font if text overflows
        self._enable_autofit(shape)
    
    def _enable_autofit(self, shape: ET.Element) -> None:
        """Enable auto-fit on shape to shrink text to fit."""
        # Find text body
        txBody = shape.find('.//p:txBody', PPTX_NAMESPACES)
        if txBody is None:
            return
        
        # Find or create body properties
        bodyPr = txBody.find('a:bodyPr', PPTX_NAMESPACES)
        if bodyPr is None:
            # Create bodyPr as first child of txBody
            bodyPr = ET.Element('{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr')
            txBody.insert(0, bodyPr)
        
        # Remove existing auto-fit settings (noAutofit, spAutoFit, normAutofit)
        for child in list(bodyPr):
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag in ('noAutofit', 'spAutoFit', 'normAutofit'):
                bodyPr.remove(child)
        
        # Add normAutofit to enable font scaling (shrink to fit)
        # fontScale and lnSpcReduction allow shrinking up to 25%
        autofit = ET.SubElement(
            bodyPr,
            '{http://schemas.openxmlformats.org/drawingml/2006/main}normAutofit'
        )
        autofit.set('fontScale', '25000')  # Allow shrinking to 25% of original
        autofit.set('lnSpcReduction', '20000')  # Allow 20% line spacing reduction


class XlsxXMLHandler(OfficeXMLHandler):
    """Handler for XLSX (Excel) documents."""
    
    def __init__(self, xlsx_path: Path):
        self._namespaces = XLSX_NAMESPACES
        super().__init__(xlsx_path)
        self._shared_strings: list[str] = []
        self._shared_strings_modified = False
    
    def extract(self) -> ExtractionResult:
        """Extract text from XLSX document."""
        xml_files, other_files = self._read_archive()
        segments = []
        
        # Load shared strings
        if 'xl/sharedStrings.xml' in xml_files:
            self._load_shared_strings(xml_files['xl/sharedStrings.xml'])
        
        # Find all worksheet XML files
        sheet_files = sorted([
            name for name in xml_files
            if re.match(r'xl/worksheets/sheet\d+\.xml', name)
        ])
        
        for sheet_path in sheet_files:
            sheet_num = int(re.search(r'sheet(\d+)', sheet_path).group(1))
            segments.extend(
                self._extract_from_sheet(xml_files[sheet_path], sheet_path, sheet_num)
            )
        
        return ExtractionResult(
            segments=segments,
            xml_files=xml_files,
            other_files=other_files,
        )
    
    def _load_shared_strings(self, content: bytes) -> None:
        """Load shared strings table."""
        root = ET.fromstring(content)
        self._shared_strings = []
        
        for si in root.findall('.//x:si', XLSX_NAMESPACES):
            text_parts = []
            for t in si.findall('.//x:t', XLSX_NAMESPACES):
                if t.text:
                    text_parts.append(t.text)
            self._shared_strings.append(''.join(text_parts))
    
    def _extract_from_sheet(
        self,
        content: bytes,
        xml_path: str,
        sheet_num: int,
    ) -> list[TextSegment]:
        """Extract text from worksheet."""
        segments = []
        root = ET.fromstring(content)
        
        for row in root.findall('.//x:row', XLSX_NAMESPACES):
            row_num = int(row.get('r', 0))
            
            for cell in row.findall('.//x:c', XLSX_NAMESPACES):
                cell_ref = cell.get('r', '')
                cell_type = cell.get('t', '')
                
                value_elem = cell.find('.//x:v', XLSX_NAMESPACES)
                if value_elem is not None and value_elem.text:
                    if cell_type == 's':
                        # Shared string reference
                        idx = int(value_elem.text)
                        if idx < len(self._shared_strings):
                            text = self._shared_strings[idx]
                        else:
                            text = ''
                    else:
                        # Inline value
                        text = value_elem.text
                    
                    if text.strip():
                        segments.append(TextSegment(
                            block_id=f"sh{sheet_num}_{cell_ref}",
                            text=text,
                            xml_path=xml_path,
                            xpath=f'.//x:c[@r="{cell_ref}"]',
                            element_index=row_num * 1000 + self._col_to_num(cell_ref),
                            metadata={
                                'type': 'cell',
                                'sheet': sheet_num,
                                'cell': cell_ref,
                                'cell_type': cell_type,
                            },
                        ))
        
        return segments
    
    def _col_to_num(self, cell_ref: str) -> int:
        """Convert column letter to number."""
        col_str = ''.join(c for c in cell_ref if c.isalpha())
        num = 0
        for c in col_str:
            num = num * 26 + (ord(c.upper()) - ord('A') + 1)
        return num
    
    def update(
        self,
        output_path: Path,
        translations: dict[str, str],
        extraction: ExtractionResult,
    ) -> None:
        """Update XLSX with translations."""
        updated_xml = dict(extraction.xml_files)
        
        # Update shared strings first
        new_shared_strings = list(self._shared_strings)
        string_to_index = {s: i for i, s in enumerate(new_shared_strings)}
        
        # Process each segment
        segments_by_file: dict[str, list[TextSegment]] = {}
        for seg in extraction.segments:
            if seg.xml_path not in segments_by_file:
                segments_by_file[seg.xml_path] = []
            segments_by_file[seg.xml_path].append(seg)
            
            # Update shared strings mapping
            if seg.metadata.get('cell_type') == 's':
                new_text = translations.get(seg.block_id, seg.text)
                if new_text not in string_to_index:
                    string_to_index[new_text] = len(new_shared_strings)
                    new_shared_strings.append(new_text)
        
        # Write updated shared strings
        if 'xl/sharedStrings.xml' in updated_xml:
            updated_xml['xl/sharedStrings.xml'] = self._build_shared_strings(
                new_shared_strings
            )
        
        # Update each sheet
        for xml_path, segments in segments_by_file.items():
            if xml_path in updated_xml:
                content = updated_xml[xml_path]
                updated_content = self._update_sheet_content(
                    content, segments, translations, string_to_index
                )
                updated_xml[xml_path] = updated_content
        
        self._write_archive(output_path, updated_xml, extraction.other_files)
    
    def _build_shared_strings(self, strings: list[str]) -> bytes:
        """Build shared strings XML."""
        root = ET.Element('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sst')
        root.set('count', str(len(strings)))
        root.set('uniqueCount', str(len(strings)))
        
        for s in strings:
            si = ET.SubElement(root, '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si')
            t = ET.SubElement(si, '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
            t.text = s
        
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    
    def _update_sheet_content(
        self,
        content: bytes,
        segments: list[TextSegment],
        translations: dict[str, str],
        string_to_index: dict[str, int],
    ) -> bytes:
        """Update sheet XML content."""
        root = ET.fromstring(content)
        
        # Build lookup
        cell_translations = {}
        for seg in segments:
            cell_ref = seg.metadata.get('cell', '')
            new_text = translations.get(seg.block_id, seg.text)
            cell_translations[cell_ref] = (new_text, seg.metadata.get('cell_type'))
        
        # Update cells
        for cell in root.findall('.//x:c', XLSX_NAMESPACES):
            cell_ref = cell.get('r', '')
            if cell_ref in cell_translations:
                new_text, cell_type = cell_translations[cell_ref]
                value_elem = cell.find('.//x:v', XLSX_NAMESPACES)
                
                if value_elem is not None:
                    if cell_type == 's':
                        # Update shared string index
                        if new_text in string_to_index:
                            value_elem.text = str(string_to_index[new_text])
                    else:
                        # Update inline value
                        value_elem.text = new_text
        
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)


def get_handler(office_path: Path) -> OfficeXMLHandler:
    """
    Get appropriate handler for Office file type.
    
    Args:
        office_path: Path to Office document.
        
    Returns:
        Appropriate handler instance.
    """
    ext = office_path.suffix.lower()
    
    if ext == '.docx':
        return DocxXMLHandler(office_path)
    elif ext == '.pptx':
        return PptxXMLHandler(office_path)
    elif ext == '.xlsx':
        return XlsxXMLHandler(office_path)
    else:
        raise ValueError(f"Unsupported Office format: {ext}")
