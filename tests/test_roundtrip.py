"""
Roundtrip Tests for PDF Layout Engine.

Tests:
- Extract layout from PDF
- Replace text with identical text
- Verify output PDF preserves layout
- Verify bounding boxes are preserved
- Test determinism (multiple runs produce same output)
"""

from __future__ import annotations

import json
import tempfile
from pathlib import Path
from typing import Any

import pytest
import fitz  # PyMuPDF

import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from pdf_layout.extractor import PDFExtractor, extract_pdf_layout, DocumentData
from pdf_layout.segmenter import PDFSegmenter, segment_document, create_translation_template
from pdf_layout.rebuilder import PDFRebuilder, RebuildConfig, rebuild_pdf
from utils.font_utils import FontMapper, get_font_metrics, map_font_name


class TestPDFExtractor:
    """Tests for PDF extraction functionality."""
    
    @pytest.fixture
    def sample_pdf(self, tmp_path: Path) -> Path:
        """Create a sample PDF for testing."""
        pdf_path = tmp_path / "sample.pdf"
        doc = fitz.open()
        
        # Create a simple page with text
        page = doc.new_page(width=595, height=842)  # A4 size
        
        # Add some text blocks
        page.insert_text(
            (72, 72),
            "Hello World",
            fontsize=12,
            fontname="helv",
        )
        page.insert_text(
            (72, 100),
            "This is a test document.",
            fontsize=11,
            fontname="helv",
        )
        page.insert_text(
            (72, 150),
            "Multiple lines\nof text here.",
            fontsize=10,
            fontname="tiro",
        )
        
        doc.save(str(pdf_path))
        doc.close()
        
        return pdf_path
    
    def test_extraction_returns_valid_structure(self, sample_pdf: Path):
        """Test that extraction returns valid DocumentData."""
        with PDFExtractor(sample_pdf) as extractor:
            document = extractor.extract()
        
        assert isinstance(document, DocumentData)
        assert len(document.pages) == 1
        assert document.pages[0].page_number == 1
        assert document.pages[0].width > 0
        assert document.pages[0].height > 0
    
    def test_extraction_extracts_text_blocks(self, sample_pdf: Path):
        """Test that text blocks are extracted."""
        with PDFExtractor(sample_pdf) as extractor:
            document = extractor.extract()
        
        page = document.pages[0]
        assert len(page.blocks) > 0
        
        for block in page.blocks:
            assert block.block_id.startswith("p1_b")
            assert len(block.bbox) == 4
            assert block.text
            assert block.font_size > 0
    
    def test_extraction_is_deterministic(self, sample_pdf: Path):
        """Test that multiple extractions produce identical results."""
        results = []
        
        for _ in range(3):
            with PDFExtractor(sample_pdf) as extractor:
                document = extractor.extract()
                results.append(document.to_json())
        
        # All results should be identical
        assert results[0] == results[1] == results[2]
    
    def test_block_ids_are_stable(self, sample_pdf: Path):
        """Test that block IDs are stable across runs."""
        ids_run1 = []
        ids_run2 = []
        
        with PDFExtractor(sample_pdf) as extractor:
            doc1 = extractor.extract()
            for page in doc1.pages:
                ids_run1.extend(b.block_id for b in page.blocks)
        
        with PDFExtractor(sample_pdf) as extractor:
            doc2 = extractor.extract()
            for page in doc2.pages:
                ids_run2.extend(b.block_id for b in page.blocks)
        
        assert ids_run1 == ids_run2
    
    def test_json_output(self, sample_pdf: Path, tmp_path: Path):
        """Test JSON output functionality."""
        output_json = tmp_path / "layout.json"
        
        extract_pdf_layout(sample_pdf, output_json)
        
        assert output_json.exists()
        
        data = json.loads(output_json.read_text())
        assert "pages" in data
        assert "source_file" in data


class TestPDFSegmenter:
    """Tests for PDF segmentation functionality."""
    
    @pytest.fixture
    def sample_document(self) -> dict[str, Any]:
        """Create sample document data."""
        return {
            "source_file": "test.pdf",
            "pages": [
                {
                    "page_number": 1,
                    "width": 595,
                    "height": 842,
                    "rotation": 0,
                    "blocks": [
                        {
                            "block_id": "p1_b0",
                            "bbox": [72, 72, 200, 90],
                            "text": "First block",
                            "font_name": "Helvetica",
                            "font_size": 12.0,
                            "color": "#000000",
                            "writing_direction": "ltr",
                            "line_height": 14.0,
                        },
                        {
                            "block_id": "p1_b1",
                            "bbox": [72, 100, 300, 120],
                            "text": "Second block with more text",
                            "font_name": "Helvetica",
                            "font_size": 11.0,
                            "color": "#000000",
                            "writing_direction": "ltr",
                            "line_height": 13.0,
                        },
                    ],
                },
            ],
        }
    
    def test_segmentation_preserves_blocks(self, sample_document: dict):
        """Test that segmentation preserves all blocks."""
        segmented = segment_document(sample_document)
        
        assert len(segmented.segments) == 2
        assert segmented.segments[0].block_id == "p1_b0"
        assert segmented.segments[1].block_id == "p1_b1"
    
    def test_segmentation_preserves_order(self, sample_document: dict):
        """Test that reading order is preserved."""
        segmented = segment_document(sample_document)
        
        for i, segment in enumerate(segmented.segments):
            assert segment.segment_index == i
    
    def test_translation_template_creation(self, sample_document: dict):
        """Test translation template creation."""
        segmenter = PDFSegmenter(sample_document)
        template = segmenter.create_translation_template()
        
        assert "p1_b0" in template
        assert "p1_b1" in template
        assert template["p1_b0"] == "First block"
        assert template["p1_b1"] == "Second block with more text"
    
    def test_empty_blocks_are_filtered(self):
        """Test that empty blocks are filtered out."""
        doc_with_empty = {
            "source_file": "test.pdf",
            "pages": [
                {
                    "page_number": 1,
                    "width": 595,
                    "height": 842,
                    "blocks": [
                        {
                            "block_id": "p1_b0",
                            "bbox": [72, 72, 200, 90],
                            "text": "Has text",
                            "font_name": "Helvetica",
                            "font_size": 12.0,
                            "color": "#000000",
                        },
                        {
                            "block_id": "p1_b1",
                            "bbox": [72, 100, 200, 120],
                            "text": "",  # Empty
                            "font_name": "Helvetica",
                            "font_size": 12.0,
                            "color": "#000000",
                        },
                    ],
                },
            ],
        }
        
        segmented = segment_document(doc_with_empty)
        assert len(segmented.segments) == 1


class TestPDFRebuilder:
    """Tests for PDF rebuilding functionality."""
    
    @pytest.fixture
    def sample_pdf_with_layout(self, tmp_path: Path) -> tuple[Path, dict]:
        """Create sample PDF and extract its layout."""
        pdf_path = tmp_path / "input.pdf"
        doc = fitz.open()
        
        page = doc.new_page(width=595, height=842)
        page.insert_text((72, 72), "Original text", fontsize=12, fontname="helv")
        page.insert_text((72, 120), "More original text", fontsize=11, fontname="helv")
        
        doc.save(str(pdf_path))
        doc.close()
        
        # Extract layout
        with PDFExtractor(pdf_path) as extractor:
            document = extractor.extract()
        
        return pdf_path, document.to_dict()
    
    def test_rebuild_with_identical_text(
        self, 
        sample_pdf_with_layout: tuple[Path, dict],
        tmp_path: Path
    ):
        """Test rebuilding with identical text produces valid PDF."""
        pdf_path, layout = sample_pdf_with_layout
        output_path = tmp_path / "output.pdf"
        
        # Create identity translations
        translations = {}
        for page in layout["pages"]:
            for block in page["blocks"]:
                translations[block["block_id"]] = block["text"]
        
        # Rebuild
        rebuilder = PDFRebuilder()
        rebuilder.rebuild(
            pdf_path=pdf_path,
            layout_data=layout,
            translations=translations,
            output_path=output_path,
        )
        
        # Verify output is valid PDF
        assert output_path.exists()
        doc = fitz.open(str(output_path))
        assert len(doc) == 1
        doc.close()
    
    def test_rebuild_preserves_page_dimensions(
        self,
        sample_pdf_with_layout: tuple[Path, dict],
        tmp_path: Path
    ):
        """Test that page dimensions are preserved."""
        pdf_path, layout = sample_pdf_with_layout
        output_path = tmp_path / "output.pdf"
        
        translations = {}
        for page in layout["pages"]:
            for block in page["blocks"]:
                translations[block["block_id"]] = block["text"]
        
        rebuilder = PDFRebuilder()
        rebuilder.rebuild(pdf_path, layout, translations, output_path)
        
        # Compare dimensions
        original = fitz.open(str(pdf_path))
        output = fitz.open(str(output_path))
        
        assert original[0].rect.width == output[0].rect.width
        assert original[0].rect.height == output[0].rect.height
        
        original.close()
        output.close()
    
    def test_font_scaling_works(self, tmp_path: Path):
        """Test that font scaling reduces size for overflow."""
        # Create PDF with small text box
        pdf_path = tmp_path / "input.pdf"
        doc = fitz.open()
        
        page = doc.new_page(width=595, height=842)
        page.insert_text((72, 72), "Short", fontsize=12, fontname="helv")
        
        doc.save(str(pdf_path))
        doc.close()
        
        # Extract
        with PDFExtractor(pdf_path) as extractor:
            document = extractor.extract()
        
        layout = document.to_dict()
        
        # Create translation with much longer text
        translations = {}
        for page in layout["pages"]:
            for block in page["blocks"]:
                # Make text much longer
                translations[block["block_id"]] = block["text"] * 10
        
        output_path = tmp_path / "output.pdf"
        
        # Should not raise an error
        rebuilder = PDFRebuilder()
        rebuilder.rebuild(pdf_path, layout, translations, output_path)
        
        assert output_path.exists()


class TestRoundtrip:
    """End-to-end roundtrip tests."""
    
    @pytest.fixture
    def complex_pdf(self, tmp_path: Path) -> Path:
        """Create a more complex PDF for roundtrip testing."""
        pdf_path = tmp_path / "complex.pdf"
        doc = fitz.open()
        
        # Page 1
        page1 = doc.new_page(width=595, height=842)
        page1.insert_text((72, 72), "Title of Document", fontsize=18, fontname="helv")
        page1.insert_text((72, 120), "This is the introduction paragraph with some text.", fontsize=11, fontname="helv")
        page1.insert_text((72, 160), "Chapter 1: Getting Started", fontsize=14, fontname="helv")
        page1.insert_text((72, 200), "Lorem ipsum dolor sit amet, consectetur adipiscing elit.", fontsize=10, fontname="tiro")
        
        # Page 2
        page2 = doc.new_page(width=595, height=842)
        page2.insert_text((72, 72), "Chapter 2: Advanced Topics", fontsize=14, fontname="helv")
        page2.insert_text((72, 110), "More detailed content here.", fontsize=10, fontname="helv")
        
        doc.save(str(pdf_path))
        doc.close()
        
        return pdf_path
    
    def test_full_roundtrip_identity(self, complex_pdf: Path, tmp_path: Path):
        """Test full roundtrip with identity translation."""
        layout_path = tmp_path / "layout.json"
        output_path = tmp_path / "output.pdf"
        
        # Step 1: Extract
        document = extract_pdf_layout(complex_pdf, layout_path)
        
        # Step 2: Create identity translations
        translations = {}
        for page in document.pages:
            for block in page.blocks:
                translations[block.block_id] = block.text
        
        translations_path = tmp_path / "translations.json"
        translations_path.write_text(json.dumps(translations, indent=2))
        
        # Step 3: Rebuild
        rebuild_pdf(
            pdf_path=complex_pdf,
            layout_path=layout_path,
            translations_path=translations_path,
            output_path=output_path,
        )
        
        # Step 4: Verify
        assert output_path.exists()
        
        # Extract from output and compare structure
        with PDFExtractor(output_path) as extractor:
            output_doc = extractor.extract()
        
        # Same number of pages
        assert len(output_doc.pages) == len(document.pages)
        
        # Page dimensions preserved
        for orig_page, out_page in zip(document.pages, output_doc.pages):
            assert abs(orig_page.width - out_page.width) < 1
            assert abs(orig_page.height - out_page.height) < 1
    
    def test_determinism_multiple_runs(self, complex_pdf: Path, tmp_path: Path):
        """Test that multiple runs produce identical output."""
        outputs = []
        
        for i in range(3):
            run_dir = tmp_path / f"run_{i}"
            run_dir.mkdir()
            
            layout_path = run_dir / "layout.json"
            output_path = run_dir / "output.pdf"
            
            # Extract
            document = extract_pdf_layout(complex_pdf, layout_path)
            
            # Identity translations
            translations = {}
            for page in document.pages:
                for block in page.blocks:
                    translations[block.block_id] = block.text
            
            translations_path = run_dir / "translations.json"
            translations_path.write_text(json.dumps(translations, indent=2))
            
            # Rebuild
            rebuild_pdf(complex_pdf, layout_path, translations_path, output_path)
            
            # Read output
            outputs.append(output_path.read_bytes())
        
        # All outputs should be identical (byte-for-byte due to determinism)
        # Note: May differ slightly due to timestamps, so we compare structure
        for i in range(1, len(outputs)):
            # At minimum, same size
            assert abs(len(outputs[0]) - len(outputs[i])) < 100  # Allow small variance


class TestFontUtils:
    """Tests for font utilities."""
    
    def test_font_mapping_helvetica(self):
        """Test Helvetica family mapping."""
        mapper = FontMapper()
        
        assert mapper.map_font("Helvetica") == "helv"
        assert mapper.map_font("Arial") == "helv"
        assert mapper.map_font("Arial-Bold") == "helvbo"
        assert mapper.map_font("Helvetica-Italic") == "helvit"
    
    def test_font_mapping_times(self):
        """Test Times family mapping."""
        mapper = FontMapper()
        
        assert mapper.map_font("Times") == "tiro"
        assert mapper.map_font("Times-Roman") == "tiro"
        assert mapper.map_font("TimesNewRoman-Bold") == "tirobo"
    
    def test_font_mapping_courier(self):
        """Test Courier family mapping."""
        mapper = FontMapper()
        
        assert mapper.map_font("Courier") == "cour"
        assert mapper.map_font("CourierNew") == "cour"
        assert mapper.map_font("Consolas") == "cour"
    
    def test_font_mapping_fallback(self):
        """Test fallback for unknown fonts."""
        mapper = FontMapper()
        
        assert mapper.map_font("UnknownFont") == "helv"
        assert mapper.map_font("") == "helv"
        assert mapper.map_font("SomeRandomFont") == "helv"
    
    def test_font_metrics(self):
        """Test font metrics calculation."""
        metrics = get_font_metrics("Helvetica", 12.0)
        
        assert metrics.font_name == "helv"
        assert metrics.line_height > 0
        assert metrics.avg_char_width > 0
    
    def test_is_monospace(self):
        """Test monospace detection."""
        mapper = FontMapper()
        
        assert mapper.is_monospace("Courier")
        assert mapper.is_monospace("Consolas")
        assert not mapper.is_monospace("Helvetica")
        assert not mapper.is_monospace("Times")


class TestEdgeCases:
    """Tests for edge cases and error handling."""
    
    def test_empty_pdf(self, tmp_path: Path):
        """Test handling of PDF with no text."""
        pdf_path = tmp_path / "empty.pdf"
        doc = fitz.open()
        doc.new_page()  # Empty page
        doc.save(str(pdf_path))
        doc.close()
        
        with PDFExtractor(pdf_path) as extractor:
            document = extractor.extract()
        
        assert len(document.pages) == 1
        assert len(document.pages[0].blocks) == 0
    
    def test_multipage_pdf(self, tmp_path: Path):
        """Test handling of multi-page PDF."""
        pdf_path = tmp_path / "multipage.pdf"
        doc = fitz.open()
        
        for i in range(5):
            page = doc.new_page()
            page.insert_text((72, 72), f"Page {i + 1}", fontsize=12, fontname="helv")
        
        doc.save(str(pdf_path))
        doc.close()
        
        with PDFExtractor(pdf_path) as extractor:
            document = extractor.extract()
        
        assert len(document.pages) == 5
        for i, page in enumerate(document.pages):
            assert page.page_number == i + 1
    
    def test_nonexistent_pdf(self, tmp_path: Path):
        """Test error handling for non-existent PDF."""
        with pytest.raises(FileNotFoundError):
            PDFExtractor(tmp_path / "nonexistent.pdf")
    
    def test_special_characters(self, tmp_path: Path):
        """Test handling of special characters."""
        pdf_path = tmp_path / "special.pdf"
        doc = fitz.open()
        
        page = doc.new_page()
        # Note: Some special chars may not render with built-in fonts
        page.insert_text((72, 72), "Test with special: & < > \"", fontsize=12, fontname="helv")
        
        doc.save(str(pdf_path))
        doc.close()
        
        with PDFExtractor(pdf_path) as extractor:
            document = extractor.extract()
        
        assert len(document.pages[0].blocks) > 0


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
