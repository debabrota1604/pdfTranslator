"""
Direct PDF Pipeline - Current implementation.

Uses PyMuPDF for direct PDF text replacement with Unicode support.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from .base import (
    TranslationPipeline,
    PipelineConfig,
    ExtractResult,
    MergeResult,
)
from ..extractor import extract_pdf_layout
from ..rebuilder_unicode import (
    PDFRebuilder,
    RebuildConfig,
    RenderMethod,
    _find_unicode_font,
)
from ..translation_io import (
    generate_translate_file,
    generate_translated_template,
    parse_translated_file,
    get_translation_prompt,
)


@dataclass
class DirectPDFConfig(PipelineConfig):
    """Configuration specific to Direct PDF pipeline."""
    
    # Render method
    render_method: RenderMethod = RenderMethod.LINE_BY_LINE
    
    # Font settings
    unicode_font_path: Optional[Path] = field(default_factory=_find_unicode_font)
    fallback_font: str = "helv"
    
    # Visual
    overlay_color: tuple[float, float, float] = (1.0, 1.0, 1.0)


class DirectPDFPipeline(TranslationPipeline):
    """
    Direct PDF manipulation pipeline.
    
    Uses PyMuPDF to:
    1. Extract text blocks with bounding boxes
    2. Redact original text
    3. Insert translated text with Unicode font support
    
    Pros:
    - Fast, no external dependencies beyond PyMuPDF
    - Preserves original PDF structure (images, vectors, etc.)
    
    Cons:
    - Line-by-line replacement may not handle length changes well
    - Limited text reflow capabilities
    """
    
    def __init__(self, config: DirectPDFConfig):
        super().__init__(config)
        self.config: DirectPDFConfig = config
    
    @property
    def name(self) -> str:
        return "Direct PDF"
    
    @property
    def description(self) -> str:
        return "Direct PDF text replacement using PyMuPDF with Unicode support"
    
    def extract(self, input_path: Path) -> ExtractResult:
        """Extract layout and generate translation files."""
        paths = self.derive_paths(input_path)
        
        # Extract layout
        extract_pdf_layout(input_path, paths["layout"])
        
        # Generate translation file
        generate_translate_file(paths["layout"], paths["translate"])
        
        # Generate template for translated output
        generate_translated_template(paths["layout"], paths["translated"])
        
        return ExtractResult(
            layout_path=paths["layout"],
            translate_path=paths["translate"],
            translated_template_path=paths["translated"],
        )
    
    def merge(
        self,
        input_path: Path,
        output_path: Path,
        translated_path: Path,
        layout_path: Path,
    ) -> MergeResult:
        """Merge translations and rebuild PDF."""
        paths = self.derive_paths(input_path)
        
        # Parse translations
        translations = parse_translated_file(
            translated_path,
            layout_path,
            paths["translations"],
        )
        
        # Build rebuild config
        rebuild_config = RebuildConfig(
            render_method=self.config.render_method,
            min_font_size=self.config.min_font_size,
            font_step=self.config.font_step,
            unicode_font_path=self.config.unicode_font_path,
            fallback_font=self.config.fallback_font,
            overlay_color=self.config.overlay_color,
        )
        
        # Rebuild PDF
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
    
    def get_translation_prompt(self, block_count: int) -> str:
        """Get LLM prompt for translation."""
        return get_translation_prompt(block_count, self.config.target_language)


def create_direct_pdf_pipeline(
    target_language: str = "Hindi",
    render_method: RenderMethod = RenderMethod.LINE_BY_LINE,
    min_font_size: float = 6.0,
    font_step: float = 0.5,
) -> DirectPDFPipeline:
    """Factory function to create DirectPDF pipeline with common settings."""
    config = DirectPDFConfig(
        target_language=target_language,
        render_method=render_method,
        min_font_size=min_font_size,
        font_step=font_step,
    )
    return DirectPDFPipeline(config)
