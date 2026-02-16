"""
Pipeline base classes and enums for PDF translation.

Supported Pipelines:
- DIRECT_PDF: Direct PDF text replacement using PyMuPDF
- DOCX_ROUNDTRIP: PDF → DOCX → translate XML → DOCX → PDF
- XLIFF: Generate XLIFF format for professional CAT tools
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Optional


class PipelineType(Enum):
    """Available translation pipelines."""
    DIRECT_PDF = "direct"       # Direct PDF manipulation (current approach)
    DOCX_ROUNDTRIP = "docx"     # PDF → DOCX → translate → DOCX → PDF
    XLIFF = "xliff"             # Generate XLIFF format for CAT tools


@dataclass
class PipelineConfig:
    """Base configuration for all pipelines."""
    
    # Pipeline selection
    pipeline_type: PipelineType = PipelineType.DIRECT_PDF
    
    # Common settings
    target_language: str = "Hindi"
    min_font_size: float = 6.0
    font_step: float = 0.5
    
    # Output settings
    output_dir: Optional[Path] = None  # Uses input file directory if None


@dataclass
class ExtractResult:
    """Result from extract phase."""
    layout_path: Path
    translate_path: Path
    translated_template_path: Path
    extra_files: dict[str, Path] = field(default_factory=dict)  # Pipeline-specific files


@dataclass
class MergeResult:
    """Result from merge phase."""
    output_path: Path
    blocks_processed: int
    warnings: list[str] = field(default_factory=list)


class TranslationPipeline(ABC):
    """
    Abstract base class for translation pipelines.
    
    All pipelines follow a two-phase workflow:
    1. Extract: Generate translation files from input document
    2. Merge: Rebuild document with translations
    """
    
    def __init__(self, config: PipelineConfig):
        self.config = config
    
    @property
    @abstractmethod
    def name(self) -> str:
        """Human-readable pipeline name."""
        pass
    
    @property
    @abstractmethod
    def description(self) -> str:
        """Pipeline description for help text."""
        pass
    
    @abstractmethod
    def extract(self, input_path: Path) -> ExtractResult:
        """
        Extract text blocks and generate translation files.
        
        Args:
            input_path: Path to input document (PDF, DOCX, etc.)
            
        Returns:
            ExtractResult with paths to generated files.
        """
        pass
    
    @abstractmethod
    def merge(
        self,
        input_path: Path,
        output_path: Path,
        translated_path: Path,
        layout_path: Path,
    ) -> MergeResult:
        """
        Merge translations and rebuild document.
        
        Args:
            input_path: Path to original input document.
            output_path: Path for output document.
            translated_path: Path to file with translations.
            layout_path: Path to layout/metadata file.
            
        Returns:
            MergeResult with output path and statistics.
        """
        pass
    
    @abstractmethod
    def get_translation_prompt(self, block_count: int) -> str:
        """
        Get the LLM prompt for translation.
        
        Args:
            block_count: Number of text blocks to translate.
            
        Returns:
            Prompt string for the LLM.
        """
        pass
    
    def derive_paths(self, input_path: Path) -> dict[str, Path]:
        """
        Derive intermediate file paths from input path.
        
        Can be overridden by subclasses for custom naming.
        """
        base = input_path.parent / input_path.name
        return {
            "layout": Path(f"{base}_layout.json"),
            "translate": Path(f"{base}_translate.txt"),
            "translated": Path(f"{base}_translated.txt"),
            "translations": Path(f"{base}_translations.json"),
        }
