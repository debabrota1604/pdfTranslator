"""
PDF Translation Pipelines.

Available pipelines:
- DirectPDFPipeline: Direct PDF text replacement (default)
- DocxRoundtripPipeline: PDF → DOCX → translate → DOCX → PDF
- OfficeRoundtripPipeline: PDF → Office (auto-detect DOCX/PPTX/XLSX) → translate → PDF
- XLIFFPipeline: Generate XLIFF format for CAT tools

Office XML Handlers:
- DocxXMLHandler: Extract/update text in Word documents
- PptxXMLHandler: Extract/update text in PowerPoint presentations
- XlsxXMLHandler: Extract/update text in Excel spreadsheets
"""

from .base import (
    PipelineType,
    PipelineConfig,
    TranslationPipeline,
    ExtractResult,
    MergeResult,
)
from .direct_pdf import (
    DirectPDFPipeline,
    DirectPDFConfig,
    create_direct_pdf_pipeline,
)
from .docx_roundtrip import (
    DocxRoundtripPipeline,
    DocxRoundtripConfig,
    create_docx_roundtrip_pipeline,
)
from .office_roundtrip import (
    OfficeRoundtripPipeline,
    OfficeRoundtripConfig,
    OfficeFormat,
    create_office_roundtrip_pipeline,
)
from .office_xml import (
    OfficeXMLHandler,
    DocxXMLHandler,
    PptxXMLHandler,
    XlsxXMLHandler,
    TextSegment,
    ExtractionResult,
    get_handler,
)
from .xliff_format import (
    XLIFFPipeline,
    XLIFFConfig,
    XLIFFVersion,
    create_xliff_pipeline,
)


def create_pipeline(
    pipeline_type: PipelineType,
    target_language: str = "Hindi",
    **kwargs,
) -> TranslationPipeline:
    """
    Factory function to create a pipeline by type.
    
    Args:
        pipeline_type: Type of pipeline to create.
        target_language: Target language for translation.
        **kwargs: Additional pipeline-specific arguments.
        
    Returns:
        Configured TranslationPipeline instance.
    """
    if pipeline_type == PipelineType.DIRECT_PDF:
        return create_direct_pdf_pipeline(
            target_language=target_language,
            **kwargs,
        )
    elif pipeline_type == PipelineType.DOCX_ROUNDTRIP:
        # Use new Office pipeline with auto-detection
        return create_office_roundtrip_pipeline(
            target_language=target_language,
            **kwargs,
        )
    elif pipeline_type == PipelineType.XLIFF:
        return create_xliff_pipeline(
            target_language=target_language,
            **kwargs,
        )
    else:
        raise ValueError(f"Unknown pipeline type: {pipeline_type}")


__all__ = [
    # Base classes
    "PipelineType",
    "PipelineConfig",
    "TranslationPipeline",
    "ExtractResult",
    "MergeResult",
    # DirectPDF
    "DirectPDFPipeline",
    "DirectPDFConfig",
    "create_direct_pdf_pipeline",
    # DOCX Roundtrip (legacy)
    "DocxRoundtripPipeline",
    "DocxRoundtripConfig",
    "create_docx_roundtrip_pipeline",
    # Office Roundtrip (auto-detect DOCX/PPTX/XLSX)
    "OfficeRoundtripPipeline",
    "OfficeRoundtripConfig",
    "OfficeFormat",
    "create_office_roundtrip_pipeline",
    # Office XML Handlers
    "OfficeXMLHandler",
    "DocxXMLHandler",
    "PptxXMLHandler",
    "XlsxXMLHandler",
    "TextSegment",
    "ExtractionResult",
    "get_handler",
    # XLIFF
    "XLIFFPipeline",
    "XLIFFConfig",
    "XLIFFVersion",
    "create_xliff_pipeline",
    # Factory
    "create_pipeline",
]
