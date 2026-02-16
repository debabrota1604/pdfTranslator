"""
Translation I/O Module.

Minimal tagged format for efficient LLM translation.
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Optional, Union


def generate_translate_file(
    layout_path: Union[str, Path],
    output_path: Union[str, Path],
) -> tuple[str, list[str]]:
    """
    Generate minimal tagged file for translation.
    
    Format: <id>text</id> (one block per line when possible)
    Also updates layout.json with block_order for mapping.
    
    Args:
        layout_path: Path to layout.json file.
        output_path: Path for output file.
        
    Returns:
        Tuple of (content, list of block_ids in order).
    """
    layout_path = Path(layout_path)
    output_path = Path(output_path)
    
    layout = json.loads(layout_path.read_text(encoding="utf-8"))
    
    lines = []
    block_ids = []
    
    for page in layout.get("pages", []):
        for block in page.get("blocks", []):
            block_id = block.get("block_id", "")
            text = block.get("text", "")
            
            if not text.strip():
                continue
            
            block_ids.append(block_id)
            idx = len(block_ids) - 1
            
            # Replace newlines with \\n for compact single-line format
            text_compact = text.replace("\n", "\\n")
            lines.append(f"<{idx}>{text_compact}</{idx}>")
    
    content = "\n".join(lines)
    
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(content, encoding="utf-8")
    
    # Store block_order in layout.json itself
    layout["block_order"] = block_ids
    layout_path.write_text(
        json.dumps(layout, indent=2, ensure_ascii=False),
        encoding="utf-8"
    )
    
    return content, block_ids


def generate_translated_template(
    layout_path: Union[str, Path],
    output_path: Union[str, Path],
) -> str:
    """
    Generate minimal template for translated text.
    
    Args:
        layout_path: Path to layout.json file.
        output_path: Path for output file.
        
    Returns:
        The template content.
    """
    layout_path = Path(layout_path)
    output_path = Path(output_path)
    
    layout = json.loads(layout_path.read_text(encoding="utf-8"))
    
    lines = []
    idx = 0
    
    for page in layout.get("pages", []):
        for block in page.get("blocks", []):
            text = block.get("text", "")
            if not text.strip():
                continue
            lines.append(f"<{idx}></{idx}>")
            idx += 1
    
    content = "\n".join(lines)
    
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(content, encoding="utf-8")
    
    return content


def parse_translated_file(
    translated_path: Union[str, Path],
    layout_path: Union[str, Path],
    output_path: Optional[Union[str, Path]] = None,
) -> dict[str, str]:
    """
    Parse translated file back to translations dictionary.
    
    Args:
        translated_path: Path to translated file.
        layout_path: Path to layout.json (contains block_order mapping).
        output_path: Optional path to save translations.json.
        
    Returns:
        Dictionary mapping block_id to translated text.
    """
    translated_path = Path(translated_path)
    layout_path = Path(layout_path)
    
    content = translated_path.read_text(encoding="utf-8")
    layout = json.loads(layout_path.read_text(encoding="utf-8"))
    block_ids = layout.get("block_order", [])
    
    translations = {}
    
    # Pattern: <idx>content</idx>
    pattern = r'<(\d+)>(.*?)</\1>'
    matches = re.findall(pattern, content, re.DOTALL)
    
    for idx_str, text in matches:
        idx = int(idx_str)
        if idx < len(block_ids):
            # Convert \\n back to actual newlines
            text_expanded = text.replace("\\n", "\n")
            translations[block_ids[idx]] = text_expanded.strip()
    
    if output_path:
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(
            json.dumps(translations, indent=2, ensure_ascii=False),
            encoding="utf-8"
        )
    
    return translations


def apply_translations_to_layout(
    layout_path: Union[str, Path],
    translations: Union[dict[str, str], str, Path],
    output_path: Union[str, Path],
) -> dict[str, Any]:
    """
    Apply translations to layout JSON.
    """
    layout_path = Path(layout_path)
    output_path = Path(output_path)
    
    layout = json.loads(layout_path.read_text(encoding="utf-8"))
    
    if isinstance(translations, (str, Path)):
        translations_path = Path(translations)
        translations = json.loads(translations_path.read_text(encoding="utf-8"))
    
    for page in layout.get("pages", []):
        for block in page.get("blocks", []):
            block_id = block.get("block_id", "")
            if block_id in translations:
                block["text"] = translations[block_id]
    
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(layout, indent=2, ensure_ascii=False),
        encoding="utf-8"
    )
    
    return layout


# Path to editable prompt template
PROMPT_TEMPLATE_PATH = Path(__file__).parent / "prompt_template.txt"


def get_translation_prompt(total_blocks: int, target_language: str = "Hindi") -> str:
    """Generate the prompt for LLM translation.
    
    Reads from prompt_template.txt which can be edited by developers.
    Supports {target_language} placeholder in the template.
    """
    if PROMPT_TEMPLATE_PATH.exists():
        template = PROMPT_TEMPLATE_PATH.read_text(encoding="utf-8")
        return template.format(target_language=target_language)
    
    # Fallback if template file is missing
    return f"""Translate to {target_language}. Rules:
- Format: <N>text</N> where N is number
- Translate ONLY text between tags
- Keep <N> and </N> tags exactly as-is
- \\n = line break, preserve them
- One block per line, no extra text"""
