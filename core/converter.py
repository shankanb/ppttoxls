import json
from typing import Dict, Any, Union, BinaryIO, Tuple
import openpyxl

from core.models import MappingConfig
from core.ppt_extractor import extract_presentation_content, DeckContent
from core.llm_extractor import extract_data_with_llm
from core.excel_builder import build_excel_workbook, export_workbook_to_bytes


def convert_ppt_to_xls(
    ppt_source: Union[str, BinaryIO, bytes],
    mapping_source: Union[str, Dict[str, Any], bytes, BinaryIO],
    api_key: str = None,
    extracted_data_override: Dict[str, Any] = None,
) -> Tuple[openpyxl.Workbook, Dict[str, Any], DeckContent]:
    """
    Main orchestration function converting a PPT presentation to a styled Excel workbook.
    Returns (workbook, extracted_data_dict, deck_content).
    """
    # 1. Parse Mapping Configuration
    if isinstance(mapping_source, (str, bytes)):
        if isinstance(mapping_source, bytes):
            mapping_dict = json.loads(mapping_source.decode("utf-8"))
        elif mapping_source.strip().startswith("{"):
            mapping_dict = json.loads(mapping_source)
        else:
            # Assume it's a file path
            with open(mapping_source, "r", encoding="utf-8") as f:
                mapping_dict = json.load(f)
    elif isinstance(mapping_source, dict):
        mapping_dict = mapping_source
    elif hasattr(mapping_source, "read"):
        content = mapping_source.read()
        if isinstance(content, bytes):
            content = content.decode("utf-8")
        mapping_dict = json.loads(content)
    else:
        raise ValueError(f"Unsupported mapping_source type: {type(mapping_source)}")

    config = MappingConfig(**mapping_dict)

    # 2. Extract PPT Content
    deck = extract_presentation_content(ppt_source)

    # 3. Extract Data using Gemini LLM (or use override if testing/mocking)
    if extracted_data_override is not None:
        extracted_data = extracted_data_override
    else:
        extracted_data = extract_data_with_llm(deck=deck, config=config, api_key=api_key)

    # 4. Build Styled Excel Workbook
    wb = build_excel_workbook(config=config, extracted_data=extracted_data)

    return wb, extracted_data, deck
