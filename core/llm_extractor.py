import os
import json
import logging
from typing import Dict, Any, Optional
from dotenv import load_dotenv

from core.models import MappingConfig
from core.ppt_extractor import DeckContent

load_dotenv()
logger = logging.getLogger(__name__)


def build_extraction_prompt(deck: DeckContent, config: MappingConfig) -> str:
    """Constructs a deterministic prompt detailing the slide content and target JSON schema."""
    deck_text = deck.to_markdown()

    # Describe expected schema structure
    schema_desc = {}
    for sheet in config.sheets:
        for cell_map in sheet.cells:
            schema_desc[cell_map.key] = {
                "type": cell_map.type,
                "description": cell_map.description or f"Value for cell {cell_map.cell}",
            }
        for table in sheet.tables:
            col_specs = {}
            for col in table.columns:
                if col.type != "formula":  # formulas are computed by Python/Excel
                    col_specs[col.key] = {
                        "type": col.type,
                        "description": col.description or f"Column '{col.header}'",
                    }
            schema_desc[table.key] = {
                "type": "list of objects",
                "row_fields": col_specs,
                "description": table.description or f"Data rows for table '{table.key}'",
            }

    prompt = f"""You are an expert, deterministic data extraction assistant.
Extract information from the provided PowerPoint presentation content according to the target schema.

=== PRESENTATION CONTENT ===
{deck_text}

=== TARGET SCHEMA REQUIREMENTS ===
Extract the data into a single JSON object matching this schema:
{json.dumps(schema_desc, indent=2)}

=== RULES FOR DETERMINISM & ACCURACY ===
1. Return ONLY a valid JSON object matching the requested keys.
2. Clean all numbers: convert formatted strings like "$1,200,000", "50k", or "12.5%" to pure numbers (e.g. 1200000, 50000, 0.125 or 12.5 according to context).
3. For table arrays, extract each individual data row into an object with the specified column keys. Do not include summary/total rows if they should be formulas.
4. If a field cannot be found, use null or 0.
5. Output pure JSON without markdown code fences or conversational text.
"""
    return prompt


def extract_data_with_llm(
    deck: DeckContent,
    config: MappingConfig,
    api_key: Optional[str] = None,
    model_name: str = "gemini-2.5-flash",
) -> Dict[str, Any]:
    """
    Calls Gemini API with temperature=0.0 for deterministic structured extraction.
    """
    final_api_key = api_key or os.getenv("GEMINI_API_KEY", "")

    if not final_api_key or "YOUR_GEMINI_API_KEY" in final_api_key:
        raise ValueError(
            "Gemini API key is not configured. Please set GEMINI_API_KEY in your .env file or pass it to the converter."
        )

    # Import google-genai
    try:
        from google import genai
        from google.genai import types

        client = genai.Client(api_key=final_api_key)
        prompt = build_extraction_prompt(deck, config)

        response = client.models.generate_content(
            model=model_name,
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0.0,
                response_mime_type="application/json",
            ),
        )

        response_text = response.text.strip()
        # Clean any accidental markdown wrap
        if response_text.startswith("```json"):
            response_text = response_text[7:]
        if response_text.startswith("```"):
            response_text = response_text[3:]
        if response_text.endswith("```"):
            response_text = response_text[:-3]

        return json.loads(response_text.strip())

    except Exception as e:
        logger.error(f"Error during Gemini LLM extraction: {e}")
        raise
