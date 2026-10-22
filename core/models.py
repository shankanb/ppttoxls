from typing import List, Dict, Optional, Any, Union
from pydantic import BaseModel, Field


class StyleDef(BaseModel):
    font: str = "Calibri"
    size: int = 11
    bold: bool = False
    italic: bool = False
    font_color: Optional[str] = None  # e.g., "#FFFFFF" or "FFFFFF"
    bg_color: Optional[str] = None    # e.g., "#2F5597" or "2F5597"
    align: str = "left"               # "left", "center", "right"
    valign: str = "center"            # "top", "center", "bottom"
    border: Optional[str] = None      # "thin", "medium", "double"
    num_format: Optional[str] = None  # e.g., "$#,##0", "0.0%", "#,##0.00"


class CellMapping(BaseModel):
    cell: str                          # e.g., "B2"
    key: str                           # key for LLM output, e.g. "total_budget"
    type: str = "string"               # "string", "number", "integer"
    description: str = ""              # LLM extraction context
    style: Optional[str] = None        # Style key in styles dict
    label: Optional[str] = None        # Optional text label to place (e.g. at A2)
    label_cell: Optional[str] = None   # Target coordinate for label if specified
    label_style: Optional[str] = None  # Style key for label


class ColumnMapping(BaseModel):
    header: str                        # Column title, e.g. "Department"
    key: str                           # Key in extracted table row, e.g. "dept_name"
    type: str = "string"               # "string", "number", "integer", "formula"
    description: str = ""              # Guidance for LLM
    style: Optional[str] = None        # Style for data cells
    header_style: Optional[str] = None # Style for header cell
    formula: Optional[str] = None      # Formula template, e.g. "={planned}-{actual}"
    width: Optional[int] = None        # Column width in Excel characters


class TableMapping(BaseModel):
    key: str                           # Key for LLM output list, e.g. "departments"
    start_cell: str = "A5"             # Header start cell coordinate, e.g. "A5"
    description: str = ""              # Context for LLM
    columns: List[ColumnMapping] = []


class SheetMapping(BaseModel):
    name: str = "Sheet1"
    styles: Dict[str, StyleDef] = {}
    cells: List[CellMapping] = []
    tables: List[TableMapping] = []


class MappingConfig(BaseModel):
    title: Optional[str] = "PPT to XLS Export"
    sheets: List[SheetMapping]
