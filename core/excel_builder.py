import re
from typing import Dict, Any, Union, Optional
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import coordinate_to_tuple, get_column_letter

from core.models import MappingConfig, SheetMapping, StyleDef, CellMapping, TableMapping


def _clean_hex(hex_code: Optional[str]) -> Optional[str]:
    """Clean hex string to 6 or 8 character format for openpyxl."""
    if not hex_code:
        return None
    cleaned = hex_code.strip().lstrip("#")
    if len(cleaned) == 6:
        return f"FF{cleaned.upper()}"
    elif len(cleaned) == 8:
        return cleaned.upper()
    return cleaned.upper()


def apply_cell_style(cell, style_def: Optional[StyleDef]):
    """Applies font, fill, border, alignment, and number formatting to an openpyxl cell."""
    if not style_def:
        return

    font_kwargs: Dict[str, Any] = {
        "name": style_def.font or "Calibri",
        "size": style_def.size or 11,
        "bold": style_def.bold,
        "italic": style_def.italic,
    }
    hex_color = _clean_hex(style_def.font_color)
    if hex_color:
        font_kwargs["color"] = hex_color
    cell.font = Font(**font_kwargs)

    # Background Fill
    if style_def.bg_color:
        bg_hex = _clean_hex(style_def.bg_color)
        cell.fill = PatternFill(start_color=bg_hex, end_color=bg_hex, fill_type="solid")

    # Alignment
    cell.alignment = Alignment(
        horizontal=style_def.align or "left",
        vertical=style_def.valign or "center",
        wrap_text=True,
    )

    # Border
    if style_def.border:
        side = Side(style=style_def.border, color="FFD9D9D9")
        cell.border = Border(left=side, right=side, top=side, bottom=side)

    # Number Format
    if style_def.num_format:
        cell.number_format = style_def.num_format


def _cast_value(val: Any, val_type: str) -> Any:
    """Casts extracted values to proper numeric or string types."""
    if val is None:
        return ""
    if val_type in ("number", "float"):
        try:
            if isinstance(val, (int, float)):
                return float(val)
            clean_str = str(val).replace("$", "").replace(",", "").replace("%", "").strip()
            return float(clean_str)
        except (ValueError, TypeError):
            return val
    elif val_type in ("integer", "int"):
        try:
            if isinstance(val, (int, float)):
                return int(val)
            clean_str = str(val).replace(",", "").strip()
            return int(float(clean_str))
        except (ValueError, TypeError):
            return val
    return str(val)


def build_excel_workbook(config: MappingConfig, extracted_data: Dict[str, Any]) -> openpyxl.Workbook:
    """
    Deterministically constructs an Excel workbook based on MappingConfig and extracted data.
    """
    wb = openpyxl.Workbook()
    # Remove default sheet
    default_sheet = wb.active

    for sheet_idx, sheet_def in enumerate(config.sheets):
        if sheet_idx == 0 and default_sheet is not None:
            ws = default_sheet
            ws.title = sheet_def.name
        else:
            ws = wb.create_sheet(title=sheet_def.name)

        styles = sheet_def.styles

        # 1. Populate Individual Cells
        for cell_map in sheet_def.cells:
            target_row, target_col = coordinate_to_tuple(cell_map.cell)

            # Write label if defined
            if cell_map.label:
                if cell_map.label_cell:
                    lbl_row, lbl_col = coordinate_to_tuple(cell_map.label_cell)
                else:
                    # Place label directly to the left if col > 1
                    lbl_row, lbl_col = target_row, max(1, target_col - 1)
                lbl_cell = ws.cell(row=lbl_row, column=lbl_col, value=cell_map.label)
                if cell_map.label_style and cell_map.label_style in styles:
                    apply_cell_style(lbl_cell, styles[cell_map.label_style])

            # Write target value
            raw_val = extracted_data.get(cell_map.key)
            typed_val = _cast_value(raw_val, cell_map.type)
            data_cell = ws.cell(row=target_row, column=target_col, value=typed_val)

            if cell_map.style and cell_map.style in styles:
                apply_cell_style(data_cell, styles[cell_map.style])

        # 2. Populate Tables
        for table_map in sheet_def.tables:
            start_row, start_col = coordinate_to_tuple(table_map.start_cell)

            # Map column key to its column letter in Excel for formula resolution
            col_letter_map: Dict[str, str] = {}
            for col_idx, col_def in enumerate(table_map.columns):
                col_letter = get_column_letter(start_col + col_idx)
                col_letter_map[col_def.key] = col_letter

                # Set custom width if specified
                if col_def.width:
                    ws.column_dimensions[col_letter].width = col_def.width
                else:
                    # Default sensible width
                    current_w = ws.column_dimensions[col_letter].width
                    if not current_w or current_w < 15:
                        ws.column_dimensions[col_letter].width = 16

            # Write Table Headers
            for col_idx, col_def in enumerate(table_map.columns):
                h_cell = ws.cell(row=start_row, column=start_col + col_idx, value=col_def.header)
                h_style_name = col_def.header_style or "header"
                if h_style_name in styles:
                    apply_cell_style(h_cell, styles[h_style_name])

            # Write Data Rows
            table_rows = extracted_data.get(table_map.key, [])
            if not isinstance(table_rows, list):
                table_rows = []

            for row_offset, row_item in enumerate(table_rows):
                current_row = start_row + 1 + row_offset
                for col_idx, col_def in enumerate(table_map.columns):
                    cell = ws.cell(row=current_row, column=start_col + col_idx)

                    # Handle Formula Column
                    if col_def.type == "formula" and col_def.formula:
                        # Template e.g. "={planned}-{actual}" -> "=B7-C7"
                        formula_str = col_def.formula
                        for key, letter in col_letter_map.items():
                            formula_str = formula_str.replace(f"{{{key}}}", f"{letter}{current_row}")
                        cell.value = formula_str
                    else:
                        val = row_item.get(col_def.key) if isinstance(row_item, dict) else ""
                        cell.value = _cast_value(val, col_def.type)

                    # Apply data cell style
                    if col_def.style and col_def.style in styles:
                        apply_cell_style(cell, styles[col_def.style])

    return wb


def export_workbook_to_bytes(wb: openpyxl.Workbook) -> bytes:
    """Serializes an openpyxl Workbook into in-memory bytes."""
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()
