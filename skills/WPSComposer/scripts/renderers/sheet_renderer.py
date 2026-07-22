"""Sheet renderer — StructuredDocument to WPS Spreadsheets to XLSX.

Each H2 section with a table becomes a separate worksheet.
Tables without a heading go to a default "Data" sheet.
"""

from __future__ import annotations

import re

from ..document_model import (
    StructuredDocument, Section, TableBlock,
)
from ..sheet import SheetComposer


def render(doc: StructuredDocument, output_path: str,
           preset=None, composer_factory=SheetComposer) -> str:
    """Render StructuredDocument tables to a styled XLSX file.

    Args:
        doc: StructuredDocument.
        output_path: Path for the output .xlsx file.
        preset: Optional DesignPreset for header colour.

    Returns:
        Absolute path to the generated file.
    """
    tables = doc.all_tables
    if not tables:
        raise ValueError("No tables found in document. Nothing to render to XLSX.")

    with composer_factory() as s:
        header_color = preset.get_color("primary") if preset else "#4472C4"

        # Map each table section to a sheet
        sheet_index = 0
        table_idx = 0

        for sec_idx, section in enumerate(doc.sections):
            for elem in section.elements:
                if not isinstance(elem, TableBlock):
                    continue

                # Create/select sheet
                sheet_name = section.heading if section.has_heading else f"Table{table_idx + 1}"
                # Excel sheet names: 31 chars max, no []:*?/\
                sheet_name = re.sub(r"[\[\]:*?/\\]", " ", sheet_name).strip()[:31]
                sheet_name = sheet_name or f"Table{table_idx + 1}"

                if sheet_index == 0:
                    # Rename default sheet
                    try:
                        s.rename_sheet(1, sheet_name)
                    except Exception:
                        pass
                else:
                    s.add_sheet(sheet_name)

                sheet_index += 1

                # Write table data
                data = ([elem.headers] if elem.headers else []) + elem.rows
                s.write_table(
                    start_row=1, start_col=1,
                    data=data,
                    header_shade=header_color,
                    header_font_color="#FFFFFF",
                )

                table_idx += 1

        if sheet_index > 1:
            s.select_sheet(1)

        s.autofit()
        return s.save_xlsx(output_path)
