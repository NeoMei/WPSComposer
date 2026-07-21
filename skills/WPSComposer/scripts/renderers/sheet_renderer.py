"""Sheet renderer — StructuredDocument to WPS Spreadsheets to XLSX.

Each H2 section with a table becomes a separate worksheet.
Tables without a heading go to a default "Data" sheet.
"""

from __future__ import annotations

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
                # Sanitize sheet name (31 char limit)
                sheet_name = sheet_name[:31]

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
                data = [elem.headers] + elem.rows
                s.write_table(
                    start_row=1, start_col=1,
                    data=data,
                    header_shade=True,
                    header_font_color="#FFFFFF",
                )
                s.set_column_width("A", 15)
                if len(elem.headers) > 1:
                    for c in range(2, len(elem.headers) + 1):
                        col_letter = chr(64 + c) if c <= 26 else "Z"
                        s.set_column_width(col_letter, 12)

                table_idx += 1

        if sheet_index > 1:
            s.select_sheet(1)

        s.autofit()
        return s.save_xlsx(output_path)
