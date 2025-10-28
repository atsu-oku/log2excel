"""
Minimal XLSX writer that creates worksheets using inline strings only.

This module avoids third-party dependencies by building the required Open XML
structures manually.  It supports basic text content without styling.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, List
from xml.sax.saxutils import escape
import io
import zipfile


@dataclass
class CellValue:
    value: str = ""
    formula: str | None = None
    data_type: str | None = None


@dataclass
class SheetData:
    name: str
    rows: List[List[CellValue]]


def _column_letter(index: int) -> str:
    """Convert a 1-based column index to the Excel column letter."""
    if index < 1:
        raise ValueError("Column index must be positive")
    letters = []
    while index:
        index, remainder = divmod(index - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def _build_sheet_xml(sheet: SheetData) -> bytes:
    max_cols = max((len(row) for row in sheet.rows), default=0)
    if max_cols == 0:
        max_cols = 1

    lines: List[str] = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        "  <sheetData>",
    ]

    for row_index, row in enumerate(sheet.rows, start=1):
        has_values = any(
            (cell.value or cell.formula)
            for cell in row
        )
        if not has_values:
            lines.append(f'    <row r="{row_index}" spans="1:{max_cols}"/>')
            continue

        lines.append(f'    <row r="{row_index}" spans="1:{max_cols}">')
        for col_index, cell in enumerate(row, start=1):
            if not cell.value and not cell.formula:
                continue
            cell_ref = f"{_column_letter(col_index)}{row_index}"
            if cell.formula:
                attrs = [f'r="{cell_ref}"']
                if cell.data_type:
                    attrs.append(f't="{cell.data_type}"')
                lines.append(f"      <c {' '.join(attrs)}>")
                lines.append(f"        <f>{escape(cell.formula)}</f>")
                if cell.value:
                    lines.append(f"        <v>{escape(cell.value)}</v>")
                lines.append("      </c>")
            else:
                value = cell.value
                escaped = escape(value).replace("\n", "&#10;")
                preserve = ' xml:space="preserve"' if value.strip() != value else ""
                lines.append(
                    f'      <c r="{cell_ref}" t="inlineStr"><is><t{preserve}>{escaped}</t></is></c>'
                )
        lines.append("    </row>")

    lines.append("  </sheetData>")
    lines.append("</worksheet>")
    return "\n".join(lines).encode("utf-8")


def _build_workbook_xml(sheet_names: Iterable[str]) -> bytes:
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        "  <sheets>",
    ]
    for index, name in enumerate(sheet_names, start=1):
        safe_name = escape(name)
        lines.append(
            f'    <sheet name="{safe_name}" sheetId="{index}" r:id="rId{index}"/>'
        )
    lines.append("  </sheets>")
    lines.append("</workbook>")
    return "\n".join(lines).encode("utf-8")


def _build_content_types_xml(sheet_count: int) -> bytes:
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '  <Default Extension="xml" ContentType="application/xml"/>',
        '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    ]
    for index in range(1, sheet_count + 1):
        lines.append(
            f'  <Override PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    lines.append(
        '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    )
    lines.append("</Types>")
    return "\n".join(lines).encode("utf-8")


def _build_root_rels_xml() -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>\n'
        "</Relationships>\n"
    ).encode("utf-8")


def _build_workbook_rels_xml(sheet_count: int) -> bytes:
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ]
    for index in range(1, sheet_count + 1):
        lines.append(
            f'  <Relationship Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )
    lines.append(
        '  <Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'.format(
            sheet_count + 1
        )
    )
    lines.append("</Relationships>")
    return "\n".join(lines).encode("utf-8")


def _build_styles_xml() -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'
        '  <fonts count="1">\n'
        "    <font>\n"
        '      <sz val="11"/>\n'
        '      <color theme="1"/>\n'
        '      <name val="Calibri"/>\n'
        '      <family val="2"/>\n'
        "    </font>\n"
        "  </fonts>\n"
        '  <fills count="2">\n'
        "    <fill><patternFill patternType=\"none\"/></fill>\n"
        "    <fill><patternFill patternType=\"gray125\"/></fill>\n"
        "  </fills>\n"
        '  <borders count="1">\n'
        "    <border><left/><right/><top/><bottom/><diagonal/></border>\n"
        "  </borders>\n"
        '  <cellStyleXfs count="1">\n'
        '    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>\n'
        "  </cellStyleXfs>\n"
        '  <cellXfs count="1">\n'
        '    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>\n'
        "  </cellXfs>\n"
        '  <cellStyles count="1">\n'
        '    <cellStyle name="Normal" xfId="0" builtinId="0"/>\n'
        "  </cellStyles>\n"
        "</styleSheet>\n"
    ).encode("utf-8")


def write_xlsx(path: str | io.BufferedIOBase, sheets: Iterable[SheetData]) -> None:
    sheets = list(sheets)
    if not sheets:
        raise ValueError("Workbook must contain at least one sheet")

    sheet_xml_data = [_build_sheet_xml(sheet) for sheet in sheets]
    workbook_xml = _build_workbook_xml(sheet.name for sheet in sheets)
    content_types_xml = _build_content_types_xml(len(sheets))
    root_rels_xml = _build_root_rels_xml()
    workbook_rels_xml = _build_workbook_rels_xml(len(sheets))
    styles_xml = _build_styles_xml()

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", root_rels_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        zf.writestr("xl/styles.xml", styles_xml)
        for index, data in enumerate(sheet_xml_data, start=1):
            zf.writestr(f"xl/worksheets/sheet{index}.xml", data)
