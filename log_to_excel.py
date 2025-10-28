"""
Compare STG and PRD log files and export the differences to an XLSX workbook.

Usage example:

    python log_to_excel.py --input-dir ref --output comparison.xlsx
"""

from __future__ import annotations

import argparse
import difflib
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Set, Tuple

try:
    from .xlsx_writer import CellValue, SheetData, write_xlsx
except ImportError:
    from xlsx_writer import CellValue, SheetData, write_xlsx


@dataclass
class LogPair:
    """Metadata describing the staged/production log files that should be compared."""
    base_name: str
    stg_host: str
    prd_host: str
    stg_path: Path
    prd_path: Path


def _extract_host_name(path: Path) -> str:
    stem = path.stem
    if "diff_" in stem:
        stem = stem.split("diff_", 1)[1]
    return stem


def _discover_pairs(input_dir: Path) -> List[LogPair]:
    stg_files: Dict[str, Tuple[str, Path]] = {}
    prd_files: Dict[str, Tuple[str, Path]] = {}

    for entry in input_dir.glob("*.log"):
        host = _extract_host_name(entry)
        if not host:
            continue
        suffix = host[-1].lower()
        base = host[:-1]
        if suffix == "s":
            stg_files[base] = (host, entry)
        elif suffix == "p":
            prd_files[base] = (host, entry)

    pairs: List[LogPair] = []
    for base in sorted(set(stg_files) & set(prd_files)):
        stg_host, stg_path = stg_files[base]
        prd_host, prd_path = prd_files[base]
        pairs.append(LogPair(base, stg_host, prd_host, stg_path, prd_path))
    return pairs


def _read_log_lines(path: Path) -> List[str]:
    with path.open("r", encoding="utf-8") as fh:
        return [line.rstrip("\r\n") for line in fh]


def _align_logs(
    stg_lines: List[str], prd_lines: List[str]
) -> List[Tuple[str, str, int | None, int | None]]:
    """Align two sequences of log lines using difflib.SequenceMatcher."""
    matcher = difflib.SequenceMatcher(a=stg_lines, b=prd_lines, autojunk=False)
    aligned: List[Tuple[str, str, int | None, int | None]] = []

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for offset, (s_idx, p_idx) in enumerate(zip(range(i1, i2), range(j1, j2))):
                aligned.append(
                    (
                        stg_lines[s_idx],
                        prd_lines[p_idx],
                        s_idx + 1,
                        p_idx + 1,
                    )
                )
        elif tag == "replace":
            span = max(i2 - i1, j2 - j1)
            for offset in range(span):
                s_idx = i1 + offset
                p_idx = j1 + offset
                stg_line = stg_lines[s_idx] if s_idx < i2 else ""
                prd_line = prd_lines[p_idx] if p_idx < j2 else ""
                stg_no = s_idx + 1 if s_idx < i2 else None
                prd_no = p_idx + 1 if p_idx < j2 else None
                aligned.append((stg_line, prd_line, stg_no, prd_no))
        elif tag == "delete":
            for s_idx in range(i1, i2):
                aligned.append((stg_lines[s_idx], "", s_idx + 1, None))
        elif tag == "insert":
            for p_idx in range(j1, j2):
                aligned.append(("", prd_lines[p_idx], None, p_idx + 1))

    return aligned


SHEET_NAME_FORMULA = 'RIGHT(CELL("filename",A1),LEN(CELL("filename",A1))-FIND("]",CELL("filename",A1)))'


def _build_header_rows(pair: LogPair) -> List[List[CellValue]]:
    return [
        [
            CellValue(value="InfoOne延命プロジェクト"),
            CellValue(),
            CellValue(),
            CellValue(),
            CellValue(),
        ],
        [
            CellValue(value="オンプレ単体テスト（NewSTG基盤）"),
            CellValue(),
            CellValue(),
            CellValue(),
            CellValue(),
        ],
        [CellValue(), CellValue(), CellValue(), CellValue(), CellValue()],
        [
            CellValue(),
            CellValue(value="対象サーバ"),
            CellValue(value="備考"),
            CellValue(),
            CellValue(),
        ],
        [
            CellValue(),
            CellValue(value="", formula=SHEET_NAME_FORMULA, data_type="str"),
            CellValue(),
            CellValue(),
            CellValue(),
        ],
        [CellValue(), CellValue(), CellValue(), CellValue(), CellValue()],
        [
            CellValue(),
            CellValue(value=f"現行サーバ（{pair.stg_host}）"),
            CellValue(value="差分有無"),
            CellValue(value=f"新基盤（{pair.prd_host}）"),
            CellValue(value="備考"),
        ],
    ]


def _build_rows_for_pair(pair: LogPair) -> List[List[CellValue]]:
    stg_lines = _read_log_lines(pair.stg_path)
    prd_lines = _read_log_lines(pair.prd_path)
    aligned = _align_logs(stg_lines, prd_lines)

    rows: List[List[CellValue]] = _build_header_rows(pair)
    current_row = len(rows) + 1

    for stg_line, prd_line, *_unused_indices in aligned:
        match_flag = "1" if stg_line == prd_line else "0"
        note = "" if match_flag == "1" else "差異あり"
        rows.append(
            [
                CellValue(),
                CellValue(value=stg_line),
                CellValue(
                    value=match_flag,
                    formula=f"B{current_row}=D{current_row}",
                    data_type="b",
                ),
                CellValue(value=prd_line),
                CellValue(value=note),
            ]
        )
        current_row += 1
    return rows


def _column_index_from_ref(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    index = 0
    for char in letters.upper():
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index


def _load_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    strings: List[str] = []
    for si in root.findall("s:si", ns):
        text = "".join(t.text or "" for t in si.findall(".//s:t", ns))
        strings.append(text)
    return strings


def _parse_sheet_rows(xml_bytes: bytes, shared_strings: List[str]) -> List[List[CellValue]]:
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    root = ET.fromstring(xml_bytes)
    rows: List[List[CellValue]] = []
    for row in root.findall("s:sheetData/s:row", ns):
        cells: Dict[int, CellValue] = {}
        max_col = 0
        for cell in row.findall("s:c", ns):
            cell_ref = cell.get("r")
            if not cell_ref:
                continue
            col_index = _column_index_from_ref(cell_ref)
            data_type = cell.get("t")
            formula_node = cell.find("s:f", ns)
            formula = formula_node.text if formula_node is not None else None
            value_node = cell.find("s:v", ns)
            text = ""
            if data_type == "s" and value_node is not None:
                idx = int(value_node.text or "0")
                text = shared_strings[idx] if 0 <= idx < len(shared_strings) else ""
            elif data_type == "inlineStr":
                t_node = cell.find(".//s:t", ns)
                text = "" if t_node is None else t_node.text or ""
            elif value_node is not None:
                text = value_node.text or ""
            cells[col_index] = CellValue(value=text, formula=formula, data_type=data_type)
            max_col = max(max_col, col_index)
        if max_col == 0:
            rows.append([])
            continue
        row_values = [CellValue() for _ in range(max_col)]
        for col_index, cell_value in cells.items():
            row_values[col_index - 1] = cell_value
        rows.append(row_values)
    return rows


def _load_existing_sheets(output_path: Path) -> List[SheetData]:
    ns = {
        "w": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }
    sheets: List[SheetData] = []
    with zipfile.ZipFile(output_path) as zf:
        shared_strings = _load_shared_strings(zf)
        wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
        rel_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rel_map = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rel_root.findall(
                "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
            )
        }
        for sheet in wb_root.findall("w:sheets/w:sheet", ns):
            rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            name = sheet.attrib["name"]
            target = rel_map.get(rid)
            if not target:
                continue
            rows = _parse_sheet_rows(zf.read(f"xl/{target}"), shared_strings)
            sheets.append(SheetData(name=name, rows=rows))
    return sheets


def _make_unique_sheet_name(base_name: str, used_names: Set[str]) -> str:
    if base_name not in used_names:
        used_names.add(base_name)
        return base_name
    counter = 2
    while True:
        suffix = f".v{counter}"
        max_len = 31 - len(suffix)
        candidate = base_name[:max_len] + suffix
        if candidate not in used_names:
            used_names.add(candidate)
            return candidate
        counter += 1


def generate_workbook(input_dir: Path, output_path: Path) -> None:
    """Create or extend an XLSX workbook with STG/PRD comparisons discovered under input_dir."""
    pairs = _discover_pairs(input_dir)
    if not pairs:
        raise FileNotFoundError(
            f"No STG/PRD log pairs were discovered under '{input_dir}'."
        )

    existing_sheets: List[SheetData] = []
    used_names: Set[str] = set()
    if output_path.exists():
        existing_sheets = _load_existing_sheets(output_path)
        used_names.update(sheet.name for sheet in existing_sheets)

    new_sheets: List[SheetData] = []
    for pair in pairs:
        rows = _build_rows_for_pair(pair)
        base_name = pair.base_name[:31]
        sheet_name = _make_unique_sheet_name(base_name, used_names)
        new_sheets.append(SheetData(name=sheet_name, rows=rows))

    write_xlsx(str(output_path), existing_sheets + new_sheets)


def parse_args() -> argparse.Namespace:
    """Return the CLI arguments for building the workbook."""
    parser = argparse.ArgumentParser(description="Compare STG/PRD logs and export XLSX.")
    parser.add_argument(
        "--input-dir",
        default=".",
        type=Path,
        help="Directory containing STG/PRD log files (default: current directory).",
    )
    parser.add_argument(
        "--output",
        default="comparison.xlsx",
        type=Path,
        help="Path to the XLSX file to generate (default: comparison.xlsx).",
    )
    return parser.parse_args()


def main() -> None:
    """CLI entry point for generating the comparison workbook."""
    args = parse_args()
    input_dir: Path = args.input_dir
    output_path: Path = args.output

    if not input_dir.is_dir():
        raise NotADirectoryError(f"Input directory '{input_dir}' does not exist.")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    generate_workbook(input_dir, output_path)
    print(f"Workbook generated: {output_path.resolve()}")


if __name__ == "__main__":
    main()
