"""
Utility to verify that every log line from STG/PRD inputs appears somewhere in the
generated workbook output.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Dict, List, Set
import xml.etree.ElementTree as ET
import zipfile

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

# pylint: disable=wrong-import-position
from log_to_excel import _discover_pairs, _read_log_lines  # type: ignore
# pylint: enable=wrong-import-position


def _load_sheet_texts(xlsx_path: Path) -> Dict[str, List[str]]:
    """Return the textual contents of every cell for each sheet in the workbook."""
    with zipfile.ZipFile(xlsx_path) as zf:
        workbook_xml = ET.fromstring(zf.read("xl/workbook.xml"))
        ns = {
            "w": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        }
        sheets = []
        for sheet in workbook_xml.findall("w:sheets/w:sheet", ns):
            rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            name = sheet.attrib["name"]
            sheets.append((rid, name))

        rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rels = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels_root.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
        }

        sheet_texts: Dict[str, List[str]] = {}
        for rid, name in sheets:
            sheet_target = rels[rid]
            sheet_xml = ET.fromstring(zf.read(f"xl/{sheet_target}"))
            ns_sheet = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            texts: List[str] = []
            for t in sheet_xml.findall(".//s:t", ns_sheet):
                texts.append(t.text or "")
            sheet_texts[name] = texts

    return sheet_texts


def verify_logs_against_workbook(input_dir: Path, workbook_path: Path) -> bool:
    """Confirm every line from STG/PRD logs appears somewhere in the workbook text."""
    pairs = _discover_pairs(input_dir)
    sheet_texts = _load_sheet_texts(workbook_path)

    all_ok = True
    for pair in pairs:
        sheet_name = pair.base_name[:31]
        sheet_values = sheet_texts.get(sheet_name, [])
        value_set: Set[str] = set(sheet_values)

        for line in _read_log_lines(pair.stg_path):
            if line and line not in value_set:
                print(f"[STG MISSING] {pair.stg_host}: {line}")
                all_ok = False
        for line in _read_log_lines(pair.prd_path):
            if line and line not in value_set:
                print(f"[PRD MISSING] {pair.prd_host}: {line}")
                all_ok = False

    return all_ok


if __name__ == "__main__":
    default_input_dir = Path("ref")
    default_workbook = Path("comparison.xlsx")
    verification_passed = verify_logs_against_workbook(default_input_dir, default_workbook)
    print("Verification result:", "PASSED" if verification_passed else "FAILED")
