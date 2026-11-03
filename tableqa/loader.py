"""Load one business's files into raw tables, reusing the verified step 1 parsers.

Every row is kept with its source row number. Nothing is guessed: no header row,
no filled-down cells, no deduplication, no date or money parsing. CSV/XLSX values
stay strings as stored (Excel serials stay numbers-as-text), SQLite values keep
their SQLite types. Absent XLSX cells are missing (NaN); explicit empty cells are "".
"""

from dataclasses import dataclass, field
from pathlib import Path
import re
import xml.etree.ElementTree as ET

import pandas as pd

from ingestion.reader import NS, SUPPORTED, Scanner

# OOXML built-in number formats that display dates or times.
DATE_FORMAT_IDS = set(range(14, 23)) | {45, 46, 47}


def letters(column):
    text = ""
    while column:
        column, rem = divmod(column - 1, 26)
        text = chr(65 + rem) + text
    return text


def is_date_format(code):
    # Quoted literals and [colour]/[locale] sections are not date tokens.
    code = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", code or "")
    return bool(re.search(r"[dmyhs]", code, re.I)) and code.lower() != "general"


@dataclass
class Table:
    name: str
    source: str
    cells: pd.DataFrame
    date_cells: set = field(default_factory=set)
    hidden: bool = False
    merged: list = field(default_factory=list)


@dataclass
class Business:
    id: str
    readme: str
    tables: dict
    errors: list


class _Collector(Scanner):
    """Receives the reader's records in memory and keeps only compact rows."""

    def __init__(self):
        self.errors = []
        self.sheets = []
        self.strings = []
        self.date_styles = set()
        self.sql_columns = {}

    def error(self, entry, message, **location):
        self.errors.append({"source": entry["path"], "message": str(message), **location})

    def emit(self, entry, kind, **fields):
        if kind == "csv_row":
            self.sheets[-1]["rows"].append((fields["row"], dict(zip(map(letters, range(1, len(fields["values"]) + 1)), fields["values"]))))
        elif kind == "xlsx_styles":
            root = ET.fromstring(fields["xml"])
            custom = {int(f.get("numFmtId")): f.get("formatCode") for f in root.iter(NS + "numFmt")}
            xfs = root.find(NS + "cellXfs")
            for index, xf in enumerate([] if xfs is None else xfs):
                fmt = int(xf.get("numFmtId", 0))
                if fmt in DATE_FORMAT_IDS or (fmt in custom and is_date_format(custom[fmt])):
                    self.date_styles.add(str(index))
        elif kind == "xlsx_shared_string":
            self.strings.append(fields["text"])
        elif kind == "xlsx_sheet":
            self.sheets.append({"name": fields["sheet"], "hidden": fields["attributes"].get("state", "visible") != "visible",
                                "rows": [], "dates": set(), "merged": []})
        elif kind == "xlsx_row":
            sheet, values = self.sheets[-1], {}
            for cell in fields["cells"]:
                col = re.match(r"[A-Z]+", cell["coordinate"])[0]
                kind_attr = cell["attributes"].get("t")
                if cell["inline_string"] is not None:
                    values[col] = cell["inline_string"]["text"]
                elif kind_attr == "s":
                    values[col] = self.strings[int(cell["raw_value"])]
                else:
                    values[col] = "" if cell["raw_value"] is None else cell["raw_value"]
                    if kind_attr in (None, "n") and cell["attributes"].get("s") in self.date_styles and cell["raw_value"]:
                        sheet["dates"].add((fields["row"], col))
            sheet["rows"].append((fields["row"], values))
        elif kind == "xlsx_sheet_metadata" and fields["xml"].lstrip("<").split(" ")[0].endswith("mergeCells"):
            self.sheets[-1]["merged"] += [m.get("ref") for m in ET.fromstring(fields["xml"]).iter(NS + "mergeCell")]
        elif kind == "sqlite_table":
            # Hidden (virtual) columns are not returned by SELECT *.
            self.sql_columns[fields["table"]] = [c[1] for c in fields["columns"] if not c[6]]
            self.sheets.append({"name": fields["table"], "hidden": False, "rows": [], "dates": set(), "merged": []})
        elif kind == "sqlite_row":
            values = [v.get("value", v.get("base64", v.get("special"))) for v in fields["values"]]
            self.sheets[-1]["rows"].append((fields["ordinal"], dict(zip(self.sql_columns[fields["table"]], values))))


def _frame(rows, columns=None):
    frame = pd.DataFrame([values for _, values in rows], index=pd.Index([row for row, _ in rows], dtype="int64"), columns=columns)
    if columns is None:
        # Spreadsheet columns in sheet order (A, B, ..., Z, AA), not alphabetical.
        frame = frame[sorted(frame.columns, key=lambda c: (len(c), c))]
    frame.index.name = "row"
    return frame.astype(object)


def load_file(path, rel):
    """One CSV/XLSX/SQLite file -> ({table name: Table}, errors)."""
    entry = {"path": rel, "rows": 0, "cells": 0}
    collector = _Collector()
    suffix = Path(path).suffix.lower()
    if suffix == ".csv":
        collector.sheets.append({"name": None, "hidden": False, "rows": [], "dates": set(), "merged": []})
    try:
        # SQLite is opened by file URI, which needs an absolute path.
        getattr(collector, "scan_" + suffix[1:])(Path(path).resolve(), entry)
    except Exception as exc:
        collector.error(entry, f"{type(exc).__name__}: {exc}")
    tables = {}
    for sheet in collector.sheets:
        name = rel if sheet["name"] is None else f"{rel} / {sheet['name']}"
        columns = collector.sql_columns.get(sheet["name"]) if suffix == ".sqlite" else None
        tables[name] = Table(name, rel, _frame(sheet["rows"], columns), sheet["dates"], sheet["hidden"], sheet["merged"])
    return tables, collector.errors


def load_business(business_dir):
    business_dir = Path(business_dir)
    readme = business_dir / "README.md"
    tables, errors = {}, []
    for path in sorted(p for p in business_dir.rglob("*") if p.is_file() and p.suffix.lower() in SUPPORTED):
        found, problems = load_file(path, path.relative_to(business_dir).as_posix())
        tables.update(found)
        errors += problems
    return Business(business_dir.name, readme.read_text(encoding="utf-8-sig") if readme.exists() else "", tables, errors)


def load_rules(split_dir):
    """Split-level context read at run time: rulebook text and the raw calendar table."""
    split_dir = Path(split_dir)
    calendar, _ = load_file(split_dir / "calendar.csv", "calendar.csv")
    return (split_dir / "rulebook.md").read_text(encoding="utf-8-sig"), calendar["calendar.csv"]
