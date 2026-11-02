"""Stream source records without interpreting headers, dates, or business facts."""

import argparse
import base64
import csv
from contextlib import closing, contextmanager
import datetime as dt
import hashlib
import gzip
import io
import json
import math
import os
from pathlib import Path
import posixpath
import re
import sqlite3
import stat
import sys
import xml.etree.ElementTree as ET
import zipfile

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
SUPPORTED = {".csv", ".xlsx", ".sqlite"}


@contextmanager
def record_writer(path):
    with path.open("wb") as raw:
        with gzip.GzipFile(filename="", fileobj=raw, mode="wb", compresslevel=1, mtime=0) as compressed:
            with io.TextIOWrapper(compressed, encoding="utf-8", newline="\n") as text:
                yield text


def dump(handle, obj):
    handle.write(json.dumps(obj, ensure_ascii=False, allow_nan=False, separators=(",", ":")) + "\n")


def digest(path):
    h = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def reject_links(path):
    """Reject links/reparse points in every existing component (including junctions)."""
    for part in (path, *path.parents):
        try:
            info = part.lstat()
        except FileNotFoundError:
            continue
        if stat.S_ISLNK(info.st_mode) or getattr(info, "st_file_attributes", 0) & 0x400:
            raise ValueError(f"symlink/reparse path is not allowed: {part}")


def within(path, parent):
    return path == parent or parent in path.parents


def quote(name):
    return '"' + name.replace('"', '""') + '"'


def xlsx_member(relation):
    if relation.get("TargetMode") == "External":
        raise ValueError("external workbook part is unsupported")
    target = relation.get("Target", "")
    member = posixpath.normpath(target.lstrip("/") if target.startswith("/") else "xl/" + target)
    if not member.startswith("xl/"):
        raise ValueError("workbook relationship escapes xl/")
    return member


def string_text(element):
    # Phonetic rPh annotations are metadata, not part of the visible cell text.
    return "".join(child.text or "" if child.tag == NS + "t" else
                   (child.findtext(NS + "t") or "") if child.tag == NS + "r" else ""
                   for child in element)


class Scanner:
    def __init__(self, split, output):
        self.split = Path(os.path.abspath(split))
        self.output = Path(os.path.abspath(output))
        reject_links(self.split)
        reject_links(self.output)
        self.split = self.split.resolve(strict=True)
        self.output = self.output.resolve()
        if within(self.output, self.split) or within(self.split, self.output):
            raise ValueError("output must be outside split and cannot be its ancestor")
        if self.output.exists() and (not self.output.is_dir() or any(self.output.iterdir())):
            raise ValueError("output directory must be new or empty")
        self.output.mkdir(parents=True, exist_ok=True)
        self.manifest = {"contract_version": 1, "split": str(self.split),
                         "status": "incomplete", "files": [], "error_count": 0,
                         "errors_file": "errors.jsonl",
                         "records_file": "records.jsonl.gz",
                         "excluded": [], "businesses": [],
                         "ordering": "paths lexical; XLSX workbook order; SQL rows unordered"}
        self.records = None
        self.errors = None

    def error(self, entry, message, **location):
        err = {"source": entry["path"], "message": str(message), **location}
        self.manifest["error_count"] += 1
        entry["error_count"] += 1
        entry["status"] = "incomplete"
        dump(self.errors, err)

    def emit(self, entry, kind, **fields):
        dump(self.records, {"source": entry["path"], "kind": kind, **fields})
        entry["records"] += 1

    def discover(self):
        paths = []
        # os.walk never follows links. Reject linked directories rather than quietly skipping.
        def discovery_error(exc):
            paths.append((Path(exc.filename), "unsafe", f"directory discovery failed: {exc}"))
        for root, dirs, files in os.walk(self.split, followlinks=False, onerror=discovery_error):
            dirs.sort()
            if Path(root) == self.split / "data":
                self.manifest["businesses"] = list(dirs)
            for name in list(dirs):
                child = Path(root) / name
                try:
                    reject_links(child)
                except ValueError as exc:
                    dirs.remove(name)
                    paths.append((child, "unsafe", str(exc)))
            for name in sorted(files):
                child = Path(root) / name
                rel = child.relative_to(self.split).as_posix()
                if name.lower() == "answers.jsonl":
                    self.manifest["excluded"].append({"path": rel, "reason": "answer labels; not opened or hashed"})
                    continue
                kind = "unsupported"
                if rel in {"rulebook.md", "calendar.csv", "questions.jsonl"}:
                    kind = {"rulebook.md": "text", "calendar.csv": "csv", "questions.jsonl": "questions"}[rel]
                elif len(child.relative_to(self.split).parts) >= 3 and child.relative_to(self.split).parts[0] == "data":
                    if name == "README.md":
                        kind = "text"
                    elif child.suffix.lower() in SUPPORTED:
                        kind = child.suffix.lower()[1:]
                paths.append((child, kind, None))
        # All text and CSV are scanned before detailed workbook/database access.
        rank = {"text": 0, "csv": 1, "questions": 2, "xlsx": 3, "sqlite": 4}
        return sorted(paths, key=lambda p: (rank.get(p[1], 5), p[0].relative_to(self.split).as_posix()))

    def scan_csv(self, path, entry):
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            sample = handle.read(65536)
            handle.seek(0)
            # Sniffer handles quoting; fallback only for a genuinely single-column file.
            try:
                dialect = csv.Sniffer().sniff(sample.replace("\r\n", "\n").replace("\r", "\n"), delimiters=",;\t|")
                delimiter = dialect.delimiter
            except csv.Error:
                if any(c in sample for c in ",;\t|"):
                    raise ValueError("CSV delimiter could not be validated")
                delimiter = ","
                dialect = csv.excel
            with path.open("rb") as binary:
                entry["encoding"] = "utf-8-sig" if binary.read(3) == b"\xef\xbb\xbf" else "utf-8"
            entry["delimiter"] = delimiter
            entry["dialect"] = {key: getattr(dialect, key) for key in
                                ("quotechar", "doublequote", "skipinitialspace", "quoting")}
            raw_lines = []
            def lines():
                for line in handle:
                    raw_lines.append(line)
                    yield line
            reader = csv.reader(lines(), dialect=dialect, delimiter=delimiter, strict=True)
            width = None
            previous_line = 0
            for ordinal, values in enumerate(reader, 1):
                if width is None:
                    width = len(values)
                self.emit(entry, "csv_row", row=ordinal, line_start=previous_line + 1,
                          line_end=reader.line_num, values=values, raw_record="".join(raw_lines))
                raw_lines.clear()
                previous_line = reader.line_num
                entry["rows"] += 1
                entry["cells"] += len(values)
                if len(values) != width:
                    self.error(entry, f"CSV row width {len(values)} differs from first row {width}", row=ordinal)
            entry["columns"] = width or 0

    def scan_questions(self, path, entry):
        # Disk-backed duplicate detection keeps question count from growing memory usage.
        with closing(sqlite3.connect(self.output / "_question_ids.sqlite")) as ids:
            ids.execute("CREATE TABLE ids (id TEXT PRIMARY KEY)")
            with path.open("r", encoding="utf-8-sig", newline="") as handle:
                for line_no, line in enumerate(handle, 1):
                    entry["rows"] += 1
                    try:
                        def pairs(items):
                            obj = {}
                            for key, value in items:
                                if key in obj:
                                    raise ValueError(f"duplicate JSON key: {key}")
                                obj[key] = value
                            return obj
                        obj = json.loads(line, object_pairs_hook=pairs,
                                         parse_constant=lambda value: (_ for _ in ()).throw(ValueError(value)))
                        if not isinstance(obj, dict):
                            raise ValueError("question must be an object")
                        required = {"id", "business_id", "question", "as_of_date"}
                        if set(obj) != required:
                            raise ValueError("question fields must be exactly id, business_id, question, as_of_date (labels prohibited)")
                        if any(not isinstance(obj[k], str) or not obj[k].strip() for k in required):
                            raise ValueError("required question fields must be nonempty strings")
                        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", obj["as_of_date"]):
                            raise ValueError("as_of_date must be YYYY-MM-DD")
                        dt.date.fromisoformat(obj["as_of_date"])
                        if obj["business_id"] not in self.manifest["businesses"]:
                            raise ValueError("question references unknown business_id")
                        ids.execute("INSERT INTO ids VALUES (?)", (obj["id"],))
                        self.emit(entry, "question", line=line_no, value=obj)
                        entry["valid_questions"] = entry.get("valid_questions", 0) + 1
                    except (ValueError, sqlite3.IntegrityError) as exc:
                        self.error(entry, f"invalid question: {exc}", line=line_no)
        (self.output / "_question_ids.sqlite").unlink()

    def scan_xlsx(self, path, entry):
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
            if len(names) != len(set(names)):
                raise ValueError("duplicate ZIP member names")
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            rels = {item.get("Id"): item for item in relationships}
            self.emit(entry, "xlsx_workbook", xml=ET.tostring(workbook, encoding="unicode"))
            # Styles retain number formats and xf references without converting Excel serials.
            parts = {r.get("Type", "").rsplit("/", 1)[-1]: r for r in relationships}
            if "styles" in parts:
                self.emit(entry, "xlsx_styles", xml=archive.read(xlsx_member(parts["styles"])).decode("utf-8"))
            entry["shared_strings"] = 0
            if "sharedStrings" in parts:
                with archive.open(xlsx_member(parts["sharedStrings"])) as source:
                    context = ET.iterparse(source, events=("start", "end"))
                    _, root = next(context)
                    index = 0
                    for event, elem in context:
                        if event == "end" and elem.tag == NS + "si":
                            self.emit(entry, "xlsx_shared_string", index=index,
                                      text=string_text(elem),
                                      xml=ET.tostring(elem, encoding="unicode"))
                            index += 1
                            root.clear()
                    entry["shared_strings"] = index
            entry["sheets"] = []
            sheets = workbook.find(NS + "sheets")
            if sheets is None:
                raise ValueError("workbook has no sheets element")
            for sheet in sheets:
                meta = {"name": sheet.get("name"), "state": sheet.get("state", "visible"),
                        "sheet_id": sheet.get("sheetId"), "rows": 0, "cells": 0,
                        "max_row": 0, "max_column": 0, "status": "complete"}
                entry["sheets"].append(meta)
                try:
                    relation = rels[sheet.get(REL + "id")]
                    if not relation.get("Type", "").endswith("/worksheet"):
                        raise ValueError("sheet is not an internal worksheet")
                    member = xlsx_member(relation)
                    meta["member"] = member
                    self.emit(entry, "xlsx_sheet", sheet=meta["name"], attributes=dict(sheet.attrib), member=member)
                    self.scan_sheet(archive, member, entry, meta)
                except Exception as exc:
                    meta["status"] = "incomplete"
                    self.error(entry, f"{type(exc).__name__}: {exc}", sheet=meta["name"])

    def scan_sheet(self, archive, member, entry, meta):
        with archive.open(member) as source:
            context = ET.iterparse(source, events=("start", "end"))
            stack = []
            for event, elem in context:
                if event == "start":
                    stack.append(elem)
                    continue
                if elem.tag == NS + "row":
                    row = int(elem.get("r", meta["max_row"] + 1))
                    if row <= meta["max_row"]:
                        raise ValueError("non-increasing worksheet row coordinates")
                    cells = []
                    for cell in elem.findall(NS + "c"):
                        coord = cell.get("r", "")
                        match = re.fullmatch(r"([A-Z]+)([1-9]\d*)", coord)
                        if not match or int(match[2]) != row:
                            raise ValueError(f"invalid cell coordinate: {coord}")
                        col = 0
                        for letter in match[1]:
                            col = col * 26 + ord(letter) - 64
                        if col <= (cells[-1]["column"] if cells else 0):
                            raise ValueError("non-increasing cell coordinates")
                        value = cell.find(NS + "v")
                        if cell.get("t") == "s" and (value is None or value.text is None or
                                not value.text.isdigit() or int(value.text) >= entry["shared_strings"]):
                            raise ValueError(f"invalid shared string reference at {coord}")
                        formula = cell.find(NS + "f")
                        inline = cell.find(NS + "is")
                        cells.append({"coordinate": coord, "column": col, "attributes": dict(cell.attrib),
                                      "value_present": value is not None,
                                      "raw_value": value.text if value is not None else None,
                                      "formula": None if formula is None else {"text": formula.text, "attributes": dict(formula.attrib)},
                                      "inline_string": None if inline is None else {
                                          "text": string_text(inline),
                                          "xml": ET.tostring(inline, encoding="unicode")}})
                        meta["max_column"] = max(meta["max_column"], col)
                    self.emit(entry, "xlsx_row", sheet=meta["name"], row=row,
                              attributes=dict(elem.attrib), cells=cells)
                    meta["rows"] += 1
                    meta["cells"] += len(cells)
                    meta["max_row"] = row
                    entry["rows"] += 1
                    entry["cells"] += len(cells)
                    if len(stack) > 1:
                        stack[-2].remove(elem)
                    elem.clear()
                elif len(stack) == 2 and elem.tag != NS + "sheetData":
                    self.emit(entry, "xlsx_sheet_metadata", sheet=meta["name"],
                              xml=ET.tostring(elem, encoding="unicode"))
                    stack[-2].remove(elem)
                    elem.clear()
                stack.pop()

    def scan_sqlite(self, path, entry):
        for suffix in ("-wal", "-journal"):
            sidecar = Path(str(path) + suffix)
            if sidecar.exists() and sidecar.stat().st_size:
                raise ValueError("active SQLite WAL/journal is unsupported; provide a closed checkpointed database")
        # immutable avoids journal/SHM writes; only used once sidecars are ruled out.
        with closing(sqlite3.connect(path.as_uri() + "?mode=ro&immutable=1", uri=True)) as conn:
            conn.execute("PRAGMA query_only=ON")
            schema = conn.execute("SELECT type,name,tbl_name,rootpage,sql FROM sqlite_schema ORDER BY type,name").fetchall()
            self.emit(entry, "sqlite_schema", columns=["type", "name", "table", "rootpage", "sql"], objects=schema)
            entry["tables"] = []
            for kind, name, _, _, sql in schema:
                if kind != "table":
                    continue
                table = {"name": name, "rows": 0, "cells": 0, "status": "complete", "ordering": "unspecified traversal; ordinal is not a source key"}
                entry["tables"].append(table)
                try:
                    columns = conn.execute(f"PRAGMA table_xinfo({quote(name)})").fetchall()
                    foreign = conn.execute(f"PRAGMA foreign_key_list({quote(name)})").fetchall()
                    self.emit(entry, "sqlite_table", table=name, sql=sql,
                              columns=columns, foreign_keys=foreign,
                              column_fields=["cid", "name", "declared_type", "not_null", "default", "pk_position", "hidden"])
                    cursor = conn.execute(f"SELECT * FROM {quote(name)}")
                    result_columns = [c[0] for c in cursor.description]
                    table["columns"] = result_columns
                    for ordinal, row in enumerate(cursor, 1):
                        values = []
                        for value in row:
                            if value is None:
                                values.append({"type": "null", "value": None})
                            elif isinstance(value, bytes):
                                values.append({"type": "blob", "base64": base64.b64encode(value).decode("ascii")})
                            elif isinstance(value, float) and not math.isfinite(value):
                                values.append({"type": "real", "special": str(value)})
                            else:
                                values.append({"type": "text" if isinstance(value, str) else "integer" if isinstance(value, int) else "real", "value": value})
                        self.emit(entry, "sqlite_row", table=name, ordinal=ordinal, values=values)
                        table["rows"] += 1
                        table["cells"] += len(values)
                        entry["rows"] += 1
                        entry["cells"] += len(values)
                except Exception as exc:
                    table["status"] = "incomplete"
                    self.error(entry, f"{type(exc).__name__}: {exc}", table=name)

    def run(self):
        paths = self.discover()
        with record_writer(self.output / "records.jsonl.gz") as self.records, \
                (self.output / "errors.jsonl").open("w", encoding="utf-8", newline="\n") as self.errors:
            available = {p.relative_to(self.split).as_posix() for p, _, _ in paths}
            required = {"rulebook.md", "calendar.csv", "questions.jsonl"}
            required.update(f"data/{business}/README.md" for business in self.manifest["businesses"])
            if not (self.split / "data").is_dir() or not self.manifest["businesses"]:
                required.add("data/<business_id>/")
            for missing in sorted(required - available):
                entry = {"path": missing, "kind": "missing", "status": "incomplete", "error_count": 0, "rows": 0, "cells": 0, "records": 0}
                self.manifest["files"].append(entry)
                self.error(entry, "required input missing")
            for path, kind, issue in paths:
                entry = {"path": path.relative_to(self.split).as_posix(), "kind": kind,
                         "status": "complete", "error_count": 0, "rows": 0, "cells": 0, "records": 0}
                self.manifest["files"].append(entry)
                if kind in {"unsupported", "unsafe"}:
                    self.error(entry, issue or "unsupported file; content not read")
                    continue
                try:
                    reject_links(path)
                    entry["bytes"] = path.stat().st_size
                    entry["sha256_before"] = digest(path)
                    if kind == "text":
                        # Stream lines; endings retained, no markdown processing.
                        with path.open("r", encoding="utf-8-sig", newline="") as handle:
                            for line, value in enumerate(handle, 1):
                                self.emit(entry, "text_line", line=line, text=value)
                                entry["rows"] += 1
                    else:
                        getattr(self, "scan_" + kind)(path, entry)
                except Exception as exc:
                    self.error(entry, f"{type(exc).__name__}: {exc}")
                finally:
                    if "sha256_before" in entry:
                        try:
                            entry["sha256_after"] = digest(path)
                            entry["source_unchanged"] = entry["sha256_before"] == entry["sha256_after"]
                            if not entry["source_unchanged"]:
                                self.error(entry, "source changed during scan")
                        except OSError as exc:
                            self.error(entry, f"post-scan hash failed: {exc}")
            self.manifest["status"] = "complete" if not self.manifest["error_count"] else "incomplete"
        files = self.manifest["files"]
        self.manifest["totals"] = {"files": len(files), "files_complete": sum(f["status"] == "complete" for f in files),
                                   "rows": sum(f["rows"] for f in files), "cells": sum(f["cells"] for f in files),
                                   "records": sum(f["records"] for f in files), "errors": self.manifest["error_count"],
                                   "sheets": sum(len(f.get("sheets", [])) for f in files),
                                   "tables": sum(len(f.get("tables", [])) for f in files)}
        with (self.output / "manifest.json").open("w", encoding="utf-8") as handle:
            json.dump(self.manifest, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
        with (self.output / "report.md").open("w", encoding="utf-8") as handle:
            handle.write(f"# Raw ingestion scan\n\nStatus: {self.manifest['status']}\n\n")
            handle.write("This is structural scanning, not semantic auditing or question answering.\n\n")
            handle.write("```json\n" + json.dumps(self.manifest["totals"], indent=2) + "\n```\n\n")
            handle.write("See manifest.json for every source hash, sheet/table coverage and error.\n")
        return self.manifest


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--split", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args(argv)
    try:
        manifest = Scanner(args.split, args.output).run()
    except (OSError, ValueError) as exc:
        print(f"ingestion refused: {exc}", file=sys.stderr)
        return 2
    print(json.dumps({"status": manifest["status"], **manifest["totals"], "output": str(args.output.resolve())}))
    return 0 if manifest["status"] == "complete" else 1
