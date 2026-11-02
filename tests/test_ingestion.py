import datetime
from contextlib import closing
import json
import gzip
from pathlib import Path
import sqlite3
import tempfile
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as ET
import zipfile

from openpyxl import Workbook
from ingestion.reader import Scanner, main


class IngestionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.previous_tempdir = tempfile.tempdir
        cls.fixture_root = Path(__file__).resolve().parents[1] / ".venv" / "tmp" / "ingestion-tests"
        cls.fixture_root.mkdir(parents=True, exist_ok=True)
        tempfile.tempdir = str(cls.fixture_root)

    @classmethod
    def tearDownClass(cls):
        tempfile.tempdir = cls.previous_tempdir

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.split = self.root / "split"
        self.business = self.split / "data" / "business"
        self.business.mkdir(parents=True)
        (self.business / "README.md").write_text("# Tiệm\n\nRaw notes.\n", encoding="utf-8")
        (self.split / "rulebook.md").write_text("# Rules\n", encoding="utf-8")
        (self.split / "calendar.csv").write_text("date,holiday\n2026-01-01,Tết\n", encoding="utf-8")
        self.question = {"id": "q1", "business_id": "business", "question": "How much?", "as_of_date": "2026-01-01"}
        self.questions([self.question])
        # Excluded input is deliberately invalid UTF-8; the reader must never decode it.
        (self.split / "answers.jsonl").write_bytes(b"\xff\x00secret-label")
        self.output = self.root / "output"

    def questions(self, objects):
        (self.split / "questions.jsonl").write_text("".join(json.dumps(o) + "\n" for o in objects), encoding="utf-8")

    def scan(self):
        manifest = Scanner(self.split, self.output).run()
        with gzip.open(self.output / "records.jsonl.gz", "rt", encoding="utf-8") as handle:
            records = [json.loads(line) for line in handle]
        errors = [json.loads(line) for line in (self.output / "errors.jsonl").read_text(encoding="utf-8").splitlines()]
        return manifest, records, errors

    def test_csv_preserves_text_quoting_blank_duplicates_and_provenance(self):
        raw = 'id;name;note;empty\r\n001;"Cà; phê";"line1\nline2";""\r\n;NULL;;\r\n;NULL;;\r\n'
        file = self.business / "sales.csv"
        file.write_bytes(b"\xef\xbb\xbf" + raw.encode("utf-8"))
        before = file.read_bytes()
        manifest, records, errors = self.scan()
        self.assertEqual(errors, [])
        rows = [r for r in records if r["source"].endswith("sales.csv")]
        self.assertEqual(rows[1]["values"], ["001", "Cà; phê", "line1\nline2", ""])
        self.assertEqual(rows[2]["values"], ["", "NULL", "", ""])
        self.assertEqual(rows[2]["values"], rows[3]["values"])
        self.assertIn(';""\r\n', rows[1]["raw_record"])
        self.assertEqual(rows[1]["line_start"], 2)
        self.assertEqual(rows[1]["line_end"], 3)
        self.assertEqual(file.read_bytes(), before)
        self.assertEqual(manifest["status"], "complete")
        self.assertTrue(all(f["source_unchanged"] for f in manifest["files"]))
        self.assertEqual(manifest["excluded"][0]["path"], "answers.jsonl")
        self.assertFalse(any("secret-label" in str(r) for r in records))

    def test_single_column_csv(self):
        (self.business / "one.csv").write_text("header\nTiệm\nNULL\n", encoding="utf-8")
        manifest, records, errors = self.scan()
        self.assertEqual(errors, [])
        self.assertEqual(manifest["status"], "complete")

    def test_nondefault_csv_quotes(self):
        (self.business / "quotes.csv").write_text("id; name\n1; 'A; B'\n", encoding="utf-8")
        manifest, records, errors = self.scan()
        self.assertEqual(errors, [])
        row = next(r for r in records if r["source"].endswith("quotes.csv") and r["row"] == 2)
        self.assertEqual(row["values"], ["1", "A; B"])
        self.assertEqual(row["raw_record"].rstrip("\r\n"), "1; 'A; B'")

    def test_xlsx_all_sheets_rows_types_formulas_and_metadata(self):
        path = self.business / "book.xlsx"
        book = Workbook()
        sheet = book.active
        sheet.title = "Tháng 1"
        sheet["A1"] = "Report"
        sheet.merge_cells("A1:C1")
        sheet.append([None])
        sheet.append(["Order", "Date", "Amount", "Formula"])
        sheet.append(["001", datetime.datetime(2026, 1, 2), 42, "=C4*2"])
        sheet.append([None, "02/01", 42, None])
        sheet.append([None, "02/01", 42, None])
        sheet.row_dimensions[5].hidden = True
        sheet.column_dimensions["C"].hidden = True
        hidden = book.create_sheet("Hidden data")
        hidden.sheet_state = "veryHidden"
        hidden["A1"] = 46024
        hidden["A1"].number_format = "yyyy-mm-dd"
        book.save(path)
        # Inject a cached result; openpyxl intentionally does not calculate formulas.
        with zipfile.ZipFile(path) as source:
            members = [(info, source.read(info.filename)) for info in source.infolist()]
        with zipfile.ZipFile(path, "w") as target:
            for info, content in members:
                if info.filename == "xl/worksheets/sheet1.xml":
                    xml = ET.fromstring(content)
                    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
                    xml.find(f".//{ns}c[@r='D4']/{ns}v").text = "84"
                    content = ET.tostring(xml, encoding="utf-8")
                target.writestr(info, content)
        manifest, records, errors = self.scan()
        self.assertEqual(errors, [])
        entry = next(f for f in manifest["files"] if f["kind"] == "xlsx")
        self.assertEqual([s["name"] for s in entry["sheets"]], ["Tháng 1", "Hidden data"])
        self.assertEqual(entry["sheets"][1]["state"], "veryHidden")
        rows = [r for r in records if r["kind"] == "xlsx_row" and r["sheet"] == "Tháng 1"]
        self.assertEqual(next(r for r in rows if r["row"] == 3)["cells"][0]["inline_string"]["text"], "Order")
        row4 = next(r for r in rows if r["row"] == 4)
        self.assertEqual(row4["cells"][1]["raw_value"], "46024")
        self.assertEqual(row4["cells"][3]["formula"]["text"], "C4*2")
        self.assertEqual(row4["cells"][3]["raw_value"], "84")
        row5 = next(r for r in rows if r["row"] == 5)
        self.assertEqual(row5["cells"][0]["coordinate"], "B5")  # A stays absent/blank.
        self.assertEqual(row5["attributes"]["hidden"], "1")
        self.assertEqual([c["raw_value"] for c in row5["cells"]],
                         [c["raw_value"] for c in next(r for r in rows if r["row"] == 6)["cells"]])
        metadata = "".join(r["xml"] for r in records if r["kind"] == "xlsx_sheet_metadata")
        self.assertIn("A1:C1", metadata)
        self.assertIn('hidden="1"', metadata)
        self.assertTrue(any(r["kind"] == "xlsx_styles" and "numFmt" in r["xml"] for r in records))

    def test_sqlite_quoted_names_without_rowid_null_blob_types(self):
        path = self.business / "records.sqlite"
        with closing(sqlite3.connect(path)) as conn, conn:
            conn.execute('CREATE TABLE "parent" (id TEXT PRIMARY KEY) WITHOUT ROWID')
            conn.execute('INSERT INTO parent VALUES (?)', ("key",))
            conn.execute('CREATE TABLE "odd"" name" (id TEXT PRIMARY KEY, value BLOB, missing TEXT, n REAL, FOREIGN KEY(id) REFERENCES parent(id)) WITHOUT ROWID')
            conn.execute('INSERT INTO "odd"" name" VALUES (?,?,?,?)', ("key", b"\x00\xff", None, 2.5))
        manifest, records, errors = self.scan()
        self.assertEqual(errors, [])
        row = next(r for r in records if r["kind"] == "sqlite_row" and r["table"] == 'odd" name')
        self.assertEqual(row["values"], [{"type": "text", "value": "key"}, {"type": "blob", "base64": "AP8="},
                                         {"type": "null", "value": None}, {"type": "real", "value": 2.5}])
        table = next(r for r in records if r["kind"] == "sqlite_table" and r["table"] == 'odd" name')
        self.assertEqual(table["columns"][0][5], 1)
        self.assertEqual(table["foreign_keys"][0][2], "parent")
        self.assertEqual(manifest["totals"]["tables"], 2)

    def test_shared_strings_keep_indices_and_rich_text(self):
        path = self.business / "shared.xlsx"
        book = Workbook()
        book.active["A1"] = "placeholder"
        book.save(path)
        ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        with zipfile.ZipFile(path) as source:
            members = [(info, source.read(info.filename)) for info in source.infolist()]
        with zipfile.ZipFile(path, "w") as target:
            for info, content in members:
                if info.filename == "xl/worksheets/sheet1.xml":
                    xml = ET.fromstring(content)
                    cell = xml.find(f".//{ns}c[@r='A1']")
                    cell.remove(cell.find(ns + "is"))
                    cell.set("t", "s")
                    ET.SubElement(cell, ns + "v").text = "0"
                    content = ET.tostring(xml, encoding="utf-8")
                elif info.filename == "xl/_rels/workbook.xml.rels":
                    xml = ET.fromstring(content)
                    package_ns = "{http://schemas.openxmlformats.org/package/2006/relationships}"
                    for relationship in xml:
                        if relationship.get("Type", "").endswith("/styles"):
                            relationship.set("Target", "/xl/styles-alt.xml")
                    ET.SubElement(xml, package_ns + "Relationship", Id="strings",
                                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
                                  Target="texts/strings.xml")
                    content = ET.tostring(xml, encoding="utf-8")
                elif info.filename == "xl/styles.xml":
                    target.writestr("xl/styles-alt.xml", content)
                    continue
                elif info.filename == "[Content_Types].xml":
                    xml = ET.fromstring(content)
                    ct = "{http://schemas.openxmlformats.org/package/2006/content-types}"
                    for override in xml:
                        if override.get("PartName") == "/xl/styles.xml":
                            override.set("PartName", "/xl/styles-alt.xml")
                    ET.SubElement(xml, ct + "Override", PartName="/xl/texts/strings.xml",
                                  ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
                    content = ET.tostring(xml, encoding="utf-8")
                target.writestr(info, content)
            target.writestr("xl/texts/strings.xml", '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><t xml:space="preserve"> Cà </t></r><r><t>phê</t></r><rPh sb="0" eb="1"><t>phonetic</t></rPh></si></sst>')
        manifest, records, errors = self.scan()
        self.assertEqual(errors, [])
        shared = next(r for r in records if r["kind"] == "xlsx_shared_string")
        self.assertEqual((shared["index"], shared["text"]), (0, " Cà phê"))
        self.assertIn("space", shared["xml"])
        self.assertIn("phonetic", shared["xml"])
        self.assertTrue(any(r["kind"] == "xlsx_styles" for r in records))
        cell = next(r for r in records if r["kind"] == "xlsx_row")["cells"][0]
        self.assertEqual(cell["attributes"]["t"], "s")
        self.assertEqual(cell["raw_value"], "0")
        with zipfile.ZipFile(path) as source:
            members = [(info, source.read(info.filename)) for info in source.infolist()]
        with zipfile.ZipFile(path, "w") as target:
            for info, content in members:
                if info.filename == "xl/worksheets/sheet1.xml":
                    xml = ET.fromstring(content)
                    xml.find(f".//{ns}c[@r='A1']/{ns}v").text = "9"
                    content = ET.tostring(xml, encoding="utf-8")
                target.writestr(info, content)
        self.output = self.root / "bad-reference"
        manifest, records, errors = self.scan()
        self.assertEqual(manifest["status"], "incomplete")
        self.assertTrue(any("invalid shared string reference" in e["message"] for e in errors))

    def test_missing_required_input_still_scans_available_files(self):
        (self.split / "rulebook.md").unlink()
        manifest, records, errors = self.scan()
        self.assertEqual(manifest["status"], "incomplete")
        self.assertTrue(any(e["source"] == "rulebook.md" and e["message"] == "required input missing" for e in errors))
        self.assertTrue(any(r["kind"] == "question" for r in records))

    def test_bad_questions_continue_and_exclude_labels(self):
        self.questions([self.question, self.question, {**self.question, "id": "bad_date", "as_of_date": "2026-02-30"},
                        {**self.question, "id": 12}, {**self.question, "answer": "secret"}])
        with (self.split / "questions.jsonl").open("a", encoding="utf-8") as handle:
            handle.write('{bad json}\n')
        manifest, records, errors = self.scan()
        self.assertEqual(manifest["status"], "incomplete")
        self.assertEqual(len(errors), 5)
        self.assertEqual(len([r for r in records if r["kind"] == "question"]), 1)
        self.assertFalse(any("secret" in str(r) for r in records))

    def test_corrupt_file_does_not_hide_other_files(self):
        (self.business / "bad.xlsx").write_bytes(b"not a zip")
        (self.business / "good.csv").write_text("a,b\n1,2\n", encoding="utf-8")
        (self.business / "unknown.bin").write_bytes(b"unknown")
        manifest, records, errors = self.scan()
        self.assertEqual(manifest["status"], "incomplete")
        self.assertEqual(len(errors), 2)
        self.assertTrue(any(r["source"].endswith("good.csv") for r in records))
        self.assertEqual(main(["--split", str(self.split), "--output", str(self.root / "cli")]), 1)

    def test_invalid_encoding_and_width_are_reported(self):
        (self.business / "bad.csv").write_bytes(b"a,b\n\xff,2\n")
        # First 64 KiB is regular, then a bad width must still be discovered.
        (self.business / "width.csv").write_text("a,b\n" + "1,2\n" * 20000 + "3,4,5\n", encoding="utf-8")
        manifest, records, errors = self.scan()
        self.assertEqual(manifest["status"], "incomplete")
        self.assertTrue(any("row width" in e["message"] for e in errors))
        self.assertTrue(any("UnicodeDecodeError" in e["message"] for e in errors))
        self.assertTrue(any(r.get("values") == ["3", "4", "5"] for r in records))

    def test_discovery_error_is_not_silent(self):
        import os
        real_walk = os.walk
        def walk(*args, **kwargs):
            kwargs["onerror"](PermissionError(13, "permission denied", str(self.split / "data" / "unreadable")))
            yield from real_walk(*args, **kwargs)
        with patch("ingestion.reader.os.walk", side_effect=walk):
            manifest, records, errors = self.scan()
        self.assertEqual(manifest["status"], "incomplete")
        self.assertTrue(any("directory discovery failed" in e["message"] for e in errors))
        self.assertTrue(any(r["kind"] == "question" for r in records))

    def test_output_protection_and_links(self):
        with self.assertRaisesRegex(ValueError, "outside split"):
            Scanner(self.split, self.business / "output")
        self.output.mkdir()
        marker = self.output / "keep"
        marker.write_text("untouched")
        with self.assertRaisesRegex(ValueError, "new or empty"):
            Scanner(self.split, self.output)
        self.assertEqual(marker.read_text(), "untouched")
        with self.assertRaisesRegex(ValueError, "ancestor"):
            Scanner(self.split, self.root)
        link = self.root / "link"
        try:
            link.symlink_to(self.business, target_is_directory=True)
        except OSError:
            self.skipTest("Windows symlink privilege unavailable; junction tested separately")
        with self.assertRaisesRegex(ValueError, "symlink/reparse"):
            Scanner(self.split, link / "output")


if __name__ == "__main__":
    unittest.main()
