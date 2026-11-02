# Step 1: raw ingestion

This layer reads a split, inventories its inputs and preserves source values and
locations. It does not identify headers, interpret business schemas, normalize
dates, fill grouped cells, deduplicate, merge entities, load business SQLite
tables, call models, or answer questions. A complete scan is not a semantic audit.

Run from the repository root in PowerShell, using the shared interpreter:

```powershell
& ./.venv/Scripts/python.exe -m unittest discover -s tests -v
& ./.venv/Scripts/python.exe -m ingestion --split train --output artifacts/ingestion/train
```

The output directory must be new or empty, outside the input split, and cannot be
an ancestor of the split. Existing nonempty output is refused. Choose another
output name for a second scan. Links and Windows reparse points/junctions are
rejected in input/output path components. Directory discovery failures are
reported. Source files are opened for reading only; no input is overwritten.
The CLI exits **0** only on a complete scan, **1** on a completed but incomplete
scan, and **2** when startup/path validation fails. An interruption can leave
partial outputs without a manifest; those are never evidence of completion.

## Output contract (version 1)

Every output lives under the explicitly supplied output directory:

| File | Meaning |
| --- | --- |
| `manifest.json` | Version, absolute split path, status, excluded label paths, sorted businesses, per-source coverage/status/error counts, SHA-256 before/after, aggregate totals and pointers to record/error streams. |
| `records.jsonl.gz` | Gzip-compressed UTF-8 JSON Lines, one raw record per line; stream with `gzip.open(path, "rt", encoding="utf-8")`. Gzip uses a fixed timestamp and no source filename. |
| `errors.jsonl` | UTF-8 JSON Lines, one error per line, including relative source path and row/line/sheet/table where known. This is the full error list; memory retains only counts. |
| `report.md` | Small structural scan summary, explicitly distinguishing it from semantic auditing. |

During question validation only, `_question_ids.sqlite` is a disk-backed duplicate
index under output. It is closed and removed after a successful question pass;
a failed pass may leave it as an incomplete-run artifact. No label is stored in it.

All records have `source` (POSIX path relative to split) and `kind`. File paths
are ordered lexically within phases: all README/rulebook text first, all CSV
second, questions third, XLSX fourth, SQLite fifth, unsupported/unsafe entries
last. Worksheets use original workbook order. Source rows retain their order.
SQLite tables are lexical by name; their row traversal order has no business or
relational ordering meaning, and may differ on another SQLite version.

Supported inputs are `rulebook.md`, `calendar.csv`, `questions.jsonl`, business
`README.md` and `.csv`, `.xlsx`, `.sqlite` files recursively below `data/`.
Required root inputs, at least one business, and each business README are checked.
All other files are inventoried as unsupported and make the scan incomplete.
Every file named `answers.jsonl` (case insensitive) is excluded **before any
content read or hashing**. Its path and exclusion reason alone are recorded.

### Text and CSV

- `text_line`: original `line` number and `text`, retaining newline characters.
  Only UTF-8, optionally with a BOM, is supported. The decoded BOM is removed;
  the original source hash still covers its bytes.
- `csv_row`: 1-based logical `row`, physical `line_start`/`line_end`, `values`
  (strings only) and original decoded `raw_record` with quoting and newlines.
  No header is dropped. Empty cells stay empty strings; `NULL`, `null`, `NA`
  and numbers stay text. CSV has no universal null marker: quoted empty and
  unquoted empty both parse as `""`, while `raw_record` preserves their spelling.
- Encoding is strict UTF-8/BOM, never guessed or replaced. Delimiter is inferred
  from a maximum 64 KiB sample among comma, semicolon, tab and pipe using Python
  `csv.Sniffer`. Its quote character, doubled-quote and initial-space settings
  are retained in the manifest. Newline spelling is simplified **only** for
  dialect inference; output records retain original spelling.
- A delimiter-free single-column file is allowed. Ambiguous/undetectable dialects
  are refused rather than silently selected. Other encodings, escape conventions
  and long/incomplete sniff samples can require an explicit future reader option;
  this implementation does not promise arbitrary dialect detection. Every row
  is scanned with strict CSV syntax checking. Width differences from the first
  row are recorded as errors; the offending row is still emitted unchanged.

### XLSX

The reader streams ZIP/XML directly, so Excel numeric serials never become Python
dates or float-rounded numbers. Output is sparse, matching XML storage:

- `xlsx_workbook`: workbook XML, including date epoch/workbook properties.
- `xlsx_styles`: original UTF-8 styles XML, including cell XF indexes and custom
  number formats; built-in `numFmtId` values keep their OOXML meaning.
- `xlsx_shared_string`: `index`, resolved `text`, and rich string XML. Cell
  attributes `t="s"` and `raw_value` refer to this zero-based index. The shared
  string table streams once, without holding all strings in RAM.
- `xlsx_sheet`: original name, workbook sheet attributes (`state`, sheet ID,
  relation ID) and worksheet ZIP member. Hidden and very-hidden sheets are read.
- `xlsx_row`: original sheet name, 1-based `row`, row attributes and explicit
  `cells`. Each cell has its original `coordinate`, 1-based `column`, all XML
  attributes (`t` cell type, `s` style index, etc.), `value_present`, lexical
  `raw_value`, optional `formula` text/attributes, and optional `inline_string`
  text/rich XML. Numeric values intentionally remain strings. Missing `t` means
  numeric according to OOXML, not an inferred business type.
- Formula `<v>` contents are the cached result, when stored. `value_present`
  distinguishes a missing cache element from an empty one. Shared/array formula
  attributes remain available; formulas are never recalculated or expanded.
- `xlsx_sheet_metadata`: each non-`sheetData` worksheet XML child, including
  dimensions, hidden columns, merge ranges, filters and other sheet metadata.

Absent row indexes and absent cell coordinates are implicit blanks, **not**
filled-down values. Explicit blank/styled cells remain explicit. Blank grid
positions can be reconstructed from row/cell coordinates plus worksheet
dimension/merge metadata; the enormous unused Excel grid is never expanded.
`rows` counts physical XML row elements, `cells` counts explicit XML cells;
`max_row`/`max_column` record observed bounds. Worksheet dimension is metadata,
never a reason to stop scanning. Duplicate data rows remain separate.

The supported XML namespace is transitional OOXML SpreadsheetML (the namespace
used by the supplied workbooks). Strict OOXML and non-UTF-8 styles XML are not
silently converted; unsupported structures/encodings make scanning incomplete.
This is tabular preservation, not a round-trip XLSX editor: drawings, macros,
comments, external connection contents, rendering and all ZIP members are not
copied. Non-worksheet sheets, invalid/missing coordinates, inaccessible/corrupt
ZIP/XML and duplicate ZIP member names fail explicitly. Other discoverable
sheets/files continue scanning. Original input remains the byte-level source
of truth; XML serialization may change namespace prefixes without changing data.

### SQLite

- Opened with URI `mode=ro&immutable=1` and `query_only=ON`, avoiding source-side
  journals/SHM. Nonempty WAL/journal sidecars are refused; supply a closed,
  checkpointed database. Concurrent source modifications are unsupported.
- `sqlite_schema`: all schema objects and their original SQL. Views, indexes
  and triggers are inventoried here; views are not materialized as tables.
- `sqlite_table`: every schema table (including SQLite internal tables), SQL,
  `table_xinfo` columns and `foreign_key_list`; field names for column tuples
  are included. PK positions and declared types are preserved. Identifiers are
  quoted, including embedded quotes; `WITHOUT ROWID` is supported.
- `sqlite_row`: table, 1-based traversal `ordinal` and positional `values` in
  manifest `columns` order. Values are tagged `null`, `integer`, `real`, `text`
  or `blob`. Blobs use base64; nonfinite REAL uses a `special` string. No null
  becomes empty text and no assumption that every table has a rowid is made.
- A table error marks that table/file incomplete and subsequent tables continue.

### Questions, coverage and completeness

`question` records contain original `line` and parsed `value`. Only exactly
`id`, `business_id`, `question`, `as_of_date` are allowed, all nonempty strings;
date must be a real ISO `YYYY-MM-DD` day, business must exist, ID must be unique.
Duplicate JSON keys, malformed JSON, blank lines, nonfinite JSON numbers, extra
fields (including answer labels), wrong types and bad dates are errors. Invalid
questions are not emitted; scanning continues through later lines. IDs are checked
with a disk index rather than an unbounded in-memory set.

Every supported readable file is hashed before/after its scan. `source_unchanged`
compares those hashes. Read failure, changed hash, unsupported input, discovery
failure, row-width error or invalid question makes the run incomplete. Inspect
per-file and per-sheet/table status as well as totals; partial emitted records
do not imply a successful file. File failures do not suppress later files.
Counts include text lines and question lines in `rows`; only tabular source
cells count in `cells`. Question count is separately available as
`valid_questions`. Files, schema/metadata, one row, or one text line can occupy
memory; dataset rows/shared strings/error lists are not accumulated. A single
unusually large row/metadata XML node may still require significant memory.

Fixture tests use only temporary synthetic sources under `.venv/tmp/ingestion-tests`.
They never inspect actual answer labels. Runtime ingestion uses the Python
standard library only; existing pinned `openpyxl` is used to build test fixtures.
