# Agent Overview

Purpose: convert one monthly Excel workbook into an idempotent SQLite usage
database.

## Repo Map

- `src/pipeline.py`: only application source file. Handles CLI args, workbook
  discovery, Excel parsing, SQLite schema creation, upserts, and purge mode.
- `raw/`: ignored input area. Normal runs expect exactly one `.xlsx` here.
- `data/usage_stats.sqlite`: ignored generated database.
- `DATABASE_SCHEMA.md`: database contract, current live profile, value rules,
  indexes, and common SQL.
- `README.md`: human monthly operating instructions.
- `start.bat`: Windows entrypoint; creates/reuses `.venv`, installs
  `requirements.txt`, then runs `src/pipeline.py`.
- `clean_database.bat`: Windows helper for purge mode.

## Runtime

- Language: Python.
- Dependency: `openpyxl`.
- Database: SQLite through Python stdlib `sqlite3`.
- No web service, scheduler, package build, or test suite is present.

## Main Commands

```bat
start.bat
start.bat --input "C:\path\file.xlsx"
start.bat --database "C:\path\usage_stats.sqlite"
start.bat --purge
```

Direct Python equivalent:

```bat
python src\pipeline.py --input "C:\path\file.xlsx" --database "data\usage_stats.sqlite"
```

## Input Contract

Workbook must contain:

- `Downloads`
- `API Calls`

First three headers must be:

- `Data Provider`
- `Dataset Name`
- `Data Nature`

Metadata headers are found by name:

- `Access Group`
- `Remarks`
- `ASDPID`

Month columns are parsed from headings like `May 2026` or `Jun 2026`.

## Database Contract

Schema is created in `initialize_database()` in `src/pipeline.py`.

Objects:

- `observations`: fact table.
- `import_runs`: successful import audit table.
- `observations_chronological`: ordered view over `observations`.

Important keys:

- `record_key`: unique SHA-256 idempotency key.
- `dataset_identity`: `asdp:<id>` if `ASDPID` exists, otherwise fallback from
  provider, dataset name, and data nature.

Allowed `metric` values:

- `api_calls`
- `downloads`

Allowed `value_status` values:

- `reported`
- `missing`
- `not_applicable`
- `other`

Use `DATABASE_SCHEMA.md` for full column details and example queries.

## Behavioral Notes

- Re-running the same workbook is safe; unchanged semantic rows are skipped.
- Existing rows are updated when the same `record_key` has changed semantic
  fields.
- Rows missing from later workbooks are not deleted.
- Blank cells become `missing`.
- `N/A`, `NA`, and `NOT APPLICABLE` become `not_applicable`.
- Numeric cells and numeric text become `reported`.
- Other text, currently including `TBP`, becomes `other`.
- The loader writes `PRAGMA journal_mode = WAL` and `PRAGMA synchronous = NORMAL`.

## Edit Guidance

- Keep generated files out of git: `raw/*`, `data/*.sqlite`,
  `data/*.sqlite-wal`, and `data/*.sqlite-shm` are ignored.
- Prefer changing `src/pipeline.py` over documenting behavior that the loader
  does not actually implement.
- If schema behavior changes, update `DATABASE_SCHEMA.md` in the same patch.
- Preserve idempotency semantics around `record_key`.
- Do not purge or rebuild the database unless the task explicitly requires it.
- The repository may contain user edits; inspect status before changing files.
