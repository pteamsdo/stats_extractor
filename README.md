# Monthly Excel to SQLite updater

This tool reads one monthly Excel workbook and adds new data to a SQLite database. Running it again with the same workbook is safe: it will not create duplicate rows.

Excel files are read with the common open-source Python library `openpyxl`.

The computer only needs Python available in `PATH`. On the first run, `start.bat` automatically creates a local `.venv` folder and installs the packages in `requirements.txt`. Later runs reuse that environment.

## What to do each month

1. Open the `raw` folder.
2. Remove the previous Excel workbook or move it somewhere else.
3. Put the new `.xlsx` workbook in `raw`. There must be exactly one Excel file there.
4. Double-click `start.bat`, or open Command Prompt in this folder and run:

   ```bat
   start.bat
   ```

5. When it finishes, the database is at:

   ```text
   data\usage_stats.sqlite
   ```

## What the tool expects

The Excel filename can be anything. The workbook must contain these two worksheets:

- `Downloads`
- `API Calls`

Month columns are discovered from headings such as `May 2026` or `Jun 2026`. The tool also finds `Access Group`, `Remarks`, and `ASDPID` columns by their headings, so their column letters may move when a new month is added.

## What happens during an update

- A brand-new month is inserted.
- A corrected value for an existing month is updated.
- An unchanged value is left alone.
- Blank, `N/A`, and `TBP` cells are kept as explicit statuses.
- Old database rows are not deleted when they are absent from a later workbook.

The command prints three useful numbers:

- `inserted`: new database rows
- `updated`: existing rows whose data changed
- `unchanged`: rows already stored with the same data

## Optional command-line settings

Normally you do not need these. They are useful for testing or storing the database elsewhere.

```bat
start.bat --input "C:\path\another-file.xlsx"
start.bat --database "C:\path\another-database.sqlite"
```

## Querying by date

SQLite tables do not have a permanent display order. Use the included chronological view:

```sql
SELECT *
FROM observations_chronological;
```

The `import_runs` table records each successful run and how many rows were inserted, updated, or unchanged.

For the full database contract, table schema, indexes, value rules, and example queries, see:

```text
DATABASE_SCHEMA.md
```

## Deleting the database and starting over

This permanently removes the generated SQLite database. It does not delete anything from `raw`.

Double-click:

```text
clean_database.bat
```

Or run this command:

```bat
start.bat --purge
```

The command removes `data\usage_stats.sqlite` and any SQLite `-wal`, `-shm`, or `-journal` helper files. The next normal `start.bat` run will create a new empty database and import the workbook in `raw`.

To purge a database supplied with `--database`, use:

```bat
start.bat --purge --database "C:\path\another-database.sqlite"
```
