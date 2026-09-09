# SCM Allocation Mail

Safely scans Outlook allocation-request emails, extracts supported HTML/XLSX structures into traceable `AllocationRecord` values, and optionally applies Outlook processing categories. It is deliberately independent of PostgreSQL; a later persistence layer can consume its structured records.

## Pipeline

```text
Outlook Email (or Local Sample)
            ↓
     Email Filtering
            ↓
    Content Extraction
    ├── HTML table parser
    └── Excel attachment parser
            ↓
Field Detection & Normalization
            ↓
  Normalization / SLoc defaults
            ↓
 Structured `AllocationRecord`
            ↓
      Output / Report
```

## Operating safely

Windows, Outlook desktop, and `pywin32` are required for live Outlook scans. The default command is read-only with respect to Outlook:

```powershell
.\.venv\Scripts\python.exe scripts\process_outlook_allocations.py --folder Inbox
```

Only this explicit command may persist Outlook Categories:

```powershell
.\.venv\Scripts\python.exe scripts\process_outlook_allocations.py --folder Inbox --write-categories
```

The bot never moves, deletes, replies to, forwards, changes body/attachments, or changes `UnRead`. `VALID` receives `SCM Bot - Scraped`; `PARTIAL` receives `SCM Bot - Review`; INVALID, unsupported, and Batam/out-of-scope messages receive no bot category.

Already categorized messages are skipped before extraction. Outlook Categories—not subject, attachment name, in-memory repository state, or CSV history—are the authoritative duplicate guard.

## Supported extraction formats

- Compound allocation grids.
- Destination and warehouse/destination matrices.
- Simple Excel tables and site-only review tables.
- NPI wide site/material matrices (material codes above descriptions).
- IT/LOOPS new-store tables with source/destination SLoc columns.
- Supported flat HTML and request/pivot HTML tables.

Unsupported or malformed attachments produce controlled INVALID diagnostics; they do not fabricate allocation records.

## Storage-location rules

Explicit workbook source/destination SLoc values always win. When source SLoc is absent, BGC1/BHC1/BIC1/BFC1 use `1005` for DEVICE context and `1001` for NON-DEVICE or uncertain context. A destination with no explicit SLoc defaults to `1001`.

## Output and audit

`AllocationRecord` keeps material, quantity, source/destination warehouse and SLoc values, context, parser/status/confidence/errors, and source email/file/sheet/row/cell traceability. This is ready for a PostgreSQL mapper without coupling the parser to a database.

Local audit files are ignored by Git:

- `logs/allocation_bot_history.csv`: successful live category persistence only.
- `logs/allocation_bot_runs.csv`: every scanner run and counters.

## Setup and tests

## 3. Project Structure

```text
SCM-Allocation-Mail/
├── config/              # Configuration files (filters, keyword rules, field aliases)
├── data/                # Local data storage and master reference data
├── src/
│   └── scm_allocation/  # Core package
│       ├── config/      # Configuration loaders and settings
│       ├── ingestion/   # Email adapters and raw email objects
│       ├── models/      # Canonical data models (AllocationRequest, EmailMessage)
│       ├── normalization/ # Field mapping, numeric cleaning, alias resolution
│       ├── output/      # Report generation and export handlers
│       ├── parsers/     # Text, HTML, Table, and Excel parsers
│       ├── utils/       # Common helpers, date utilities, logging
│       └── validation/  # Master data validation and discrepancy checks
├── tests/
│   ├── fixtures/        # Sanitized sample files and mock data
│   ├── integration/     # Integration tests
│   └── unit/            # Unit tests for parsers, models, and validators
├── pyproject.toml       # Build and dependency configuration
└── README.md
```

## 4. Getting Started

### Prerequisites
- Python 3.10+ (tested with Python 3.13.14 on Windows)
- Dedicated virtual environment (`.venv`)

### Setup Virtual Environment
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e ".[dev]"
```

### Run tests
```powershell
.\.venv\Scripts\pytest.exe -v --basetemp .pytest-final-validation
```

## End-to-End Demo

### 1. Prerequisites
- Windows OS with Microsoft Outlook desktop installed and logged in.
- Virtual environment activated with project dependencies (`pip install -e ".[dev]"` or `poetry install`).

### 2. Exact Command
To run a safe live scan of your Outlook Inbox in dry-run mode:
```powershell
python scripts/process_outlook_allocations.py --folder Inbox
```
You can optionally inspect only the newest $N$ messages:
```powershell
python scripts/process_outlook_allocations.py --folder Inbox --max-emails 10
```

### 3. What the Command Does
1. Connects to Outlook desktop via MAPI COM (`pywin32`).
2. Filters emails by sender domain (`erajaya.com`), keywords (`alokasi`, `allocation`), and excludes out-of-scope subjects (e.g. Batam).
3. Applies duplicate guard based on Outlook Categories (`SCM Bot - Scraped`, `SCM Bot - Review`); already categorized messages are skipped.
4. Filters non-XLSX attachments (logos, images, `.htm` signatures) regardless of attachment order.
5. Dispatches candidate `.xlsx` workbooks to appropriate Excel/HTML parsers.
6. Normalizes SLocs, plant codes, and quantities into traceable `AllocationRecord[]` objects.
7. Logs execution metrics and counters to `logs/allocation_bot_runs.csv`.

### 4. Dry-Run Safety Behavior
By default (without `--write-categories`):
- Read-only execution: does NOT send, delete, move, or modify emails.
- Preserves `UnRead` flag state.
- Does NOT apply Outlook categories.
- Does NOT write to `logs/allocation_bot_history.csv`.

### 5. Expected Output
The CLI displays progress, email summaries, parser diagnostics, and extracted records:
```text
[SCM] Starting Outlook allocation scan (folder=Inbox, max_emails=10, dry_run=True).
[SCM] Connected to Outlook.
[SCM] Scanning folder: Inbox.
Subject: Request Alokasi NPI Anker | Email ID: 0000... | Status: VALID | Parser: excel_site_material_matrix | records=10 | Planned category: SCM Bot - Scraped | Written: False
  workbook=npi_anker.xlsx sheet=Sheet1 -> matched: excel_site_material_matrix records=10
  -> AllocationRecord: material=8100294900 (ANK Charger 0) qty=1 from=<none> (sloc=1001) to=X015 (sloc=1001) context=accessories src=npi_anker.xlsx/Sheet1 row=3
Scanned: 10
Candidates: 1
Processed: 1
Skipped already processed: 0
Scraped: 1
Review: 0
Invalid/Excluded: 0
Dry-run: True
```

### 6. Category Persistence (--write-categories)
When run with the explicit opt-in flag:
```powershell
python scripts/process_outlook_allocations.py --folder Inbox --write-categories
```
- Messages yielding `VALID` extraction receive category `SCM Bot - Scraped`.
- Messages yielding `PARTIAL` extraction receive category `SCM Bot - Review`.
- Successfully categorized runs record processed email details to `logs/allocation_bot_history.csv`.

### 7. Known Limitations
- Outlook COM automation requires Windows OS with desktop Outlook running.
- Unsupported or malformed attachments yield `INVALID` status and do not fabricate records.
- Database persistence is out of scope for this repository; a downstream consumer will ingest `AllocationRecord[]` into PostgreSQL.

## Security and handoff
- Never commit credentials, passwords, OAuth tokens, or secrets.
- Never commit confidential production emails or sensitive employee personal data to version control.
