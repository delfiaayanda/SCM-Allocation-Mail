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

## Security and handoff
- Never commit credentials, passwords, OAuth tokens, or secrets.
- Never commit confidential production emails or sensitive employee personal data to version control.
