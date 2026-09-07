# SCM-Allocation-Mail

Automated collection, extraction, and validation of allocation requests sent by PM/CM to the SCM department via Microsoft Outlook.

## 1. Overview

PM/CM sends allocation requests to SCM across varied formats, including:
- Plain text email bodies
- HTML tables with varying structures, column names, and column orders
- Excel attachments (.xlsx, .xls)
- Hybrid messages (email body summary + Excel attachment)

This system automates the ingestion, parsing, normalization, validation against master reference data, and consolidation of these requests into a canonical structured format while preserving complete traceability.

## 2. Architecture & Pipeline

```text
Outlook Email (or Local Sample)
            ↓
     Email Filtering
            ↓
    Content Extraction
    ├── Plain Text Parser
    ├── HTML Body Parser
    ├── HTML Table Parser
    └── Excel Attachment Parser
            ↓
Field Detection & Normalization
            ↓
  Master Data Validation (WH, Plant, Material)
            ↓
 Structured Allocation Request
            ↓
      Output / Report
```

### Core Design Principles
- **Accuracy & Traceability**: Preserve original extracted values, email metadata, timestamps, and parser methods for auditability.
- **Strict Validation**: Distinct statuses (`VALID`, `INVALID`, `MISSING`, `AMBIGUOUS`, `UNMATCHED`). Never guess or hallucinate business data.
- **Sample-First Development**: Parsers are designed and tested against real-world sample patterns before production deployment.
- **Layer Separation**: Parsing and normalization logic are decoupled from Outlook COM / MAPI integration to enable robust offline testing.

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

### Running Tests
```powershell
.\.venv\Scripts\pytest
```

## 5. Security & Privacy
- Never commit credentials, passwords, OAuth tokens, or secrets.
- Never commit confidential production emails or sensitive employee personal data to version control.
