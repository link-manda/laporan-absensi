# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Current repository state

Repository currently minimal. No application code, README, Cursor rules, or Copilot instructions found.

Python dependencies currently tracked in `requirements.txt`.

## Commands

Create virtual environment:

```bash
python3 -m venv .venv
```

Activate virtual environment:

```bash
source .venv/bin/activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Current dependency set:
- `pandas`
- `openpyxl`

No build, lint, or test commands configured yet.

## Architecture

No application architecture exists yet.

Current repository only establishes Python environment baseline for spreadsheet work:
- `requirements.txt` defines runtime packages
- `.venv/` is local virtual environment for development
- `pandas` handles tabular data processing
- `openpyxl` provides Excel `.xlsx` read/write support used by pandas workflows

## Notes for future updates

If README.md, .cursor/rules/, .cursorrules, or .github/copilot-instructions.md are added later, fold only non-obvious repository-specific guidance into this file.
