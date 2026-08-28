# Test files

Place sample legacy source files and destination schema files here while testing the migration workflow.

Suggested layout:

- `legacy/` — XLSX, CSV, JSON, DB, or SQLite source files
- `schemas/` — destination schema files
- `expected/` — expected migrated output for comparison

Do not store production or sensitive customer data in this folder. Use anonymized fixtures only.
