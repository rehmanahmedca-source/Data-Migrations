# Relay — Data Migration Workspace

A focused, browser-based prototype for safely moving legacy data into a new schema. It supports multiple legacy files, new schema selection, drag-and-drop upload, a discovery/review step with healthy, partial, orphan, and voided record groups, per-group inclusion selection, and export configuration for CSV, JSON, XLSX, or SQLite.

## Run locally

```bash
python3 -m http.server 4173 --bind 0.0.0.0
```

Open `http://localhost:4173`.

This is a front-end workflow prototype; file processing and actual export can be wired to a backend migration engine next.
