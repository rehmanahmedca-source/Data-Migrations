#!/usr/bin/env python3
"""
Legacy -> New App data migration engine.

What it does
------------
1. Reads the NEW APP schema workbook (test-files/schemas/New App Data.xlsx) and keeps
   ONLY its structure: sheet (table) names + column names + column order.
   All dummy rows are discarded -> "empty schema".
2. Reads the LEGACY workbook (test-files/legacy/...xlsx).
3. Copies legacy rows into the new-app schema, matching by column name.
   Columns that exist only in the new schema are left NULL (except a few documented
   safe defaults).
4. Skips VOIDED rows (is_void truthy) and cascades that skip to owned child rows.
5. Repairs ORPHAN rows (rows that point at a client that does not exist in `client`):
   a synthetic client "Orphan1", "Orphan2", ... is created for every distinct unknown
   client identity, the orphan rows are re-pointed at it, and the original client name
   is preserved in the client's page_notes and in each row's note column.
6. Nulls out dangling soft foreign keys instead of deleting the row.
7. Writes XLSX / JSON / NDJSON / CSV / SQLite .db / .sql (schema + data) outputs,
   plus a comparison report.

Usage:  python3 tools/migrate.py
"""

from __future__ import annotations

import csv
import datetime as dt
import json
import os
import shutil
import sqlite3
import sys
from collections import Counter, defaultdict

import openpyxl
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LEGACY_XLSX = os.path.join(ROOT, "test-files", "legacy", "Legacy Data28-08-2026_05-31PM.xlsx")
SCHEMA_XLSX = os.path.join(ROOT, "test-files", "schemas", "New App Data.xlsx")
OUT = os.path.join(ROOT, "output")

META_SHEET = "__AMS_META__"

# ----------------------------------------------------------------------------
# Rules
# ----------------------------------------------------------------------------

# Hard ownership: if the parent row is dropped, the child row must be dropped too.
CASCADE_FKS = [
    # (child_table, child_col, parent_table)
    ("booking_item", "booking_id", "booking"),
    ("direct_sale_item", "sale_id", "direct_sale"),
    ("booking_allocation", "sale_id", "direct_sale"),
    ("booking_allocation", "sale_item_id", "direct_sale_item"),
    ("booking_allocation", "booking_item_id", "booking_item"),
    ("grn_allocation", "sale_id", "direct_sale"),
    ("grn_allocation", "sale_item_id", "direct_sale_item"),
    ("grn_allocation", "grn_item_id", "grn_item"),
    ("grn_item", "grn_id", "grn"),
    ("material_return_item", "material_return_id", "material_return"),
    ("material_return", "payment_id", "payment"),
    ("waive_off", "payment_id", "payment"),
    ("delivery_rent", "sale_id", "direct_sale"),
    ("sale_delivery_persons", "sale_id", "direct_sale"),
    ("delivery_item", "delivery_id", "delivery"),
    ("follow_up_reminder", "pending_bill_id", "pending_bill"),
    ("follow_up_contact", "pending_bill_id", "pending_bill"),
    ("follow_up_contact", "reminder_id", "follow_up_reminder"),
    ("cash_flow_entry_audit", "entry_id", "cash_flow_entry"),
    ("cash_flow_subcategory", "category_id", "cash_flow_category"),
    ("cash_flow_reconciliation_audit", "reconciliation_id", "account_reconciliation"),
    ("import_history_entry", "import_job_id", "import_job"),
    ("material_return_item", "material_return_id", "material_return"),
]

# Soft references: if the target is gone, blank the pointer but KEEP the row.
SOFT_FKS = [
    ("direct_sale", "invoice_id", "invoice"),
    ("entry", "invoice_id", "invoice"),
    ("payment", "client_id", "client"),
    ("direct_sale_item", "grn_item_id", "grn_item"),
    ("direct_sale", "payment_account_id", "account"),
    ("payment", "payment_account_id", "account"),
    ("grn", "payment_account_id", "account"),
    ("grn", "supplier_id", "supplier"),
    ("supplier_payment", "supplier_id", "supplier"),
    ("supplier_payment", "payment_account_id", "account"),
    ("booking", "receive_in_account_id", "account"),
    ("account_transaction", "from_account_id", "account"),
    ("account_transaction", "to_account_id", "account"),
    ("account_transaction", "reconciliation_id", "account_reconciliation"),
    ("material", "category_id", "material_category"),
    ("client", "transferred_to_id", "client"),
    ("audit_log", "user_id", "user"),
    ("sale_delivery_persons", "delivery_person_id", "delivery_person"),
    ("delivery_person_payment", "delivery_person_id", "delivery_person"),
    ("delivery_person_payment", "sale_id", "direct_sale"),
    ("fbm_rental", "client_id", "fbm_client"),
    ("fbm_rental", "item_id", "fbm_rental_item"),
    ("import_job", "upload_id", "import_upload"),
]

# Polymorphic source pointers: (table, type_col, id_col, {type_value: parent_table})
POLY_FKS = [
    ("account_transaction", "source_type", "source_id",
     {"Payment": "payment", "SupplierPayment": "supplier_payment",
      "Booking": "booking", "DirectSale": "direct_sale"}),
    ("entry", "source_table", "source_id",
     {"direct_sale": "direct_sale", "booking": "booking", "payment": "payment"}),
    ("pending_bill", "source_table", "source_id",
     {"direct_sale": "direct_sale", "booking": "booking", "payment": "payment"}),
]

# Tables that reference a client by free text (code and/or name) instead of by id.
# (table, code_col_or_None, name_col_or_None, note_col_or_None)
CLIENT_TEXT_REFS = [
    ("invoice", "client_code", "client_name", "note"),
    ("pending_bill", "client_code", "client_name", "note"),
    ("direct_sale", "client_code", "client_name", "note"),
    ("entry", "client_code", "client", "note"),
    ("booking", None, "client_name", "note"),
    ("payment", None, "client_name", "note"),
    ("waive_off", "client_code", "client_name", "note"),
    ("material_return", None, "client_name", "note"),
]

# Columns that only exist in the new schema and get a safe, documented default.
NEW_COLUMN_DEFAULTS = {
    ("account", "account_status"): lambda row: "active" if truthy(row.get("is_active")) else "inactive",
    ("user", "access_mode"): lambda row: "read_write",
}

ORPHAN_CODE_PREFIX = "ORPH-"
NOTE_TAG = "[MIGRATION]"


def truthy(v):
    return v in (True, 1, "1", "true", "True", "TRUE", "yes", "Y")


def norm(v):
    return str(v).strip().upper() if v is not None and str(v).strip() != "" else None


def jsonable(v):
    if isinstance(v, (dt.datetime, dt.date, dt.time)):
        return v.isoformat()
    return v


# ----------------------------------------------------------------------------
# Load
# ----------------------------------------------------------------------------

def read_workbook(path):
    """-> {sheet: {'columns': [...], 'rows': [dict, ...]}} preserving sheet order."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = {}
    for ws in wb.worksheets:
        it = ws.iter_rows(values_only=True)
        header = next(it, None) or ()
        cols = [str(c) for c in header if c is not None]
        rows = []
        for r in it:
            if all(c is None for c in r):
                continue
            rows.append({c: jsonable(v) for c, v in zip(cols, r)})
        out[ws.title] = {"columns": cols, "rows": rows}
    wb.close()
    return out


# ----------------------------------------------------------------------------
# Migration
# ----------------------------------------------------------------------------

class Report:
    def __init__(self):
        self.void_dropped = Counter()
        self.cascade_dropped = Counter()
        self.cascade_detail = defaultdict(list)
        self.soft_nulled = Counter()
        self.orphan_rows = Counter()
        self.orphan_clients = []
        self.defaults_applied = Counter()
        self.notes = []


def migrate(legacy, schema):
    rep = Report()

    # --- 0. target = schema structure, zero rows -----------------------------
    target = {t: {"columns": list(v["columns"]), "rows": []} for t, v in schema.items()}

    # --- 1. work on a mutable copy of legacy --------------------------------
    data = {t: [dict(r) for r in v["rows"]] for t, v in legacy.items()}

    # --- 2. drop voided rows -------------------------------------------------
    kept_ids = {}
    for t, rows in data.items():
        cols = legacy[t]["columns"]
        if "is_void" in cols:
            keep, dropped = [], 0
            for r in rows:
                if truthy(r.get("is_void")):
                    dropped += 1
                else:
                    keep.append(r)
            if dropped:
                rep.void_dropped[t] = dropped
            data[t] = keep
    for t, rows in data.items():
        if rows and "id" in legacy[t]["columns"]:
            kept_ids[t] = {r["id"] for r in rows}
        else:
            kept_ids[t] = {r["id"] for r in rows if "id" in r}

    # --- 3. cascade: drop children of dropped parents (fixed point) ----------
    changed = True
    while changed:
        changed = False
        for child, col, parent in CASCADE_FKS:
            if child not in data or parent not in data:
                continue
            if col not in legacy[child]["columns"]:
                continue
            alive = kept_ids.get(parent, set())
            keep, dropped = [], 0
            for r in data[child]:
                v = r.get(col)
                if v is not None and v not in alive:
                    dropped += 1
                    rep.cascade_detail[child].append(
                        {"id": r.get("id"), "column": col, "missing_parent": f"{parent}#{v}"})
                else:
                    keep.append(r)
            if dropped:
                rep.cascade_dropped[child] += dropped
                data[child] = keep
                kept_ids[child] = {r["id"] for r in keep if "id" in r}
                changed = True

    # --- 4. null dangling soft references ------------------------------------
    for tbl, col, parent in SOFT_FKS:
        if tbl not in data or parent not in data or col not in legacy[tbl]["columns"]:
            continue
        alive = kept_ids.get(parent, set())
        n = 0
        for r in data[tbl]:
            v = r.get(col)
            if v is not None and v not in alive:
                r[col] = None
                n += 1
        if n:
            rep.soft_nulled[f"{tbl}.{col}"] = n

    for tbl, tcol, icol, mapping in POLY_FKS:
        if tbl not in data:
            continue
        n = 0
        for r in data[tbl]:
            parent = mapping.get(r.get(tcol))
            v = r.get(icol)
            if parent and v is not None and v not in kept_ids.get(parent, set()):
                r[icol] = None
                n += 1
        if n:
            rep.soft_nulled[f"{tbl}.{icol} (polymorphic)"] = n

    # --- 5. orphan clients ---------------------------------------------------
    clients = data.get("client", [])
    by_code = {r["code"]: r for r in clients if r.get("code")}
    by_name = {}
    for r in clients:
        k = norm(r.get("name"))
        if k and k not in by_name:
            by_name[k] = r

    # collect distinct unknown client identities
    orphan_keys = {}           # normalised key -> {'code':..,'name':..,'refs':Counter}
    for tbl, ccol, ncol, _note in CLIENT_TEXT_REFS:
        if tbl not in data:
            continue
        cols = legacy[tbl]["columns"]
        for r in data[tbl]:
            code = r.get(ccol) if ccol and ccol in cols else None
            name = r.get(ncol) if ncol and ncol in cols else None
            if code and code in by_code:
                continue
            if not code and norm(name) in by_name:
                continue
            if not code and not norm(name):
                continue  # nothing to attach to; left as-is
            key = (code or "", norm(name) or "")
            e = orphan_keys.setdefault(key, {"code": code, "name": name, "refs": Counter()})
            e["refs"][tbl] += 1

    next_client_id = max([r["id"] for r in clients if isinstance(r.get("id"), int)] or [0]) + 1
    next_code_n = 1
    orphan_lookup = {}         # key -> synthetic client row
    for i, (key, e) in enumerate(sorted(orphan_keys.items(), key=lambda kv: -sum(kv[1]["refs"].values())), 1):
        oname = f"Orphan{i}"
        ocode = f"{ORPHAN_CODE_PREFIX}{next_code_n:05d}"
        next_code_n += 1
        original = e["name"] or "(no name)"
        original_code = e["code"] or "(no code)"
        breakdown = ", ".join(f"{k}={v}" for k, v in sorted(e["refs"].items()))
        row = {c: None for c in legacy["client"]["columns"]}
        row.update({
            "id": next_client_id,
            "code": ocode,
            "name": oname,
            "category": "Orphan",
            "opening_balance": 0,
            "is_active": True,
            "require_manual_invoice": False,
            "page_notes": (f"{NOTE_TAG} ORPHAN RECOVERY. Original client name: \"{original}\". "
                           f"Original client code: {original_code}. "
                           f"Recovered legacy rows: {breakdown}. "
                           f"This client did not exist in the legacy `client` table; "
                           f"rename/merge it manually once identified."),
            "created_at": dt.datetime.now().isoformat(timespec="seconds"),
        })
        next_client_id += 1
        clients.append(row)
        orphan_lookup[key] = row
        rep.orphan_clients.append({
            "orphan_client": oname, "code": ocode,
            "original_name": original, "original_code": original_code,
            "rows_recovered": sum(e["refs"].values()), "breakdown": breakdown,
        })

    # re-point orphan rows
    for tbl, ccol, ncol, notecol in CLIENT_TEXT_REFS:
        if tbl not in data:
            continue
        cols = legacy[tbl]["columns"]
        for r in data[tbl]:
            code = r.get(ccol) if ccol and ccol in cols else None
            name = r.get(ncol) if ncol and ncol in cols else None
            if code and code in by_code:
                continue
            if not code and norm(name) in by_name:
                continue
            if not code and not norm(name):
                continue
            key = (code or "", norm(name) or "")
            oc = orphan_lookup.get(key)
            if not oc:
                continue
            original = name or code
            if ccol and ccol in cols:
                r[ccol] = oc["code"]
            if ncol and ncol in cols:
                r[ncol] = oc["name"]
            if notecol and notecol in cols:
                tag = f'{NOTE_TAG} Orphan record. Original client: "{original}".'
                r[notecol] = f"{r[notecol]} | {tag}" if r.get(notecol) else tag
            rep.orphan_rows[tbl] += 1

    data["client"] = clients

    # --- 6. project legacy rows into the new schema --------------------------
    for tname, tinfo in target.items():
        cols = tinfo["columns"]
        if tname == META_SHEET:
            continue
        src = data.get(tname)
        if src is None:
            continue
        for r in src:
            out = {}
            for c in cols:
                if c in r:
                    out[c] = r[c]
                else:
                    fn = NEW_COLUMN_DEFAULTS.get((tname, c))
                    if fn:
                        out[c] = fn(r)
                        rep.defaults_applied[f"{tname}.{c}"] += 1
                    else:
                        out[c] = None
            tinfo["rows"].append(out)

    # --- 7. meta -------------------------------------------------------------
    if META_SHEET in target:
        target[META_SHEET]["rows"] = [
            {"key": "export_kind", "value": "literal_all"},
            {"key": "exported_at", "value": dt.datetime.now().isoformat()},
            {"key": "scope", "value": "single_store"},
            {"key": "tenant_id", "value": None},
            {"key": "tenant_name", "value": None},
            {"key": "format_version", "value": "2026-04"},
            {"key": "source", "value": "Legacy Data28-08-2026_05-31PM.xlsx (migrated, voids removed, orphans recovered)"},
        ]

    # tables present only in the new schema (no legacy source) stay empty
    rep.notes.append(sorted(set(schema) - set(legacy) - {META_SHEET}))
    return target, rep, data


# ----------------------------------------------------------------------------
# Type inference + writers
# ----------------------------------------------------------------------------

def infer_types(table):
    types = {}
    for c in table["columns"]:
        kinds = set()
        for r in table["rows"]:
            v = r.get(c)
            if v is None or v == "":
                continue
            if isinstance(v, bool):
                kinds.add("BOOLEAN")
            elif isinstance(v, int):
                kinds.add("INTEGER")
            elif isinstance(v, float):
                kinds.add("REAL")
            else:
                kinds.add("TEXT")
        if not kinds:
            types[c] = "TEXT"
        elif kinds == {"BOOLEAN"}:
            types[c] = "BOOLEAN"
        elif kinds <= {"INTEGER", "BOOLEAN"}:
            types[c] = "INTEGER"
        elif kinds <= {"INTEGER", "REAL", "BOOLEAN"}:
            types[c] = "REAL"
        else:
            types[c] = "TEXT"
    return types


def write_xlsx(target, path, with_data=True):
    wb = openpyxl.Workbook(write_only=True)
    for tname, t in target.items():
        ws = wb.create_sheet(title=tname[:31])
        ws.append(t["columns"])
        if with_data:
            for r in t["rows"]:
                ws.append([r.get(c) for c in t["columns"]])
    wb.save(path)


def write_json(target, path):
    payload = {
        "generated_at": dt.datetime.now().isoformat(),
        "format_version": "2026-04",
        "tables": {t: {"columns": v["columns"], "row_count": len(v["rows"]), "rows": v["rows"]}
                   for t, v in target.items()},
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, default=str, separators=(",", ":"))


def write_ndjson(target, path):
    with open(path, "w", encoding="utf-8") as f:
        for t, v in target.items():
            for r in v["rows"]:
                f.write(json.dumps({"_table": t, **r}, ensure_ascii=False, default=str) + "\n")


def write_csvs(target, folder):
    os.makedirs(folder, exist_ok=True)
    for t, v in target.items():
        with open(os.path.join(folder, f"{t}.csv"), "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=v["columns"], extrasaction="ignore")
            w.writeheader()
            for r in v["rows"]:
                w.writerow({c: r.get(c) for c in v["columns"]})


def ddl(target):
    stmts = []
    for t, v in target.items():
        types = infer_types(v)
        cols = []
        for c in v["columns"]:
            ty = types[c]
            pk = " PRIMARY KEY" if c == "id" and ty in ("INTEGER",) else ""
            cols.append(f'  "{c}" {ty}{pk}')
        stmts.append(f'CREATE TABLE IF NOT EXISTS "{t}" (\n' + ",\n".join(cols) + "\n);")
    return stmts


def sql_lit(v):
    if v is None:
        return "NULL"
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, (int, float)):
        return repr(v)
    return "'" + str(v).replace("'", "''") + "'"


def write_sql(target, path, schema_path):
    stmts = ddl(target)
    with open(schema_path, "w", encoding="utf-8") as f:
        f.write("-- New App schema (structure only, no data)\n")
        f.write("-- Generated by tools/migrate.py\n\n")
        f.write("\n\n".join(stmts) + "\n")
    with open(path, "w", encoding="utf-8") as f:
        f.write("-- New App data, migrated from legacy export\n")
        f.write(f"-- Generated {dt.datetime.now().isoformat()} by tools/migrate.py\n")
        f.write("BEGIN TRANSACTION;\n\n")
        f.write("\n\n".join(stmts) + "\n\n")
        for t, v in target.items():
            if not v["rows"]:
                continue
            collist = ", ".join(f'"{c}"' for c in v["columns"])
            f.write(f'-- {t}: {len(v["rows"])} rows\n')
            for r in v["rows"]:
                vals = ", ".join(sql_lit(r.get(c)) for c in v["columns"])
                f.write(f'INSERT INTO "{t}" ({collist}) VALUES ({vals});\n')
            f.write("\n")
        f.write("COMMIT;\n")


def write_sqlite(target, path):
    if os.path.exists(path):
        os.remove(path)
    con = sqlite3.connect(path)
    cur = con.cursor()
    for s in ddl(target):
        cur.execute(s)
    for t, v in target.items():
        if not v["rows"]:
            continue
        collist = ", ".join(f'"{c}"' for c in v["columns"])
        ph = ", ".join("?" * len(v["columns"]))
        cur.executemany(
            f'INSERT INTO "{t}" ({collist}) VALUES ({ph})',
            [[r.get(c) for c in v["columns"]] for r in v["rows"]],
        )
    con.commit()
    cur.execute("VACUUM")
    con.close()


# ----------------------------------------------------------------------------

def main():
    print("Reading legacy ...")
    legacy = read_workbook(LEGACY_XLSX)
    print("Reading new-app schema ...")
    schema = read_workbook(SCHEMA_XLSX)

    target, rep, data = migrate(legacy, schema)

    os.makedirs(OUT, exist_ok=True)
    for sub in ("xlsx", "json", "csv", "sqlite", "sql"):
        os.makedirs(os.path.join(OUT, sub), exist_ok=True)

    print("Writing outputs ...")
    write_xlsx(target, os.path.join(OUT, "xlsx", "New App Data - MIGRATED.xlsx"), True)
    empty = {t: {"columns": v["columns"], "rows": []} for t, v in schema.items()}
    write_xlsx(empty, os.path.join(OUT, "xlsx", "New App Data - EMPTY SCHEMA.xlsx"), False)
    write_json(target, os.path.join(OUT, "json", "new_app_data.json"))
    write_ndjson(target, os.path.join(OUT, "json", "new_app_data.ndjson"))
    write_csvs(target, os.path.join(OUT, "csv"))
    write_sqlite(target, os.path.join(OUT, "sqlite", "new_app_data.db"))
    write_sql(target,
              os.path.join(OUT, "sql", "new_app_data.sql"),
              os.path.join(OUT, "sql", "new_app_schema.sql"))

    # orphan map csv
    with open(os.path.join(OUT, "orphan_client_map.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["orphan_client", "code", "original_name",
                                          "original_code", "rows_recovered", "breakdown"])
        w.writeheader()
        for r in rep.orphan_clients:
            w.writerow(r)

    # stats json for the report step
    stats = {
        "legacy_counts": {t: len(v["rows"]) for t, v in legacy.items()},
        "dummy_counts": {t: len(v["rows"]) for t, v in schema.items()},
        "migrated_counts": {t: len(v["rows"]) for t, v in target.items()},
        "void_dropped": dict(rep.void_dropped),
        "cascade_dropped": dict(rep.cascade_dropped),
        "cascade_detail": {k: v for k, v in rep.cascade_detail.items()},
        "soft_nulled": dict(rep.soft_nulled),
        "orphan_rows": dict(rep.orphan_rows),
        "orphan_clients": rep.orphan_clients,
        "defaults_applied": dict(rep.defaults_applied),
        "new_only_tables": rep.notes[0] if rep.notes else [],
        "legacy_columns": {t: v["columns"] for t, v in legacy.items()},
        "schema_columns": {t: v["columns"] for t, v in schema.items()},
    }
    with open(os.path.join(OUT, "migration_stats.json"), "w", encoding="utf-8") as f:
        json.dump(stats, f, indent=1, default=str)

    total_in = sum(len(v["rows"]) for t, v in legacy.items() if t != META_SHEET)
    total_out = sum(len(v["rows"]) for t, v in target.items() if t != META_SHEET)
    print(f"legacy rows      : {total_in}")
    print(f"migrated rows    : {total_out}")
    print(f"voided dropped   : {sum(rep.void_dropped.values())}")
    print(f"cascade dropped  : {sum(rep.cascade_dropped.values())}")
    print(f"orphan rows fixed: {sum(rep.orphan_rows.values())}")
    print(f"orphan clients   : {len(rep.orphan_clients)}")
    print(f"outputs in       : {OUT}")


if __name__ == "__main__":
    main()
