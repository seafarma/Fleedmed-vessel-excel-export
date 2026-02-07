import io
import os
import re
import smtplib
import time
import secrets
from email.message import EmailMessage
from datetime import datetime, date, timedelta
from decimal import Decimal
from typing import List, Dict, Any, Optional, Tuple

import psycopg2
from psycopg2.extras import RealDictCursor
from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse, HTMLResponse, RedirectResponse

from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation


DATABASE_URL = os.getenv("DATABASE_URL")
if not DATABASE_URL:
    raise RuntimeError("DATABASE_URL env var is required")

TEMPLATE_PATH = os.getenv("TEMPLATE_PATH", "app/templates/inventory_template.xlsx")

INVENTORY_SHEET = os.getenv("INVENTORY_SHEET", "Inventory")
VESSEL_INFO_SHEET = os.getenv("VESSEL_INFO_SHEET", "Vessel Information")
UPLOAD_SHEET = os.getenv("UPLOAD_SHEET", "Upload")
STORAGE_SHEET = os.getenv("STORAGE_SHEET", "Storage")

MAX_ROWS = int(os.getenv("MAX_ROWS", "3000"))

SMTP_HOST = os.getenv("SMTP_HOST", "")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER", "")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
SMTP_USE_TLS = os.getenv("SMTP_USE_TLS", "true").lower() == "true"
FROM_EMAIL = os.getenv("FROM_EMAIL", SMTP_USER or "no-reply@example.com")
FROM_NAME = os.getenv("FROM_NAME", "Inventory Export")

app = FastAPI()

WAIT_HTML = """
<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <title>Preparing download…</title>
    <meta http-equiv="refresh" content="1">
    <style>
      body { font-family: Arial, sans-serif; padding: 24px; }
    </style>
  </head>
  <body>
    Preparing your Excel file… this page will refresh automatically.
  </body>
</html>
"""

# Vessel Information layout from your template (confirmed)
# label cells:
# Next resupply label at I3 => write to I4
CELL_NEXT_RESUPPLY = "I4"
# Certificate list header at K3 => start values at K4/L4 downward
CERT_PACK_COL = "K"
CERT_STATUS_COL = "L"
CERT_START_ROW = 4
CERT_MAX_ROWS = 27  # K4:L30


def db_conn():
    return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)


def fetch_one(sql: str, params: tuple) -> Optional[Dict[str, Any]]:
    with db_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(sql, params)
            return cur.fetchone()


def fetch_all(sql: str, params: tuple) -> List[Dict[str, Any]]:
    with db_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(sql, params)
            return cur.fetchall()


def get_excel_row(excel_id: str) -> Dict[str, Any]:
    row = fetch_one(
        "SELECT excel_id, vessel_id, export_token FROM vessel_excel WHERE excel_id=%s",
        (excel_id,),
    )
    if not row:
        raise HTTPException(status_code=404, detail="excel_id not found")
    return row


def get_excel_row_retry(excel_id: str, max_wait_s: float = 6.0, step_s: float = 0.25) -> Optional[Dict[str, Any]]:
    deadline = time.time() + max_wait_s
    while True:
        row = fetch_one(
            "SELECT excel_id, vessel_id, export_token FROM vessel_excel WHERE excel_id=%s",
            (excel_id,),
        )
        if row:
            return row
        if time.time() >= deadline:
            return None
        time.sleep(step_s)


def validate_token(excel_row: Dict[str, Any], token: str):
    expected = (excel_row.get("export_token") or "").strip()
    tok = (token or "").strip()
    if not expected or not secrets.compare_digest(tok, expected):
        raise HTTPException(status_code=403, detail="invalid token")


def to_bool(v) -> Optional[bool]:
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v != 0
    s = str(v).strip().lower()
    if s in ("true", "t", "yes", "y", "1", "on"):
        return True
    if s in ("false", "f", "no", "n", "0", "off"):
        return False
    return None


def yesno_or_blank(v) -> str:
    b = to_bool(v)
    if b is None:
        return ""
    return "Yes" if b else "No"


def parse_date_any(v) -> Optional[date]:
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    try:
        if len(s) >= 10 and s[4] == "-" and s[7] == "-":
            return datetime.strptime(s[:10], "%Y-%m-%d").date()
    except Exception:
        pass
    try:
        if len(s) >= 10 and s[2] == "-" and s[5] == "-":
            return datetime.strptime(s[:10], "%d-%m-%Y").date()
    except Exception:
        pass
    return None


def txt(v) -> str:
    return "" if v is None else str(v)


def num(v) -> Optional[float]:
    if v is None or v == "":
        return None
    if isinstance(v, Decimal):
        return float(v)
    try:
        return float(v)
    except Exception:
        return None


def find_sheet_fuzzy(wb, wanted: str):
    w = wanted.strip().lower()
    for n in wb.sheetnames:
        if n.strip().lower() == w:
            return wb[n]
    for n in wb.sheetnames:
        if w in n.strip().lower():
            return wb[n]
    return None


def try_get_vessel_json(vessel_id: str) -> Dict[str, Any]:
    # Prefer enriched view; fallback vessels
    try:
        row = fetch_one(
            "SELECT to_jsonb(t) AS j FROM vw_vessels_enriched t WHERE t.vessel_id=%s",
            (vessel_id,),
        )
        if row and row.get("j"):
            return row["j"]
    except Exception:
        pass

    row2 = fetch_one(
        "SELECT to_jsonb(v) AS j FROM vessels v WHERE v.vessel_id=%s",
        (vessel_id,),
    )
    if row2 and row2.get("j"):
        return row2["j"]
    return {}


def merge_base_vessel(vessel_id: str, vessel: Dict[str, Any]) -> Dict[str, Any]:
    # Do NOT overwrite enriched vessel; fill gaps only
    try:
        row = fetch_one("SELECT to_jsonb(v) AS j FROM vessels v WHERE v.vessel_id=%s", (vessel_id,))
        if row and row.get("j"):
            base_vessel = row["j"]
            return {**base_vessel, **(vessel or {})}
    except Exception:
        pass
    return vessel or {}


def looks_like_hex_id(s: str) -> bool:
    s = (s or "").strip()
    return bool(re.fullmatch(r"[0-9a-fA-F]{8,}", s))


def lookup_flag_name(flag_val: str) -> str:
    """
    Fix: your vessel_flag can be AAA-0166 (entity_code), not a name.
    We must lookup entities.entity_code -> entities.entity_name.
    """
    if not flag_val:
        return ""
    s = str(flag_val).strip()

    # If it looks like an entity_code (AAA-0166), lookup by entity_code first
    if re.fullmatch(r"[A-Z]{3}-\d{4}", s):
        try:
            row = fetch_one("SELECT entity_name AS n FROM entities WHERE entity_code=%s LIMIT 1", (s,))
            if row and row.get("n"):
                return str(row["n"])
        except Exception:
            pass
        return s

    # If it is an ID-like value, try to resolve via a few tables
    if looks_like_hex_id(s):
        candidates = [
            ("entities", "entity_name", "entity_id"),
            ("entity", "entity_name", "entity_id"),
            ("vessel_flags", "entity_name", "vessel_flag_id"),
            ("flags", "entity_name", "flag_id"),
        ]
        for table, namecol, idcol in candidates:
            try:
                row = fetch_one(f"SELECT {namecol} AS n FROM {table} WHERE {idcol}=%s LIMIT 1", (s,))
                if row and row.get("n"):
                    return str(row["n"])
            except Exception:
                continue

    # Otherwise assume it is already a name
    return s


def lookup_vessel_category_name(cat_id: str) -> str:
    """
    Fix: vessel_category can be an ID that belongs to:
    - vessel_category / vessel_categories tables, OR
    - categories table (category_id -> third_category)
    """
    if not cat_id:
        return ""
    s = str(cat_id).strip()
    if not s:
        return ""

    # First: dedicated category tables
    candidates = [
        ("vessel_categories", "category_name", "vessel_category_id"),
        ("vessel_categories", "category_name", "category_id"),
        ("vessel_category", "category_name", "vessel_category_id"),
        ("vessel_category", "category_name", "category_id"),
    ]
    for table, namecol, idcol in candidates:
        try:
            row = fetch_one(f"SELECT {namecol} AS n FROM {table} WHERE {idcol}=%s LIMIT 1", (s,))
            if row and row.get("n"):
                return str(row["n"])
        except Exception:
            continue

    # Fallback: categories table
    try:
        row = fetch_one(
            "SELECT third_category AS n FROM categories WHERE category_id=%s LIMIT 1",
            (s,),
        )
        if row and row.get("n"):
            return str(row["n"])
    except Exception:
        pass

    return ""


def resolve_vessel_category(vessel: Dict[str, Any]) -> str:
    raw = (
        vessel.get("vessel_category_name")
        or vessel.get("vessel_category")
        or vessel.get("vessel_category_id")
        or ""
    )
    s = str(raw or "").strip()
    if not s:
        return ""
    if looks_like_hex_id(s):
        return lookup_vessel_category_name(s) or ""
    return s


def get_vessel_certificates_latest_by_pack(vessel_id: str) -> List[Dict[str, Any]]:
    try:
        return fetch_all(
            """
            SELECT
              vc.pack_id,
              p.pack_name,
              MAX(vc.certificate_end_date)::date AS last_certificate_end_date
            FROM vessel_certifications vc
            LEFT JOIN packs p ON p.pack_id = vc.pack_id
            WHERE vc.vessel_id = %s
            GROUP BY vc.pack_id, p.pack_name
            ORDER BY COALESCE(p.pack_name,'') ASC
            """,
            (vessel_id,),
        )
    except Exception:
        return []


def compute_next_resupply(cert_rows: List[Dict[str, Any]]) -> Optional[date]:
    """
    Fix: you said Next resupply is empty.
    New rule:
    - If there are valid certs (>= today): earliest (end-30) among valid
    - Else: use the latest end_date overall (even if expired) - 30
    This guarantees a date if any certificate exists.
    """
    today = date.today()
    valid_candidates: List[date] = []
    latest_end: Optional[date] = None

    for r in cert_rows or []:
        end_d = parse_date_any(r.get("last_certificate_end_date"))
        if not end_d:
            continue
        if latest_end is None or end_d > latest_end:
            latest_end = end_d
        if end_d >= today:
            valid_candidates.append(end_d - timedelta(days=30))

    if valid_candidates:
        return min(valid_candidates)

    if latest_end:
        return latest_end - timedelta(days=30)

    return None


def get_vessel_storages(vessel_id: str) -> List[Tuple[str, str]]:
    if not vessel_id:
        return []
    try:
        rows = fetch_all(
            """
            SELECT
              storage_id,
              COALESCE(
                NULLIF(storage_display,''),
                NULLIF(storage_name,''),
                NULLIF(storage_location,''),
                storage_id
              ) AS storage_name
            FROM vessel_storages
            WHERE vessel_id=%s
            """,
            (vessel_id,),
        )
    except Exception:
        return []

    storages: List[Tuple[str, str]] = []
    seen: set = set()
    for r in rows or []:
        sid = (r.get("storage_id") or "").strip()
        name = (r.get("storage_name") or "").strip()
        if not sid or sid in seen:
            continue
        seen.add(sid)
        storages.append((name or sid, sid))

    storages.sort(key=lambda x: (x[0] or "").lower())
    return storages


def apply_storage_dropdown(inv_ws, storage_sheet_name: str, n_storages: int):
    if n_storages <= 0:
        return
    end_row = 1 + n_storages  # header row at 1, values start at 2
    sheet_ref = "'" + storage_sheet_name.replace("'", "''") + "'"
    formula = f"={sheet_ref}!$A$2:$A${end_row}"

    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    dv.errorTitle = "Invalid storage"
    dv.error = "Select a storage from the list."
    dv.showErrorMessage = True

    inv_ws.add_data_validation(dv)
    # Storage column is A in your Inventory sheet
    dv.add(f"A2:A{MAX_ROWS}")


def fill_storage_sheet(wb, storages: List[Tuple[str, str]]):
    ws = find_sheet_fuzzy(wb, STORAGE_SHEET)
    if ws is None:
        ws = wb.create_sheet(STORAGE_SHEET)

    # Headers must be: Storage Name | storage_id
    ws["A1"].value = "Storage Name"
    ws["B1"].value = "storage_id"

    # Clear existing rows
    for r in range(2, 1000):
        ws.cell(r, 1).value = None
        ws.cell(r, 2).value = None

    # Fill sorted by name
    for i, (name, sid) in enumerate(storages, start=2):
        ws.cell(i, 1).value = name
        ws.cell(i, 2).value = sid

    return ws


def parse_item_classification_to_flags(v: Any) -> Tuple[bool, bool, bool, bool, bool, bool]:
    """
    Returns N, M, C, D, F, O flags from item_classification_export.
    Supports code strings: NMCDFO or N,M,C
    Supports keywords:
      - narc -> N
      - malar -> M
      - cool/cold/fridge -> C
      - danger/dg/hazmat -> D
      - fem -> F
      - oxygen/oxy/o2 -> O
    """
    s = "" if v is None else str(v).strip()
    if not s:
        return (False, False, False, False, False, False)

    s_up = s.upper()
    compact = re.sub(r"[\s,;/+\-|]+", "", s_up)

    if compact and re.fullmatch(r"[NMCDFO]{1,20}", compact):
        flags = set(compact)
        return ("N" in flags, "M" in flags, "C" in flags, "D" in flags, "F" in flags, "O" in flags)

    sl = s.lower()
    n = "narc" in sl
    m = "malar" in sl
    c = ("cool" in sl) or ("cold" in sl) or ("fridge" in sl)
    d = ("danger" in sl) or ("dg" in sl) or ("hazmat" in sl)
    f = "fem" in sl
    o = ("oxygen" in sl) or ("o2" in sl) or ("oxy" in sl) or (sl.strip() == "o")
    return (n, m, c, d, f, o)


def get_export_rows(excel_id: str) -> List[Dict[str, Any]]:
    return fetch_all(
        """
        SELECT
          v.vessel_item_id,
          v.vessel_id,
          v.storage_id,
          v.storage_display,
          v.item_id,
          v.pack_id,
          v.category_id,
          v.item_barcode,
          v.vessel_item_name,
          v.vessel_item_law_code,
          v.vessel_item_quantity,
          v.totalitem_qty_sql,
          v.certificate_qty_sql,
          v.vessel_item_expiration_date,
          v.pack_name,
          v.item_classification_export
        FROM vw_vessel_excel_items v
        WHERE v.excel_id=%s
        ORDER BY
          COALESCE(v.pack_name,'') ASC,
          COALESCE(v.vessel_item_law_code,'') ASC,
          COALESCE(v.storage_display,'') ASC,
          COALESCE(v.vessel_item_name,'') ASC
        """,
        (excel_id,),
    )


def fill_inventory_sheet(inv_ws, rows: List[Dict[str, Any]]):
    # map headers (row 1)
    headers: Dict[str, int] = {}
    for c in range(1, inv_ws.max_column + 1):
        v = inv_ws.cell(1, c).value
        if isinstance(v, str) and v.strip():
            headers[v.strip().lower()] = c

    def col(name: str) -> int:
        c = headers.get(name.lower())
        if not c:
            raise HTTPException(status_code=500, detail=f"Inventory template missing column: {name}")
        return c

    c_storage = col("Storage")
    c_article = col("Article No.")
    c_item = col("Item name")
    c_qty = col("Quantity")
    c_total = col("Total quantity")
    c_cert = col("Certificate qty")
    c_exp = col("Expiry date")
    c_law = col("Law code")
    c_pack = col("Pack name")

    # required flags in your template
    c_n = col("N")
    c_m = col("M")
    c_c = col("C")
    c_d = col("D")
    c_f = col("F")
    c_o = col("O")

    # clear existing rows (keep formatting)
    for r in range(2, MAX_ROWS + 1):
        for cidx in (c_storage, c_article, c_item, c_qty, c_total, c_cert, c_exp, c_law, c_pack, c_n, c_m, c_c, c_d, c_f, c_o):
            inv_ws.cell(r, cidx).value = None

    out_r = 2
    for rr in rows:
        if out_r > MAX_ROWS:
            break

        inv_ws.cell(out_r, c_storage).value = (rr.get("storage_display") or "").strip() or "Inbound storage"
        inv_ws.cell(out_r, c_article).value = (rr.get("item_barcode") or "").strip() or "Extra"
        inv_ws.cell(out_r, c_item).value = rr.get("vessel_item_name") or ""

        q = num(rr.get("vessel_item_quantity")) or 0
        inv_ws.cell(out_r, c_qty).value = q
        inv_ws.cell(out_r, c_total).value = num(rr.get("totalitem_qty_sql")) or 0
        inv_ws.cell(out_r, c_cert).value = num(rr.get("certificate_qty_sql")) or 0

        dte = parse_date_any(rr.get("vessel_item_expiration_date"))
        if dte:
            inv_ws.cell(out_r, c_exp).value = dte
            inv_ws.cell(out_r, c_exp).number_format = "mm-yyyy"
        else:
            inv_ws.cell(out_r, c_exp).value = None

        inv_ws.cell(out_r, c_law).value = rr.get("vessel_item_law_code") or ""
        inv_ws.cell(out_r, c_pack).value = rr.get("pack_name") or ""

        n, m, c, d, f, o = parse_item_classification_to_flags(rr.get("item_classification_export"))
        inv_ws.cell(out_r, c_n).value = "N" if n else None
        inv_ws.cell(out_r, c_m).value = "M" if m else None
        inv_ws.cell(out_r, c_c).value = "C" if c else None
        inv_ws.cell(out_r, c_d).value = "D" if d else None
        inv_ws.cell(out_r, c_f).value = "F" if f else None
        inv_ws.cell(out_r, c_o).value = "O" if o else None

        out_r += 1


def fill_upload_sheet(up_ws, rows: List[Dict[str, Any]]):
    """
    Upload tab: item_classification_export is removed (do not write it).
    """
    headers: Dict[str, int] = {}
    for c in range(1, up_ws.max_column + 1):
        v = up_ws.cell(1, c).value
        if isinstance(v, str) and v.strip():
            headers[v.strip().lower()] = c

    def setv(r: int, h: str, v):
        c = headers.get(h.lower())
        if c:
            up_ws.cell(r, c).value = v

    # clear previous values
    for r in range(2, MAX_ROWS + 1):
        for c in range(1, up_ws.max_column + 1):
            up_ws.cell(r, c).value = None

    out_r = 2
    for rr in rows:
        if out_r > MAX_ROWS:
            break

        setv(out_r, "vessel_item_id", rr.get("vessel_item_id") or "")
        setv(out_r, "vessel_id", rr.get("vessel_id") or "")
        setv(out_r, "storage_id", rr.get("storage_id") or "")
        setv(out_r, "item_id", rr.get("item_id") or "")
        setv(out_r, "pack_id", rr.get("pack_id") or "")
        setv(out_r, "category_id", rr.get("category_id") or "")

        storage_disp = (rr.get("storage_display") or "").strip() or "Inbound storage"
        setv(out_r, "storage_display", storage_disp)

        setv(out_r, "item_barcode", (rr.get("item_barcode") or "").strip() or "123")
        setv(out_r, "vessel_item_name", rr.get("vessel_item_name") or "")
        setv(out_r, "vessel_item_quantity", num(rr.get("vessel_item_quantity")) or 0)
        setv(out_r, "totalitem_qty_sql", num(rr.get("totalitem_qty_sql")) or 0)
        setv(out_r, "certificate_qty_sql", num(rr.get("certificate_qty_sql")) or 0)

        dte = parse_date_any(rr.get("vessel_item_expiration_date"))
        setv(out_r, "vessel_item_expiration_date", dte or "")
        if dte:
            c = headers.get("vessel_item_expiration_date")
            if c:
                up_ws.cell(out_r, c).number_format = "dd-mm-yyyy"

        setv(out_r, "vessel_item_law_code", rr.get("vessel_item_law_code") or "")
        setv(out_r, "pack_name", rr.get("pack_name") or "")

        out_r += 1


def fill_vessel_information_sheet(wb, vessel: Dict[str, Any], excel_row: Dict[str, Any], cert_rows: List[Dict[str, Any]]):
    ws = find_sheet_fuzzy(wb, VESSEL_INFO_SHEET)
    if ws is None:
        ws = wb.create_sheet(VESSEL_INFO_SHEET)

    vessel_id = (excel_row or {}).get("vessel_id") or ""
    vessel = merge_base_vessel(vessel_id, vessel)

    # left blocks (values are under label row+1 in your template)
    ws["A4"].value = txt(vessel.get("company_name") or "")
    ws["C4"].value = txt(vessel.get("vessel_name") or "")
    ws["E4"].value = yesno_or_blank(vessel.get("malaria_area"))
    ws["G4"].value = yesno_or_blank(vessel.get("mfag"))
    ws["N4"].value = txt(vessel.get("vessel_notes") or "")

    ws["A8"].value = txt(vessel.get("vessel_contact_email") or vessel.get("email") or "")
    ws["C8"].value = txt(lookup_flag_name(vessel.get("vessel_flag_name") or vessel.get("vessel_flag") or ""))
    ws["C12"].value = txt(resolve_vessel_category(vessel))

    # Next resupply (I4)
    next_resupply = compute_next_resupply(cert_rows)
    if next_resupply:
        ws[CELL_NEXT_RESUPPLY].value = next_resupply
        ws[CELL_NEXT_RESUPPLY].number_format = "dd-mm-yyyy"
    else:
        ws[CELL_NEXT_RESUPPLY].value = ""

    # Certificate list (K4/L4)
    for r in range(CERT_START_ROW, CERT_START_ROW + CERT_MAX_ROWS):
        ws[f"{CERT_PACK_COL}{r}"].value = None
        ws[f"{CERT_STATUS_COL}{r}"].value = None
        ws[f"{CERT_STATUS_COL}{r}"].number_format = "General"

    today = date.today()
    for idx, cr in enumerate(cert_rows[:CERT_MAX_ROWS]):
        r = CERT_START_ROW + idx
        pack_name = (cr.get("pack_name") or "").strip() or (cr.get("pack_id") or "")
        end_d = parse_date_any(cr.get("last_certificate_end_date"))

        ws[f"{CERT_PACK_COL}{r}"].value = txt(pack_name)
        if end_d and end_d >= today:
            ws[f"{CERT_STATUS_COL}{r}"].value = end_d
            ws[f"{CERT_STATUS_COL}{r}"].number_format = "dd-mm-yyyy"
        elif end_d:
            ws[f"{CERT_STATUS_COL}{r}"].value = "Expired"
        else:
            ws[f"{CERT_STATUS_COL}{r}"].value = ""


def build_workbook_bytes(
    excel_row: Dict[str, Any],
    rows: List[Dict[str, Any]],
    vessel: Dict[str, Any],
    cert_rows: List[Dict[str, Any]],
    storages: List[Tuple[str, str]],
) -> bytes:
    wb = load_workbook(TEMPLATE_PATH)

    # Storage sheet + dropdown range
    ws_storage = fill_storage_sheet(wb, storages)

    ws_inv = find_sheet_fuzzy(wb, INVENTORY_SHEET)
    if ws_inv is None:
        raise HTTPException(status_code=500, detail=f"Sheet '{INVENTORY_SHEET}' not found in template")
    fill_inventory_sheet(ws_inv, rows)

    # Apply dropdown to Inventory storage column
    apply_storage_dropdown(ws_inv, ws_storage.title, len(storages))

    ws_up = find_sheet_fuzzy(wb, UPLOAD_SHEET)
    if ws_up is not None:
        fill_upload_sheet(ws_up, rows)

    fill_vessel_information_sheet(wb, vessel, excel_row, cert_rows)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


@app.get("/health")
def health():
    return {"ok": True}


@app.get("/download")
def download(excel_id: str, token: str):
    excel_row = get_excel_row_retry(excel_id, max_wait_s=6.0, step_s=0.25)
    if not excel_row:
        return HTMLResponse(WAIT_HTML, status_code=200)

    validate_token(excel_row, token)

    return RedirectResponse(
        url=f"/download_file?excel_id={excel_id}&token={token}",
        status_code=302,
    )


@app.get("/download_file")
def download_file(excel_id: str, token: str):
    excel_row = get_excel_row_retry(excel_id, max_wait_s=6.0, step_s=0.25)
    if not excel_row:
        raise HTTPException(status_code=404, detail="excel_id not found")

    validate_token(excel_row, token)

    vessel_id = excel_row["vessel_id"]
    vessel = try_get_vessel_json(vessel_id)

    cert_rows = get_vessel_certificates_latest_by_pack(vessel_id)

    storages = get_vessel_storages(vessel_id)
    storage_map = {sid: name for (name, sid) in storages if sid and name}

    rows = get_export_rows(excel_id)

    # Fill missing storage_display using storage_id mapping
    if storage_map:
        for rr in rows:
            if not (rr.get("storage_display") or "").strip():
                sid = (rr.get("storage_id") or "").strip()
                if sid and storage_map.get(sid):
                    rr["storage_display"] = storage_map[sid]

    # Skip qty=0 in both tabs
    filtered_rows: List[Dict[str, Any]] = []
    for rr in rows:
        q = num(rr.get("vessel_item_quantity"))
        q = 0 if q is None else q
        if q == 0:
            continue
        filtered_rows.append(rr)

    content = build_workbook_bytes(excel_row, filtered_rows, vessel, cert_rows, storages)

    filename = build_filename(vessel, vessel_id)
    safe_name = filename.replace('"', "").replace("\n", " ").replace("\r", " ")

    headers = {
        "Content-Disposition": f'attachment; filename="{safe_name}"',
        "Cache-Control": "no-store",
    }
    return StreamingResponse(
        io.BytesIO(content),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.get("/email")
def email(excel_id: str, token: str, to_email: str):
    excel_row = get_excel_row_retry(excel_id, max_wait_s=6.0, step_s=0.25)
    if not excel_row:
        raise HTTPException(status_code=404, detail="excel_id not found")

    validate_token(excel_row, token)

    vessel_id = excel_row["vessel_id"]
    vessel = try_get_vessel_json(vessel_id)
    cert_rows = get_vessel_certificates_latest_by_pack(vessel_id)

    storages = get_vessel_storages(vessel_id)
    storage_map = {sid: name for (name, sid) in storages if sid and name}

    rows = get_export_rows(excel_id)
    if storage_map:
        for rr in rows:
            if not (rr.get("storage_display") or "").strip():
                sid = (rr.get("storage_id") or "").strip()
                if sid and storage_map.get(sid):
                    rr["storage_display"] = storage_map[sid]

    filtered_rows: List[Dict[str, Any]] = []
    for rr in rows:
        q = num(rr.get("vessel_item_quantity"))
        q = 0 if q is None else q
        if q == 0:
            continue
        filtered_rows.append(rr)

    content = build_workbook_bytes(excel_row, filtered_rows, vessel, cert_rows, storages)
    filename = build_filename(vessel, vessel_id)

    subject = f"Medical Inventory List - {vessel.get('vessel_name','')}".strip()
    body = "Please find the medical inventory list attached."

    msg = EmailMessage()
    msg["From"] = f"{FROM_NAME} <{FROM_EMAIL}>"
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.set_content(body)
    msg.add_attachment(
        content,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename,
    )

    if SMTP_USE_TLS:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
            s.starttls()
            if SMTP_USER:
                s.login(SMTP_USER, SMTP_PASSWORD)
            s.send_message(msg)
    else:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
            if SMTP_USER:
                s.login(SMTP_USER, SMTP_PASSWORD)
            s.send_message(msg)

    return {"ok": True, "sent_to": to_email}


@app.get("/")
def index():
    return HTMLResponse("<h3>Vessel Excel Export</h3><p>Use /download?excel_id=...&token=...</p>")
