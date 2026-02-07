import io
import os
import re
import smtplib
import time
from email.message import EmailMessage
from datetime import datetime, date, timedelta
from decimal import Decimal
from typing import List, Dict, Any, Optional, Tuple

import psycopg2
from psycopg2.extras import RealDictCursor
from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse, HTMLResponse, RedirectResponse

from openpyxl import load_workbook


DATABASE_URL = os.getenv("DATABASE_URL")
if not DATABASE_URL:
    raise RuntimeError("DATABASE_URL env var is required")

TEMPLATE_PATH = os.getenv("TEMPLATE_PATH", "app/templates/inventory_template 2.03.xlsx")

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


# -----------------------------
# DB helpers
# -----------------------------
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


def get_excel_row_retry(
    excel_id: str, max_wait_s: float = 6.0, step_s: float = 0.25
) -> Optional[Dict[str, Any]]:
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
    expected = excel_row.get("export_token")
    if not expected or token != expected:
        raise HTTPException(status_code=403, detail="invalid token")


# -----------------------------
# Parsing helpers
# -----------------------------
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
    return True


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


# -----------------------------
# Classification helpers (N/M/C/D/F/O)
# -----------------------------
CLASS_KEYWORDS = {
    "N": ["narc", "narcotic", "narcotics"],
    "M": ["malar", "malaria"],
    "C": ["cool", "cold", "fridge", "refrigerat"],
    "D": ["danger", "dangerous", "dg", "hazard"],
    "F": ["female", "fem"],
    "O": ["oxygen", "o2"],
}


def parse_item_classification_flags(v: Any) -> Dict[str, bool]:
    """
    Supports values like:
      - "N"
      - "N,M"
      - "NMCDFO"
      - "malaria; cool; oxygen"
      - "['N','M']"
    Returns dict: { 'N': bool, 'M': bool, 'C': bool, 'D': bool, 'F': bool, 'O': bool }
    """
    flags = {k: False for k in ["N", "M", "C", "D", "F", "O"]}
    if v is None:
        return flags

    s = str(v).strip()
    if not s:
        return flags

    # normalize common list-like formats
    s2 = s.strip().strip("[](){}")
    s2 = s2.replace("'", "").replace('"', "")

    # direct letters
    for ch in s2.upper():
        if ch in flags:
            flags[ch] = True

    # keyword scan (case-insensitive)
    low = s2.lower()
    for letter, kws in CLASS_KEYWORDS.items():
        for kw in kws:
            if kw in low:
                flags[letter] = True

    return flags


# -----------------------------
# Data fetchers
# -----------------------------
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

    try:
        row2 = fetch_one(
            "SELECT to_jsonb(v) AS j FROM vessels v WHERE v.vessel_id=%s",
            (vessel_id,),
        )
        if row2 and row2.get("j"):
            return row2["j"]
    except Exception:
        pass

    return {}


def get_vessel_certificates_by_pack(vessel_id: str) -> List[Dict[str, Any]]:
    """One row per pack, latest end date per pack."""
    try:
        return fetch_all(
            """
            SELECT
              vc.pack_id,
              COALESCE(p.pack_name,'') AS pack_name,
              MAX(vc.certificate_end_date) AS certificate_end_date
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


def compute_next_resupply_date(cert_rows: List[Dict[str, Any]]) -> Optional[date]:
    today = date.today()
    candidates: List[date] = []
    for rr in cert_rows or []:
        end_d = parse_date_any(rr.get("certificate_end_date"))
        if end_d and end_d >= today:
            candidates.append(end_d - timedelta(days=30))
    return min(candidates) if candidates else None


def build_certificate_display(cert_rows: List[Dict[str, Any]]) -> str:
    today = date.today()
    lines: List[str] = []
    for rr in cert_rows or []:
        pack = (rr.get("pack_name") or "").strip()
        if not pack:
            pack = txt(rr.get("pack_id") or "").strip() or "Pack"
        end_d = parse_date_any(rr.get("certificate_end_date"))
        if end_d and end_d >= today:
            lines.append(f"{pack}: {end_d.strftime('%d-%m-%Y')}")
        else:
            lines.append(f"{pack}: Expired")
    return "\n".join(lines)


def get_export_rows(excel_id: str) -> List[Dict[str, Any]]:
    rows = fetch_all(
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

    # Skip quantity 0 rows (requested)
    out: List[Dict[str, Any]] = []
    for rr in rows:
        q = num(rr.get("vessel_item_quantity"))
        if q is not None and q == 0:
            continue
        out.append(rr)
    return out


def get_vessel_storage_map(vessel_id: str) -> Dict[str, str]:
    out: Dict[str, str] = {}

    # Preferred: vessel_storages table
    try:
        rows = fetch_all(
            """
            SELECT
              storage_id,
              COALESCE(NULLIF(storage_display,''), NULLIF(storage_name,''), '') AS storage_name
            FROM vessel_storages
            WHERE vessel_id = %s
            """,
            (vessel_id,),
        )
        for r in rows:
            sid = txt(r.get("storage_id")).strip()
            name = txt(r.get("storage_name")).strip()
            if sid and name:
                out[sid] = name
    except Exception:
        pass

    # Fallback: storages table
    if not out:
        try:
            rows = fetch_all(
                """
                SELECT
                  s.storage_id,
                  COALESCE(NULLIF(s.storage_display,''), NULLIF(s.storage_name,''), '') AS storage_name
                FROM storages s
                WHERE s.vessel_id = %s
                """,
                (vessel_id,),
            )
            for r in rows:
                sid = txt(r.get("storage_id")).strip()
                name = txt(r.get("storage_name")).strip()
                if sid and name:
                    out[sid] = name
        except Exception:
            pass

    return out


def get_vessel_storages_sorted(vessel_id: str) -> List[Tuple[str, str]]:
    m = get_vessel_storage_map(vessel_id)
    items = [(name, sid) for sid, name in m.items() if name]
    items.sort(key=lambda x: (x[0] or "").lower())
    # Ensure common fallback exists
    if not any((name or "").strip().lower() == "inbound storage" for name, _ in items):
        items.insert(0, ("Inbound storage", ""))
    return items


# -----------------------------
# Lookups (flag / category)
# -----------------------------
def lookup_vessel_category_name(cat_id_or_name: str) -> str:
    """Try to resolve category ID/code to a human name. Returns '' if not found."""
    if not cat_id_or_name:
        return ""

    s = str(cat_id_or_name).strip()

    candidates = [
        ("vessel_categories", "category_name", "vessel_category_id"),
        ("vessel_categories", "category_name", "category_id"),
        ("vessel_categories", "name", "vessel_category_id"),
        ("vessel_categories", "name", "category_id"),
        ("vessel_category", "category_name", "vessel_category_id"),
        ("vessel_category", "category_name", "category_id"),
        ("vw_vessel_categories", "category_name", "vessel_category_id"),
        ("vw_vessel_categories", "category_name", "category_id"),
    ]
    for table, namecol, idcol in candidates:
        try:
            row = fetch_one(
                f"SELECT {namecol} AS n FROM {table} WHERE {idcol}=%s LIMIT 1",
                (s,),
            )
            if row and row.get("n"):
                return str(row["n"])
        except Exception:
            continue
    return ""


def lookup_flag_name(flag_id_or_code_or_name: str) -> str:
    if not flag_id_or_code_or_name:
        return ""

    s = str(flag_id_or_code_or_name).strip()

    # If it already looks like a name, return it
    if " " in s and len(s) <= 80:
        return s

    queries = [
        ("SELECT entity_name AS n FROM entities WHERE entity_id=%s LIMIT 1", (s,)),
        ("SELECT entity_name AS n FROM entities WHERE entity_code=%s LIMIT 1", (s,)),
        ("SELECT entity_name AS n FROM entities WHERE code=%s LIMIT 1", (s,)),
        ("SELECT entity_name AS n FROM entity WHERE entity_id=%s LIMIT 1", (s,)),
        ("SELECT entity_name AS n FROM entity WHERE entity_code=%s LIMIT 1", (s,)),
        ("SELECT flag_name AS n FROM vessel_flags WHERE vessel_flag_id=%s LIMIT 1", (s,)),
        ("SELECT entity_name AS n FROM vessel_flags WHERE vessel_flag_id=%s LIMIT 1", (s,)),
        ("SELECT flag_name AS n FROM flags WHERE flag_id=%s LIMIT 1", (s,)),
        ("SELECT entity_name AS n FROM flags WHERE flag_id=%s LIMIT 1", (s,)),
    ]
    for sql, params in queries:
        try:
            row = fetch_one(sql, params)
            if row and row.get("n"):
                return str(row["n"])
        except Exception:
            continue

    # Fallback: show the code
    return s


# -----------------------------
# Excel filling
# -----------------------------
def build_filename(vessel: Dict[str, Any], vessel_id: str) -> str:
    raw_name = vessel.get("vessel_name") or ""
    raw_imo = vessel.get("vessel_IMO") or vessel.get("vessel_imo") or ""

    if vessel_id and (not raw_name or not raw_imo):
        try:
            row = fetch_one(
                'SELECT vessel_name, "vessel_IMO" AS imo FROM vessels WHERE vessel_id=%s',
                (vessel_id,),
            ) or {}
            raw_name = raw_name or (row.get("vessel_name") or "")
            raw_imo = raw_imo or (row.get("imo") or "")
        except Exception:
            pass

    vessel_name = str(raw_name or "Vessel").replace("/", "-").strip()
    imo = str(raw_imo or "").strip()
    date_str = datetime.now().strftime("%d-%m-%Y")

    if imo:
        return f"Medical Inventory List - {vessel_name} - {imo} - {date_str}.xlsx"
    return f"Medical Inventory List - {vessel_name} - {date_str}.xlsx"


def fill_storage_sheet(wb, vessel_id: str):
    ws = find_sheet_fuzzy(wb, STORAGE_SHEET)
    if ws is None:
        return

    # Clear existing list (keep headers)
    for r in range(2, MAX_ROWS + 1):
        ws.cell(r, 1).value = None
        ws.cell(r, 2).value = None

    storages = get_vessel_storages_sorted(vessel_id)
    for i, (name, sid) in enumerate(storages):
        r = 2 + i
        if r > MAX_ROWS + 1:
            break
        ws.cell(r, 1).value = name
        ws.cell(r, 2).value = sid


def fill_vessel_information_sheet(
    wb,
    vessel: Dict[str, Any],
    excel_row: Dict[str, Any],
    cert_rows: List[Dict[str, Any]],
    next_resupply: Optional[date],
    cool_items: Optional[bool],
):
    ws = find_sheet_fuzzy(wb, VESSEL_INFO_SHEET)
    if ws is None:
        ws = wb.create_sheet(VESSEL_INFO_SHEET)

    # Vessel fields
    company = vessel.get("company_name") or ""
    vname = vessel.get("vessel_name") or ""
    imo = vessel.get("vessel_IMO") or vessel.get("vessel_imo") or ""
    notes = vessel.get("vessel_notes") or ""

    # Emails
    purchasing = (
        vessel.get("purchasing_email")
        or vessel.get("purchasing_mail")
        or vessel.get("purchasing")
        or ""
    )
    if not purchasing:
        purchasing = vessel.get("vessel_contact_email") or vessel.get("email") or ""

    # Flag / category (prefer already-resolved names from enriched view)
    flag_val = vessel.get("vessel_flag_name") or vessel.get("vessel_flag") or vessel.get("flag_name") or ""
    flag_name = txt(flag_val).strip()
    if flag_name:
        flag_name = lookup_flag_name(flag_name)

    cat_val = vessel.get("vessel_category_name") or vessel.get("vessel_category") or vessel.get("category_name") or ""
    cat_name = txt(cat_val).strip()
    if cat_name:
        looked = lookup_vessel_category_name(cat_name)
        if looked:
            cat_name = looked

    malaria = yesno_or_blank(vessel.get("malaria_area"))
    mfag = yesno_or_blank(vessel.get("mfag"))
    female = yesno_or_blank(vessel.get("female_onboard"))
    dang = yesno_or_blank(vessel.get("dangerous_good"))
    narc = yesno_or_blank(vessel.get("narcotics"))
    oxygen = yesno_or_blank(vessel.get("medical_oxygen"))

    agreement = vessel.get("vessel_subscription_type") or ""

    rr = vessel.get("resupply_rate")
    em = vessel.get("expiration_months")

    # Cool Items: prefer vessel field if it exists, else computed from items
    cool_text = ""
    cool_field = vessel.get("cool_items")
    if cool_field is not None and cool_field != "":
        cool_text = yesno_or_blank(cool_field)
    elif cool_items is not None:
        cool_text = "Yes" if cool_items else "No"

    cert_display = build_certificate_display(cert_rows)

    # ---- Template cell mapping (inventory_template 2.03) ----
    # Row 4
    ws["A4"].value = txt(company)
    ws["C4"].value = txt(vname)
    ws["E4"].value = malaria
    ws["G4"].value = mfag

    # Next resupply date (I4)
    if next_resupply:
        ws["I4"].value = next_resupply
        ws["I4"].number_format = "dd-mm-yyyy"
    else:
        ws["I4"].value = ""

    # Certificates list (K4) - multi-line
    ws["K4"].value = cert_display

    # Notes (N4:N12 merged)
    ws["N4"].value = txt(notes)

    # Row 6
    ws["A6"].value = txt(vessel.get("vessel_contact_name") or "")
    ws["C6"].value = txt(imo)
    ws["E6"].value = female
    ws["G6"].value = cool_text
    ws["I6"].value = txt(agreement)

    # Row 8
    ws["A8"].value = txt(vessel.get("vessel_contact_email") or vessel.get("email") or "")
    ws["C8"].value = txt(flag_name)
    ws["E8"].value = dang
    ws["I8"].value = rr if rr not in (None, "") else ""

    # Row 10
    ws["A10"].value = txt(vessel.get("vessel_contact_phone") or "")
    ws["C10"].value = txt(vessel.get("vessel_crew_size") or "")
    ws["E10"].value = narc
    ws["I10"].value = em if em not in (None, "") else ""

    # Row 12
    ws["A12"].value = txt(purchasing)
    ws["C12"].value = txt(cat_name)
    ws["E12"].value = oxygen


def fill_inventory_sheet(ws, rows: List[Dict[str, Any]], storage_map: Dict[str, str]):
    # Map headers -> column index from row 1
    headers: Dict[str, int] = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if isinstance(v, str) and v.strip():
            headers[v.strip().lower()] = c

    def col(name: str) -> int:
        c = headers.get(name.lower())
        if not c:
            raise HTTPException(status_code=500, detail=f"Inventory template missing column: {name}")
        return c

    def col_opt(name: str) -> Optional[int]:
        return headers.get(name.lower())

    c_storage = col("Storage")
    c_article = col("Article No.")
    c_item = col("Item name")
    c_qty = col("Quantity")
    c_total = col("Total quantity")
    c_cert = col("Certificate qty")
    c_exp = col("Expiry date")
    c_law = col("Law code")
    c_pack = col("Pack name")

    c_n = col_opt("N")
    c_m = col_opt("M")
    c_c = col_opt("C")
    c_d = col_opt("D")
    c_f = col_opt("F")
    c_o = col_opt("O")

    # Clear rows
    clear_cols = [c_storage, c_article, c_item, c_qty, c_total, c_cert, c_exp, c_law, c_pack]
    for cc in [c_n, c_m, c_c, c_d, c_f, c_o]:
        if cc:
            clear_cols.append(cc)

    for r in range(2, MAX_ROWS + 1):
        for cidx in clear_cols:
            ws.cell(r, cidx).value = None

    # Fill
    for i, rr in enumerate(rows[:MAX_ROWS]):
        r = 2 + i

        storage = (rr.get("storage_display") or "").strip()
        if not storage:
            sid = txt(rr.get("storage_id") or "").strip()
            storage = (storage_map.get(sid) or "").strip()
        if not storage:
            storage = "Inbound storage"

        barcode = (rr.get("item_barcode") or "").strip() or "Extra"
        ws.cell(r, c_storage).value = storage
        ws.cell(r, c_article).value = barcode
        ws.cell(r, c_item).value = rr.get("vessel_item_name") or ""
        ws.cell(r, c_qty).value = num(rr.get("vessel_item_quantity")) or 0
        ws.cell(r, c_total).value = num(rr.get("totalitem_qty_sql")) or 0
        ws.cell(r, c_cert).value = num(rr.get("certificate_qty_sql")) or 0

        dte = parse_date_any(rr.get("vessel_item_expiration_date"))
        if dte:
            ws.cell(r, c_exp).value = dte
            ws.cell(r, c_exp).number_format = "mm-yyyy"
        else:
            ws.cell(r, c_exp).value = None

        ws.cell(r, c_law).value = rr.get("vessel_item_law_code") or ""
        ws.cell(r, c_pack).value = rr.get("pack_name") or ""

        flags = parse_item_classification_flags(rr.get("item_classification_export"))
        if c_n:
            ws.cell(r, c_n).value = "N" if flags.get("N") else ""
        if c_m:
            ws.cell(r, c_m).value = "M" if flags.get("M") else ""
        if c_c:
            ws.cell(r, c_c).value = "C" if flags.get("C") else ""
        if c_d:
            ws.cell(r, c_d).value = "D" if flags.get("D") else ""
        if c_f:
            ws.cell(r, c_f).value = "F" if flags.get("F") else ""
        if c_o:
            ws.cell(r, c_o).value = "O" if flags.get("O") else ""


def fill_upload_sheet(ws, rows: List[Dict[str, Any]], storage_map: Dict[str, str]):
    # Build header map from row 1
    headers: Dict[str, int] = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if isinstance(v, str) and v.strip():
            headers[v.strip().lower()] = c

    def setv(r: int, h: str, v):
        c = headers.get(h.lower())
        if c:
            ws.cell(r, c).value = v

    # Clear
    for r in range(2, MAX_ROWS + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None

    for i, rr in enumerate(rows[:MAX_ROWS]):
        r = 2 + i

        storage = (rr.get("storage_display") or "").strip()
        if not storage:
            sid = txt(rr.get("storage_id") or "").strip()
            storage = (storage_map.get(sid) or "").strip()
        if not storage:
            storage = "Inbound storage"

        barcode = (rr.get("item_barcode") or "").strip() or "123"
        dte = parse_date_any(rr.get("vessel_item_expiration_date"))

        setv(r, "vessel_item_id", rr.get("vessel_item_id") or "")
        setv(r, "vessel_id", rr.get("vessel_id") or "")
        setv(r, "storage_id", rr.get("storage_id") or "")
        setv(r, "storage_display", storage)
        setv(r, "item_id", rr.get("item_id") or "")
        setv(r, "pack_id", rr.get("pack_id") or "")
        setv(r, "category_id", rr.get("category_id") or "")
        setv(r, "item_barcode", barcode)
        setv(r, "vessel_item_name", rr.get("vessel_item_name") or "")
        setv(r, "vessel_item_law_code", rr.get("vessel_item_law_code") or "")
        setv(r, "vessel_item_quantity", num(rr.get("vessel_item_quantity")) or 0)
        setv(r, "totalitem_qty_sql", num(rr.get("totalitem_qty_sql")) or 0)
        setv(r, "certificate_qty_sql", num(rr.get("certificate_qty_sql")) or 0)
        setv(r, "pack_name", rr.get("pack_name") or "")

        if dte:
            setv(r, "vessel_item_expiration_date", dte)
            c = headers.get("vessel_item_expiration_date")
            if c:
                ws.cell(r, c).number_format = "dd-mm-yyyy"


def build_workbook_bytes(excel_row: Dict[str, Any], rows: List[Dict[str, Any]], vessel: Dict[str, Any]) -> bytes:
    wb = load_workbook(TEMPLATE_PATH)

    vessel_id = (excel_row or {}).get("vessel_id") or ""
    storage_map = get_vessel_storage_map(vessel_id) if vessel_id else {}

    # Fill Storage sheet list for dropdowns
    if vessel_id:
        fill_storage_sheet(wb, vessel_id)

    # Fill Inventory
    ws_inv = find_sheet_fuzzy(wb, INVENTORY_SHEET)
    if ws_inv is None:
        raise HTTPException(status_code=500, detail=f"Sheet '{INVENTORY_SHEET}' not found in template")
    fill_inventory_sheet(ws_inv, rows, storage_map)

    # Fill Upload (optional)
    ws_up = find_sheet_fuzzy(wb, UPLOAD_SHEET)
    if ws_up is not None:
        fill_upload_sheet(ws_up, rows, storage_map)

    # Certificates + next resupply + cool_items computed
    cert_rows = get_vessel_certificates_by_pack(vessel_id) if vessel_id else []
    next_resupply = compute_next_resupply_date(cert_rows)
    cool_items = None
    if rows:
        has_cool = any(parse_item_classification_flags(rr.get("item_classification_export")).get("C") for rr in rows)
        cool_items = has_cool

    fill_vessel_information_sheet(
        wb=wb,
        vessel=vessel,
        excel_row=excel_row,
        cert_rows=cert_rows,
        next_resupply=next_resupply,
        cool_items=cool_items,
    )

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# -----------------------------
# Routes
# -----------------------------
@app.get("/health")
def health():
    return {"ok": True}


@app.get("/download")
def download(excel_id: str, token: str):
    # Retry on "excel_id not found" race just after AppSheet Save
    excel_row = get_excel_row_retry(excel_id, max_wait_s=6.0, step_s=0.25)

    # If still not visible in DB, show a self-refreshing wait page (no JSON error)
    if not excel_row:
        return HTMLResponse(WAIT_HTML, status_code=200)

    validate_token(excel_row, token)

    # Server-side redirect to the real file endpoint (more reliable than JS/iframe)
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
    rows = get_export_rows(excel_id)

    content = build_workbook_bytes(excel_row, rows, vessel)
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
    rows = get_export_rows(excel_id)

    content = build_workbook_bytes(excel_row, rows, vessel)
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
