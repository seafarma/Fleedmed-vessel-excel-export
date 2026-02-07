import io
import os
import re
import smtplib
import time
import secrets
import json
from email.message import EmailMessage
from datetime import datetime, date, timedelta
from decimal import Decimal
from typing import List, Dict, Any, Optional, Set, Tuple

import psycopg2
from psycopg2.extras import RealDictCursor
from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse, HTMLResponse, RedirectResponse

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell


# ==============================
# Config
# ==============================

DATABASE_URL = os.getenv("DATABASE_URL")
if not DATABASE_URL:
    raise RuntimeError("DATABASE_URL env var is required")

TEMPLATE_PATH = os.getenv("TEMPLATE_PATH", "app/templates/inventory_template 2.05.xlsx")

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

# Inventory classification columns in template
CLASS_LETTERS = ("N", "M", "C", "D", "F", "O")

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


# ==============================
# DB helpers
# ==============================

def db_conn():
    return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)


def fetch_one(sql: str, params: tuple = ()) -> Optional[Dict[str, Any]]:
    with db_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(sql, params)
            return cur.fetchone()


def fetch_all(sql: str, params: tuple = ()) -> List[Dict[str, Any]]:
    with db_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(sql, params)
            return cur.fetchall()


# ==============================
# Generic helpers
# ==============================

_UUID_RE = re.compile(r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")
_SHORT_ID_RE = re.compile(r"^[0-9a-fA-F]{8}$")
_FLAG_CODE_RE = re.compile(r"^[A-Z]{3}-\d+$")


def looks_like_uuid(v: Any) -> bool:
    if not v:
        return False
    return bool(_UUID_RE.match(str(v).strip()))


def looks_like_short_id(v: Any) -> bool:
    if not v:
        return False
    return bool(_SHORT_ID_RE.match(str(v).strip()))


def to_bool(v) -> Optional[bool]:
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float, Decimal)):
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


def safe_set(ws, addr: str, value, number_format: Optional[str] = None):
    cell = ws[addr]
    if isinstance(cell, MergedCell):
        for r in ws.merged_cells.ranges:
            if addr in r:
                tl = ws.cell(row=r.min_row, column=r.min_col)
                tl.value = value
                if number_format and isinstance(value, (date, datetime)):
                    tl.number_format = number_format
                return
        return

    cell.value = value
    if number_format and isinstance(value, (date, datetime)):
        cell.number_format = number_format


# ==============================
# Item type -> letters
# ==============================

def letters_from_item_type(item_type: Any) -> Set[str]:
    """
    Your requirement: Excel should show C if the item is a Cool Good.
    item_type comes from AppSheet EnumList, e.g. "🔵Cool Good,🟣Female Gender"
    """
    s = str(item_type or "").lower()
    out: Set[str] = set()
    if "cool good" in s:
        out.add("C")
     if "narcotic" in s:
        out.add("N")
    if "dangerous" in s:
        out.add("D")
    if "female" in s:
        out.add("F")
    if "medical oxygen" in s or re.search(r"\boxygen\b", s):
        out.add("O")
    if "malaria" in s:
        out.add("M")

    return out
# ==============================
# excel_id lookup + token
# ==============================

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
    expected = (excel_row or {}).get("export_token") or ""
    if not expected or not secrets.compare_digest(str(expected), str(token or "")):
        raise HTTPException(status_code=403, detail="invalid token")


# ==============================
# Data retrieval
# ==============================

def _json_as_dict(v: Any) -> Dict[str, Any]:
    if not v:
        return {}
    if isinstance(v, dict):
        return v
    if isinstance(v, str):
        try:
            x = json.loads(v)
            return x if isinstance(x, dict) else {}
        except Exception:
            return {}
    return {}


def try_get_vessel_json(vessel_id: str) -> Dict[str, Any]:
    # vw_vessels_enriched already includes vessel_flag_name + vessel_category_name
    try:
        row = fetch_one(
            "SELECT to_jsonb(t) AS j FROM vw_vessels_enriched t WHERE t.vessel_id=%s",
            (vessel_id,),
        )
        j = _json_as_dict((row or {}).get("j"))
        if j:
            return j
    except Exception:
        pass

    try:
        row2 = fetch_one(
            "SELECT to_jsonb(v) AS j FROM vessels v WHERE v.vessel_id=%s",
            (vessel_id,),
        )
        j2 = _json_as_dict((row2 or {}).get("j"))
        if j2:
            return j2
    except Exception:
        pass

    return {}


def get_vessel_core_fields(vessel_id: str) -> Dict[str, Any]:
    try:
        row = fetch_one(
            'SELECT vessel_name, "vessel_IMO" AS vessel_imo, vessel_contact_name, vessel_contact_email, vessel_contact_phone, vessel_crew_size, vessel_notes, purchasing_email FROM vessels WHERE vessel_id=%s',
            (vessel_id,),
        )
        return row or {}
    except Exception:
        return {}


def get_latest_certificates_by_pack(vessel_id: str) -> List[Dict[str, Any]]:
    try:
        return fetch_all(
            """
            SELECT
              vc.pack_id,
              COALESCE(p.pack_name, '') AS pack_name,
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


def compute_next_resupply_date(certs: List[Dict[str, Any]]) -> Optional[date]:
    today = date.today()
    candidates: List[date] = []
    for c in certs or []:
        end = parse_date_any(c.get("certificate_end_date"))
        if end and end >= today:
            candidates.append(end - timedelta(days=30))
    return min(candidates) if candidates else None


def get_vessel_storages(vessel_id: str) -> List[Dict[str, Any]]:
    candidates: List[Tuple[str, Tuple[Any, ...]]] = [
        (
            """
            SELECT
              storage_id,
              COALESCE(storage_name, storage_display, '') AS storage_name
            FROM vessel_storages
            WHERE vessel_id=%s
            ORDER BY COALESCE(storage_name, storage_display, '') ASC
            """,
            (vessel_id,),
        ),
        (
            """
            SELECT
              storage_id,
              COALESCE(storage_display, '') AS storage_name
            FROM vessel_storages
            WHERE vessel_id=%s
            ORDER BY COALESCE(storage_display,'') ASC
            """,
            (vessel_id,),
        ),
    ]
    for sql, params in candidates:
        try:
            rows = fetch_all(sql, params)
            out = []
            for r in rows:
                out.append(
                    {
                        "storage_id": r.get("storage_id") or "",
                        "storage_name": (r.get("storage_name") or "").strip(),
                    }
                )
            out = [r for r in out if r["storage_name"]]
            return out
        except Exception:
            continue
    return []


def get_export_rows(excel_id: str) -> List[Dict[str, Any]]:
    sql = """
        SELECT
          v.vessel_item_id,
          v.vessel_id,
          v.storage_id,
          v.storage_display,
          v.item_id,
          v.pack_id,
          v.category_id,
          v.item_barcode,
          v.pack_name,
          v.vessel_item_name,
          v.vessel_item_law_code,
          v.vessel_item_quantity,
          v.vessel_item_expiration_date,
          v.item_status,
          v.totalitem_qty_sql,
          v.certificate_qty_sql,
          v.item_classification_export,
          v.item_type
        FROM vw_vessel_excel_items v
        WHERE v.excel_id=%s
        ORDER BY
          COALESCE(v.pack_name,'') ASC,
          COALESCE(v.vessel_item_law_code,'') ASC,
          COALESCE(v.storage_display,'') ASC,
          COALESCE(v.vessel_item_name,'') ASC
    """
    rows = fetch_all(sql, (excel_id,))
    for r in rows:
        if "item_type" not in r:
            r["item_type"] = ""
    return rows


# ==============================
# Excel writing
# ==============================

def build_filename(vessel: Dict[str, Any], vessel_id: str) -> str:
    raw_name = (vessel or {}).get("vessel_name") or ""
    raw_imo = (vessel or {}).get("vessel_IMO") or (vessel or {}).get("vessel_imo") or ""

    if vessel_id and (not raw_name or not raw_imo):
        try:
            row = fetch_one('SELECT vessel_name, "vessel_IMO" AS imo FROM vessels WHERE vessel_id=%s', (vessel_id,)) or {}
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


def fill_storage_sheet(wb, storages: List[Dict[str, Any]]):
    ws = find_sheet_fuzzy(wb, STORAGE_SHEET)
    if ws is None:
        return

    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if isinstance(v, str) and v.strip():
            headers[v.strip().lower()] = c

    c_name = headers.get("storage name") or headers.get("storage") or 1
    c_id = headers.get("storage_id") or headers.get("storage id") or 2

    for r in range(2, 2001):
        ws.cell(r, c_name).value = None
        ws.cell(r, c_id).value = None

    rows = sorted(storages or [], key=lambda x: (x.get("storage_name") or "").lower())
    for i, s in enumerate(rows[:2000]):
        r = 2 + i
        ws.cell(r, c_name).value = s.get("storage_name") or ""
        ws.cell(r, c_id).value = s.get("storage_id") or ""


def fill_vessel_information_sheet(
    wb,
    vessel: Dict[str, Any],
    vessel_id: str,
    certs_by_pack: List[Dict[str, Any]],
    next_resupply: Optional[date],
    rows_for_derived: List[Dict[str, Any]],
):
    ws = find_sheet_fuzzy(wb, VESSEL_INFO_SHEET)
    if ws is None:
        ws = wb.create_sheet(VESSEL_INFO_SHEET)

    core = get_vessel_core_fields(vessel_id) if vessel_id else {}

    def vget(*keys, default=""):
        for k in keys:
            if k in (vessel or {}) and (vessel or {}).get(k) not in (None, ""):
                return (vessel or {}).get(k)
        for k in keys:
            if k in core and core.get(k) not in (None, ""):
                return core.get(k)
        return default

    company = vget("company_name", "company")
    vname = vget("vessel_name")
    imo = vget("vessel_IMO", "vessel_imo", "vessel_imo")
    notes = vget("vessel_notes", "notes")

    purchasing = vget("purchasing_email", "purchasing_mail", "purchasing")
    if not purchasing:
        purchasing = vget("vessel_contact_email", "email")

    # use enriched names directly
    flag_name = vget("vessel_flag_name", default="")
    if not flag_name:
        flag_name = vget("vessel_flag", default="")

    cat_name = vget("vessel_category_name", default="")
    if not cat_name:
        cat_name = vget("vessel_category", default="")

    malaria = yesno_or_blank(vget("malaria_area", "malaria_medicine", default=None))
    mfag = yesno_or_blank(vget("mfag", default=None))
    female = yesno_or_blank(vget("female_onboard", default=None))
    dang = yesno_or_blank(vget("dangerous_good", default=None))
    narc = yesno_or_blank(vget("narcotics", default=None))
    oxygen = yesno_or_blank(vget("medical_oxygen", default=None))

    cool_v = vget("cool_goods", "cool_items", default=None)
    cool = yesno_or_blank(cool_v)
    if cool == "" and rows_for_derived:
        has_cool = any("cool good" in str(r.get("item_type") or "").lower() for r in rows_for_derived)
        if has_cool:
            cool = "Yes"

    agreement = vget("vessel_subscription_type", "resupply_agreement", "subscription_type")
    rr = vget("resupply_rate", default="")
    em = vget("expiration_months", default="")

    safe_set(ws, "A4", txt(company))
    safe_set(ws, "C4", txt(vname))
    safe_set(ws, "E4", malaria)
    safe_set(ws, "G4", mfag)

    if next_resupply:
        safe_set(ws, "I4", next_resupply, number_format="dd-mm-yyyy")
    else:
        safe_set(ws, "I4", "")

    safe_set(ws, "N4", txt(notes))

    safe_set(ws, "A6", txt(vget("vessel_contact_name")))
    safe_set(ws, "C6", txt(imo))
    safe_set(ws, "E6", female)
    safe_set(ws, "G6", cool)
    safe_set(ws, "I6", txt(agreement))

    safe_set(ws, "A8", txt(vget("vessel_contact_email", "email")))
    safe_set(ws, "C8", txt(flag_name))
    safe_set(ws, "E8", dang)
    safe_set(ws, "I8", rr if rr not in (None, "") else "")

    safe_set(ws, "A10", txt(vget("vessel_contact_phone")))
    safe_set(ws, "C10", txt(vget("vessel_crew_size")))
    safe_set(ws, "E10", narc)

    # months value cell is I11 (label is in I9)
    safe_set(ws, "I11", em if em not in (None, "") else "")

    safe_set(ws, "A12", txt(purchasing))
    safe_set(ws, "C12", txt(cat_name))
    safe_set(ws, "E12", oxygen)

    # certificates list: K/L rows 4-12
    today = date.today()
    for r in range(4, 13):
        ws[f"K{r}"].value = None
        ws[f"L{r}"].value = None

    certs_sorted = sorted(certs_by_pack or [], key=lambda x: (x.get("pack_name") or "").lower())
    for idx, c in enumerate(certs_sorted[:9]):
        row = 4 + idx
        pack_name = (c.get("pack_name") or "").strip() or "(Unknown pack)"
        safe_set(ws, f"K{row}", pack_name)

        end = parse_date_any(c.get("certificate_end_date"))
        if end and end >= today:
            safe_set(ws, f"L{row}", end, number_format="dd-mm-yyyy")
        else:
            safe_set(ws, f"L{row}", "Expired")


def fill_inventory_sheet(ws, rows: List[Dict[str, Any]], storages_by_id: Dict[str, str]):
    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if isinstance(v, str) and v.strip():
            headers[v.strip().lower()] = c

    def col_required(name: str) -> int:
        c = headers.get(name.lower())
        if not c:
            raise HTTPException(status_code=500, detail=f"Inventory template missing column: {name}")
        return c

    def col_optional(name: str) -> Optional[int]:
        return headers.get(name.lower())

    c_storage = col_required("Storage")
    c_article = col_required("Article No.")
    c_item = col_required("Item name")
    c_qty = col_required("Quantity")
    c_total = col_required("Total quantity")
    c_cert = col_required("Certificate qty")
    c_exp = col_required("Expiry date")
    c_law = col_required("Law code")
    c_pack = col_required("Pack name")

    c_class = {k: col_optional(k) for k in CLASS_LETTERS}

    # clear old rows
    for r in range(2, MAX_ROWS + 2):
        for cidx in (c_storage, c_article, c_item, c_qty, c_total, c_cert, c_exp, c_law, c_pack):
            ws.cell(r, cidx).value = None
        for _, cidx in c_class.items():
            if cidx:
                ws.cell(r, cidx).value = None

    out_i = 0
    for rr in rows or []:
        qty = num(rr.get("vessel_item_quantity")) or 0
        if qty <= 0:
            continue

        r = 2 + out_i
        if r > MAX_ROWS + 1:
            break
        out_i += 1

        storage_display = (rr.get("storage_display") or "").strip()
        storage_id = str(rr.get("storage_id") or "").strip()
        if not storage_display and storage_id and storage_id in storages_by_id:
            storage_display = storages_by_id[storage_id]
        if not storage_display:
            storage_display = "Inbound storage"

        barcode = (rr.get("item_barcode") or "").strip() or "Extra"

        ws.cell(r, c_storage).value = storage_display
        ws.cell(r, c_article).value = barcode
        ws.cell(r, c_item).value = rr.get("vessel_item_name") or ""
        ws.cell(r, c_qty).value = qty
        ws.cell(r, c_total).value = num(rr.get("totalitem_qty_sql")) or 0
        ws.cell(r, c_cert).value = num(rr.get("certificate_qty_sql")) or 0

        d = parse_date_any(rr.get("vessel_item_expiration_date"))
        if d:
            ws.cell(r, c_exp).value = d
            ws.cell(r, c_exp).number_format = "mm-yyyy"
        else:
            ws.cell(r, c_exp).value = None

        ws.cell(r, c_law).value = rr.get("vessel_item_law_code") or ""
        ws.cell(r, c_pack).value = rr.get("pack_name") or ""

        flags: Set[str] = set()
        flags |= letters_from_item_type(rr.get("item_type"))

        for letter in CLASS_LETTERS:
            cidx = c_class.get(letter)
            if not cidx:
                continue
            ws.cell(r, cidx).value = letter if letter in flags else None


def fill_upload_sheet(ws, rows: List[Dict[str, Any]], storages_by_id: Dict[str, str]):
    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if isinstance(v, str) and v.strip():
            headers[v.strip().lower()] = c

    def setv(r: int, h: str, v):
        c = headers.get(h.lower())
        if c:
            ws.cell(r, c).value = v

    for r in range(2, MAX_ROWS + 2):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None

    out_i = 0
    for rr in rows or []:
        qty = num(rr.get("vessel_item_quantity")) or 0
        if qty <= 0:
            continue

        r = 2 + out_i
        if r > MAX_ROWS + 1:
            break
        out_i += 1

        storage_display = (rr.get("storage_display") or "").strip()
        storage_id = str(rr.get("storage_id") or "").strip()
        if not storage_display and storage_id and storage_id in storages_by_id:
            storage_display = storages_by_id[storage_id]
        if not storage_display:
            storage_display = "Inbound storage"

        barcode = (rr.get("item_barcode") or "").strip() or "123"
        d = parse_date_any(rr.get("vessel_item_expiration_date"))

        setv(r, "vessel_item_id", rr.get("vessel_item_id") or "")
        setv(r, "vessel_id", rr.get("vessel_id") or "")
        setv(r, "storage_id", rr.get("storage_id") or "")
        setv(r, "storage_display", storage_display)
        setv(r, "item_id", rr.get("item_id") or "")
        setv(r, "pack_id", rr.get("pack_id") or "")
        setv(r, "category_id", rr.get("category_id") or "")
        setv(r, "item_barcode", barcode)
        setv(r, "vessel_item_name", rr.get("vessel_item_name") or "")
        setv(r, "vessel_item_law_code", rr.get("vessel_item_law_code") or "")
        setv(r, "vessel_item_quantity", qty)
        setv(r, "totalitem_qty_sql", num(rr.get("totalitem_qty_sql")) or 0)
        setv(r, "certificate_qty_sql", num(rr.get("certificate_qty_sql")) or 0)
        setv(r, "pack_name", rr.get("pack_name") or "")
        setv(r, "item_type", rr.get("item_type") or "")

        if d:
            setv(r, "vessel_item_expiration_date", d)
            c = headers.get("vessel_item_expiration_date")
            if c:
                ws.cell(r, c).number_format = "dd-mm-yyyy"


def build_workbook_bytes(
    rows: List[Dict[str, Any]],
    vessel: Dict[str, Any],
    vessel_id: str,
    certs_by_pack: List[Dict[str, Any]],
    storages: List[Dict[str, Any]],
) -> bytes:
    wb = load_workbook(TEMPLATE_PATH)

    storages_by_id = {
        str(s.get("storage_id")): (s.get("storage_name") or "")
        for s in (storages or [])
        if s.get("storage_id")
    }

    fill_storage_sheet(wb, storages)

    ws_inv = find_sheet_fuzzy(wb, INVENTORY_SHEET)
    if ws_inv is None:
        raise HTTPException(status_code=500, detail=f"Sheet '{INVENTORY_SHEET}' not found in template")
    fill_inventory_sheet(ws_inv, rows, storages_by_id)

    ws_up = find_sheet_fuzzy(wb, UPLOAD_SHEET)
    if ws_up is not None:
        fill_upload_sheet(ws_up, rows, storages_by_id)

    next_resupply = compute_next_resupply_date(certs_by_pack)
    fill_vessel_information_sheet(wb, vessel, vessel_id, certs_by_pack, next_resupply, rows)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ==============================
# Routes
# ==============================

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

    rows = get_export_rows(excel_id)
    rows = [r for r in (rows or []) if (num(r.get("vessel_item_quantity")) or 0) > 0]

    certs = get_latest_certificates_by_pack(vessel_id)
    storages = get_vessel_storages(vessel_id)

    content = build_workbook_bytes(rows, vessel, vessel_id, certs, storages)
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
    rows = [r for r in (rows or []) if (num(r.get("vessel_item_quantity")) or 0) > 0]

    certs = get_latest_certificates_by_pack(vessel_id)
    storages = get_vessel_storages(vessel_id)

    content = build_workbook_bytes(rows, vessel, vessel_id, certs, storages)
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
