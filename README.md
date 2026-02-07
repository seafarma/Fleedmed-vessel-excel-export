# Vessel Excel Export (FastAPI)

Service that generates an Excel “Medical Inventory List” for a vessel using a template workbook and data from PostgreSQL.

## Endpoints

- `GET /health` → `{ "ok": true }`
- `GET /download?excel_id=...&token=...`
  - If the row is not yet visible in DB, returns a self-refreshing HTML “Preparing…” page.
  - Otherwise redirects (302) to `/download_file`.
- `GET /download_file?excel_id=...&token=...` → downloads the Excel file
- `GET /email?excel_id=...&token=...&to_email=...` → emails the Excel file as attachment

## Required environment variables

- `DATABASE_URL` (required)  
  Example: `postgresql://user:pass@host:5432/dbname`

## Optional environment variables

- `TEMPLATE_PATH` (default: `app/templates/inventory_template.xlsx`)
- `INVENTORY_SHEET` (default: `Inventory`)
- `VESSEL_INFO_SHEET` (default: `Vessel Information`)
- `UPLOAD_SHEET` (default: `Upload`)
- `STORAGE_SHEET` (default: `Storage`)
- `MAX_ROWS` (default: `3000`)

### SMTP (only needed for `/email`)
- `SMTP_HOST`
- `SMTP_PORT` (default: `587`)
- `SMTP_USER`
- `SMTP_PASSWORD`
- `SMTP_USE_TLS` (default: `true`)
- `FROM_EMAIL` (default: `SMTP_USER` or `no-reply@example.com`)
- `FROM_NAME` (default: `Inventory Export`)

## Template expectations

Workbook should include at least:

### Inventory sheet
Headers must contain:
- `Storage`, `Article No.`, `Item name`, `Quantity`, `Total quantity`, `Certificate qty`, `Expiry date`, `Law code`, `Pack name`
And classification columns:
- `N`, `M`, `C`, `D`, `F`, `O`

### Vessel Information sheet (template 2.03 mapping)
- Company: `A4`
- Vessel name: `C4`
- Malaria medicine: `E4`
- MFAG: `G4`
- Next resupply: `I4`
- Certificates (multi-line): `K4`
- Notes: `N4` (merged area)
- Female onboard: `E6`
- Cool Items: `G6`
- Flag: `C8`
- Vessel Category: `C12`
- Medical oxygen: `E12`

### Storage sheet
Columns:
- `A1` = `Storage Name`
- `B1` = `storage_id`
The service fills rows 2..N with vessel storages (sorted).

## Run locally

```bash
pip install -r requirements.txt
export DATABASE_URL="postgresql://..."
uvicorn app.main:app --host 0.0.0.0 --port 8080
