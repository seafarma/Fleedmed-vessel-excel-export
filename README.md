# Vessel Excel Export (FleetMed)

FastAPI service that generates an Excel file (based on a template) for a vessel inventory export.

## Endpoints

- `GET /health`  
  Returns `{"ok": true}`

- `GET /download?excel_id=...&token=...`  
  Returns a redirect to `/download_file`.  
  If the `excel_id` is not yet available (race after AppSheet save), returns a self-refreshing wait page.

- `GET /download_file?excel_id=...&token=...`  
  Generates and downloads the Excel file.

- `GET /email?excel_id=...&token=...&to_email=...`  
  Generates the Excel file and sends it as an email attachment.

## Environment variables

Required:
- `DATABASE_URL`

Optional:
- `TEMPLATE_PATH` (default: `app/templates/inventory_template.xlsx`)
- `INVENTORY_SHEET` (default: `Inventory`)
- `VESSEL_INFO_SHEET` (default: `Vessel Information`)
- `UPLOAD_SHEET` (default: `Upload`)
- `STORAGE_SHEET` (default: `Storage`)
- `MAX_ROWS` (default: `3000`)

SMTP (only needed if using `/email`):
- `SMTP_HOST`
- `SMTP_PORT` (default `587`)
- `SMTP_USER`
- `SMTP_PASSWORD`
- `SMTP_USE_TLS` (default `true`)
- `FROM_EMAIL`
- `FROM_NAME`

## Run locally

```bash
pip install -r requirements.txt
export DATABASE_URL="postgresql://..."
uvicorn app.main:app --reload --port 8080
