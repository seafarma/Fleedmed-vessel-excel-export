Fleetmed vessel excel export

What this service does
- Generates an Excel file based on app/templates/inventory_template.xlsx
- Uses data from Cloud SQL (PostgreSQL)
- Protects downloads with a per-export token stored in the database table vessel_excel

Endpoints
- GET /health
  Returns {"ok": true}

- GET /download?excel_id=...&token=...
  User-friendly entry endpoint.
  It waits briefly for the excel_id record to appear (race after AppSheet Save).
  If found, it redirects to /download_file.

- GET /download_file?excel_id=...&token=...
  Returns the Excel file as an attachment.

- GET /email?excel_id=...&token=...&to_email=...
  Sends the Excel file to the given email address.
  Requires SMTP env vars to be set.

Authentication / token
- The token is not an environment variable.
- The token must match vessel_excel.export_token for the given excel_id.
- Example query:

  SELECT excel_id, vessel_id, export_token
  FROM vessel_excel
  WHERE excel_id = '976bc0bf';

Download test (Cloud Shell)
- Replace <TOKEN> with the export_token from the database:

  TOKEN="<TOKEN>"
  URL="https://vessel-excel-export-799072059168.europe-west4.run.app/download?excel_id=976bc0bf&token=${TOKEN}"
  curl -L -o ~/test.xlsx "$URL"
  file ~/test.xlsx

If file prints "Microsoft Excel 2007+" then the download is correct.

Environment variables
Required
- DATABASE_URL
  Example (Cloud SQL unix socket):
  postgresql://USER:PASSWORD@/DBNAME?host=/cloudsql/PROJECT:REGION:INSTANCE

Optional
- TEMPLATE_PATH (default: app/templates/inventory_template.xlsx)
- INVENTORY_SHEET (default: Inventory)
- VESSEL_INFO_SHEET (default: Vessel Information)
- UPLOAD_SHEET (default: Upload)

Email (only required for /email)
- SMTP_HOST
- SMTP_PORT (default 587)
- SMTP_USER
- SMTP_PASSWORD
- SMTP_USE_TLS (default true)
- FROM_EMAIL (default SMTP_USER or no-reply@example.com)
- FROM_NAME (default Inventory Export)

Template requirements
The template is loaded and then filled.
Your sheet names must match:
- Inventory
- Vessel Information
- Upload

The Inventory sheet must have the expected headers in row 1 (exact names matter because the code matches by header text).

Docker / Cloud Run
Dockerfile builds a Python 3.11 container and runs:
uvicorn app.main:app --host 0.0.0.0 --port 8080

Local run (optional)
- Set DATABASE_URL and make sure your template is present:
  export DATABASE_URL="..."
  uvicorn app.main:app --host 0.0.0.0 --port 8080
