# FleetMed Vessel Excel Export (FastAPI + Cloud Run)

Purpose
Generate a “Medical Inventory List” Excel for a single vessel using:
- excel_id + token (security)
- selectors stored in vessel_excel (storage_id and item_status)
- inventory rows from vw_vessel_excel_items
- vessel header data from vw_vessels_enriched

Output workbook contains:
- Vessel Information
- Inventory
- Upload
- Storage

## Endpoints
GET /health
Returns { "ok": true }.

GET /download?excel_id=...&token=...
Validates token and redirects to /download_file. Retries briefly if AppSheet/DB is still saving.

GET /download_file?excel_id=...&token=...
Streams the generated .xlsx file.

GET /email?excel_id=...&token=...&to_email=...
Sends the generated .xlsx by email (requires SMTP env vars).

Optional:
GET /debug_excel?excel_id=...&token=...
Returns the raw and parsed selector values read from vessel_excel (useful for debugging filter issues).

## Required environment variables
DATABASE_URL
PostgreSQL connection string.

TEMPLATE_PATH
Path to the Excel template file.
Example: app/templates/inventory_template 2.06.xlsx (or your latest)

SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, SMTP_USE_TLS
Only required for /email.

FROM_EMAIL, FROM_NAME
Optional sender settings for /email.

MAX_ROWS
Optional, default 3000.

## Database dependencies

### Table: vessel_excel
Must contain (at least):
- excel_id (key)
- vessel_id
- export_token
- storage_id (EnumList from AppSheet, comma/semicolon separated list of storage IDs)
- item_status (Enum, e.g. Active or Scrap)

The service reads selectors from this row to filter export output.

### View: vw_vessels_enriched
Must contain vessel header fields and enriched names:
- vessel_flag_name (resolved from public.entity using vessel_flag)
- vessel_category_name (resolved from public.vessel_cerficate_categories using vessel_category)

### View: vw_vessel_excel_items
Must contain:
- excel_id
- vessel_item_quantity
- storage_id
- storage_display
- item_status
- item_type (EnumList string, includes “Cool Good”, “Narcotic”, etc.)
- fields used for export columns (barcode, name, law code, expiry date, pack name, totals, etc.)

## Filtering rules
Rows are always excluded if vessel_item_quantity <= 0.

If vessel_excel.item_status is set:
- only rows with item_status matching the selected value are exported.

If vessel_excel.storage_id is set (one or more IDs):
- only rows with storage_id in the selected list are exported.

Status + Storage filters are combined with AND.

## Classification columns (N/M/C/D/F/O)
For each Inventory row, letters are derived from item_type (EnumList):
- Cool Good -> C
- Narcotic -> N
- Dangerous good -> D
- Female Gender -> F
- Medical Oxygen -> O
- Malaria -> M

## Template requirements
Inventory sheet must contain headers (exact spelling):
Storage, Article No., Item name, Quantity, Total quantity, Certificate qty, Expiry date, Law code,
N, M, C, D, F, O,
Status,
Pack name

Vessel Information sheet cells are filled according to the template mapping used in main.py.

## Deployment
Cloud Run runs the service. Ensure:
- The Cloud Run revision contains the latest main.py
- DATABASE_URL points to the correct Cloud SQL database
- TEMPLATE_PATH points to the correct template file packaged in the container

## Debug checklist
1) Verify vessel_excel row contains selectors:
   SELECT excel_id, storage_id, item_status FROM vessel_excel WHERE excel_id='...';

2) Verify expected rows in vw_vessel_excel_items:
   SELECT COUNT(*) FROM vw_vessel_excel_items WHERE excel_id='...';

3) Call /debug_excel to confirm the service reads selectors correctly.
