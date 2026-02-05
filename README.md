Cloud Run vessel Excel export (template-based, 3 sheets)

Keeps the exact look by using your uploaded XLSX as the base template and only writing cell values.

Sheets:
- Inventory: fills rows starting at template row 3; replicates row 3 styling for all rows.
- Vessel Information: fills specific cells and clears AppSheet Start/End tags.
- Upload: hidden staging sheet with IDs for matching.

Endpoints:
- GET /download?excel_id=...&token=...
- GET /email?excel_id=...&token=...

Required env:
- DATABASE_URL

Optional env:
- INCLUDE_MATCH_KEY=true  (adds 'match_key' column on Upload sheet)
SMTP env for /email:
- SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, SMTP_USE_TLS, FROM_EMAIL, FROM_NAME

DB requirements:
- vessel_excel(excel_id, vessel_id, export_token)
- vw_vessel_excel_items must contain columns selected in get_export_rows()
- vessels (or vw_vessels_enriched) for Vessel Information

Security model:
- token in query string must match vessel_excel.export_token for excel_id.
