# Sales & Rate Manager — first build

A Flask web app based on the supplied `Sale (2).xlsx` structure.

## Included
- Sales Register upload and report generation.
- SM Sale Ledger and Danish Purchase Register rate uploads. Each successful upload replaces that rate master.
- Rate = transaction amount (`Vch Type`) / quantity (number at start of `Column1`). Voucher dates are filled down to detail rows.
- Per product, latest transaction date wins; if that date occurs more than once, the last row in file order wins.
- Product mapping editor with direct in-table edits and a Save button. Mappings are stored in the database.
- Exact stock-name matching, high-confidence unique auto-mapping, and weaker fuzzy suggestions for manual review.
- Missing rate values remain blank.
- Sales reports saved in database and downloadable as Excel.
- Seed stock master extracted from the original workbook's `Stock Name` sheet.

## Run locally on Windows
1. Install Python 3.11+.
2. Extract this ZIP, open Command Prompt in the folder.
3. Run:

```bat
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
set ADMIN_USERNAME=admin
set ADMIN_PASSWORD=SetYourOwnStrongPassword
set SECRET_KEY=SetALongRandomSecret
python app.py
```

4. Open `http://127.0.0.1:5000`.

Local mode uses SQLite. Initial default credentials if environment variables are not set: `admin` / `ChangeMe123!`. Change these immediately.

## Deploy to Render
1. Push the extracted folder to a private GitHub repository.
2. Create a Render Web Service using this repository.
3. Build command: `pip install -r requirements.txt`
4. Start command: `gunicorn app:app --workers 1 --threads 4 --timeout 120`
5. Set environment variables: `SECRET_KEY` (long random value), `ADMIN_USERNAME`, `ADMIN_PASSWORD`, and `DATABASE_URL` (PostgreSQL connection string from a persistent database provider).
6. Deploy.

Do not use the Render local filesystem for persistent production data. Free database plans have provider-specific storage, inactivity and retention limits; keep regular backups.

## Important first-build notes
- This parser targets the column layouts found in the supplied workbook. It expects the first worksheet of each uploaded workbook to contain the corresponding table data and headers.
- SM Sale needs `Date`, `Particulars1`, `Column1`, `Vch Type`; Danish Purchase needs `Date`, `Particulars`, `Column1`, `Vch Type`; Sales Register needs `Particulars`, `Column1`, `Vch Type`, `Vch No.`.
- The app currently returns both SM and Danish rates as separate columns. Confirm the final report layout and any additional workbook formulas against a real test upload before using it for accounting decisions.
- High-confidence auto-mapping is intentionally conservative. Similar but uncertain matches are displayed as suggestions and are not silently confirmed.
- This is an initial working build, not a security-audited production system. Use a strong secret/password and a private repository.
