# Sales & Rate Manager — corrected register parser

A Flask web app based on the original `Sale (2).xlsx` workbook and the two fresh rate-register samples supplied by the user.

## Features
- Upload Sales Register for each reconciliation.
- Upload SM Sale and Danish Purchase Register Excel files to replace their respective rate masters.
- Product mapping editor: add/edit/delete mappings directly in the website; mappings persist in the database.
- Exact stock-name matches are automatic. Uncertain matches are suggestions only; they remain `NA` until a mapping is saved.
- Rate selection priority: Danish Purchase first, then SM Sale.
- Rate = Amount / Quantity. Printed rate columns are not used.
- For each product, latest transaction date wins; if dates tie, last valid occurrence in workbook row order wins.
- SM Sale rows with blank/invalid amount are ignored. They do not replace a valid earlier rate in the same upload.
- Danish Purchase and SM Sale rate masters are replaced only after at least one valid product-rate row has been extracted.
- Sales reports are saved and downloadable as Excel.
- Missing rates remain blank.

## Exact rate-register layouts implemented

### Danish Purchase Register (`Purchase Register` sheet)
The workbook has title rows above the column headings. The parser reads the sheet as raw cells:
- Column A: transaction date on voucher header; carried down to following detail rows.
- Column B: particulars/product name.
- Column C: quantity.
- Column D: printed rate — ignored.
- Column E: amount — used for calculation.
- Supplier/voucher header rows, `Local Purchase`, `SMART R/OFF`, `Rounded Off`, headings and blank rows are ignored because they do not contain a valid product + quantity + amount combination.

### SM Sale (`SMART MARKETING-SALE` sheet)
The workbook has title rows above the column headings. The parser reads the sheet as raw cells:
- Column A: transaction date on voucher header; carried down to following detail rows.
- Column C: product name.
- Column D: quantity.
- Column E: printed rate — ignored.
- Column F: amount — used for calculation.
- Opening/closing balances, voucher headers, `Local Sales`, `SMART R/OFF`, `(as per details)`, headings and blank rows are ignored because they do not contain a valid product + quantity + amount combination.
- Product rows with blank/invalid amount are ignored. The latest valid dated rate remains eligible.

## Final sales report columns
1. Particulars
2. Narration
3. Sale Name
4. Pairs
5. Final Name
6. SM rate or danish pur rate

`Final Name` is used for rate lookup. Danish Purchase rate is selected first; SM Sale is used only if Danish Purchase has no rate for that product.

## Run locally on Windows
1. Install Python 3.11 or 3.12.
2. Extract this ZIP and open Command Prompt in the folder.
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

Local mode uses SQLite. If no environment credentials are configured, the initial login is `admin` / `ChangeMe123!`. Change these immediately.

## Deploy to Render
1. Push the extracted folder to a private GitHub repository.
2. Create a Render Web Service using the repository.
3. Build command: `pip install -r requirements.txt`
4. Start command: `gunicorn app:app --workers 1 --threads 4 --timeout 120`
5. Set `SECRET_KEY`, `ADMIN_USERNAME`, `ADMIN_PASSWORD`, and `DATABASE_URL` to a persistent PostgreSQL connection string.
6. Deploy.

Do not rely on Render's local filesystem for persistent data. Free database availability, limits and retention can change; keep regular database backups.

## Notes
- The uploaded rate workbooks were checked against the extraction logic: Danish Purchase produced 9 valid detail rows (8 unique products after latest-date/tie selection); SM Sale produced 190 valid amount-bearing detail rows (112 unique products after latest-date/tie selection). These are validation counts for the uploaded sample files, not fixed limits.
- The rate parser expects the corresponding register to be on the first worksheet.
- Test with copies of your registers before using outputs for accounting decisions.
