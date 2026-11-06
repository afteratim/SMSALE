import os, re, io, csv, json, secrets
from datetime import datetime, date
from difflib import SequenceMatcher
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file, jsonify
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from sqlalchemy import create_engine, text
from sqlalchemy.exc import SQLAlchemyError
from openpyxl import load_workbook, Workbook

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "change-this-secret-before-deploy-" + secrets.token_hex(8))
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024
DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///sales_rate_app.db")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)
engine = create_engine(DATABASE_URL, pool_pre_ping=True, future=True)

def db(sql, params=None, fetch=False, many=False):
    with engine.begin() as c:
        result = c.execute(text(sql), params or {})
        if fetch:
            return [dict(x._mapping) for x in result.fetchall()]
        if many:
            return result.rowcount
        return result

def init_db():
    idtype = "SERIAL PRIMARY KEY" if engine.dialect.name == "postgresql" else "INTEGER PRIMARY KEY AUTOINCREMENT"
    stmts = [
    f"""CREATE TABLE IF NOT EXISTS users (id {idtype}, username VARCHAR(100) UNIQUE NOT NULL, password_hash TEXT NOT NULL)""",
    f"""CREATE TABLE IF NOT EXISTS stock_master (id {idtype}, name TEXT UNIQUE NOT NULL)""",
    f"""CREATE TABLE IF NOT EXISTS mappings (id {idtype}, sales_name TEXT UNIQUE NOT NULL, mapped_name TEXT NOT NULL, confirmed INTEGER DEFAULT 1, updated_at TEXT NOT NULL)""",
    f"""CREATE TABLE IF NOT EXISTS sm_rates (id {idtype}, product TEXT NOT NULL, rate FLOAT NOT NULL, txn_date TEXT, row_order INTEGER NOT NULL)""",
    f"""CREATE TABLE IF NOT EXISTS danish_rates (id {idtype}, product TEXT NOT NULL, rate FLOAT NOT NULL, txn_date TEXT, row_order INTEGER NOT NULL)""",
    f"""CREATE TABLE IF NOT EXISTS upload_history (id {idtype}, kind TEXT NOT NULL, filename TEXT, uploaded_at TEXT NOT NULL, row_count INTEGER DEFAULT 0, note TEXT)""",
    f"""CREATE TABLE IF NOT EXISTS sales_reports (id {idtype}, filename TEXT, created_at TEXT NOT NULL, report_json TEXT NOT NULL)"""
    ]
    with engine.begin() as c:
        for s in stmts: c.execute(text(s))
        count = c.execute(text("SELECT COUNT(*) FROM stock_master")).scalar()
        if not count:
            try:
                with open(os.path.join(os.path.dirname(__file__), "stock_master_seed.csv"), encoding="utf-8-sig") as f:
                    for row in csv.DictReader(f):
                        if row.get("name"):
                            c.execute(text("INSERT INTO stock_master(name) VALUES(:n)"), {"n":row["name"]})
            except FileNotFoundError: pass
        users = c.execute(text("SELECT COUNT(*) FROM users")).scalar()
        if not users:
            username=os.getenv("ADMIN_USERNAME","admin")
            password=os.getenv("ADMIN_PASSWORD","ChangeMe123!")
            c.execute(text("INSERT INTO users(username,password_hash) VALUES(:u,:p)"),
                      {"u":username,"p":generate_password_hash(password)})

def login_required(fn):
    @wraps(fn)
    def wrap(*a, **kw):
        if not session.get("user"): return redirect(url_for("login"))
        return fn(*a, **kw)
    return wrap

def norm(s):
    return re.sub(r"[^A-Z0-9]+","",str(s or "").upper())

def is_empty(value):
    return value is None or (isinstance(value, str) and not value.strip())

def parse_qty(value):
    if is_empty(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    match = re.search(r"-?\d+(?:\.\d+)?", str(value).replace(",", ""))
    return float(match.group()) if match else None

def parse_date(value):
    if is_empty(value):
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    value = str(value).strip()
    for fmt in ("%d-%b-%y", "%d-%b-%Y", "%d/%m/%Y", "%d/%m/%y",
                "%Y-%m-%d", "%d-%m-%Y", "%d-%m-%y"):
        try:
            return datetime.strptime(value, fmt).date().isoformat()
        except ValueError:
            pass
    return None

def open_workbook_rows(file):
    if not file or not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise ValueError("Please upload an .xlsx or .xlsm Excel file.")
    file.stream.seek(0)
    wb = load_workbook(file.stream, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    return rows

def numeric_value(value):
    if is_empty(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace(",", "").replace("Dr", "").replace("Cr", "").strip()
    if not s:
        return None
    try:
        return float(s)
    except (TypeError, ValueError):
        return None

def extract_danish_rates(raw):
    """
    Danish Purchase: A=date, B=product, C=qty, D=printed rate (ignored),
    E=amount. Voucher date is filled down to subsequent product rows.
    """
    out=[]
    current_date=None
    for i, row in enumerate(raw):
        row = tuple(row)
        parsed_date = parse_date(row[0] if len(row)>0 else None)
        if parsed_date:
            current_date = parsed_date
        product = row[1] if len(row)>1 else None
        qty = numeric_value(row[2] if len(row)>2 else None)
        amount = numeric_value(row[4] if len(row)>4 else None)
        if is_empty(product) or qty in (None, 0) or amount is None or current_date is None:
            continue
        product = str(product).strip()
        if product.upper() in {"LOCAL PURCHASE", "SMART R/OFF", "ROUNDED OFF", "ROUNDED OFF (P)"}:
            continue
        out.append({"product":product, "rate":amount/qty, "txn_date":current_date, "row_order":i})
    return select_latest_rates(out), len(out)

def extract_sm_rates(raw):
    """
    SM Sale: A=date, C=product, D=qty, E=printed rate (ignored),
    F=amount. Blank/invalid amount rows are skipped.
    """
    out=[]
    current_date=None
    for i, row in enumerate(raw):
        row = tuple(row)
        parsed_date = parse_date(row[0] if len(row)>0 else None)
        if parsed_date:
            current_date = parsed_date
        product = row[2] if len(row)>2 else None
        qty = numeric_value(row[3] if len(row)>3 else None)
        amount = numeric_value(row[5] if len(row)>5 else None)
        if is_empty(product) or qty in (None, 0) or amount is None or current_date is None:
            continue
        product = str(product).strip()
        if product.upper() in {"LOCAL SALES", "SMART R/OFF", "ROUNDED OFF", "(AS PER DETAILS)", "OPENING BALANCE", "CLOSING BALANCE"}:
            continue
        out.append({"product":product, "rate":amount/qty, "txn_date":current_date, "row_order":i})
    return select_latest_rates(out), len(out)

def select_latest_rates(rows):
    latest={}
    for item in rows:
        key=norm(item["product"])
        if not key:
            continue
        old=latest.get(key)
        if old is None or (item["txn_date"], item["row_order"]) >= (old["txn_date"], old["row_order"]):
            latest[key]=item
    return list(latest.values())

def sales_extract(raw):
    if not raw:
        raise ValueError("Sales Register is empty.")
    headers = [str(x).strip() if x is not None else "" for x in raw[0]]
    required = ["Particulars", "Column1", "Vch Type", "Vch No."]
    missing = [name for name in required if name not in headers]
    if missing:
        raise ValueError("Sales Register: missing columns: " + ", ".join(missing))
    positions = {name: headers.index(name) for name in required}
    current=None
    result=[]
    excluded={"LOCAL SALE", "CGST", "SGST", "ROUNDED OFF"}
    for row in raw[1:]:
        row=tuple(row)
        def cell(name):
            pos=positions[name]
            return row[pos] if pos < len(row) else None
        vno=cell("Vch No.")
        if not is_empty(vno):
            party=cell("Particulars")
            v=str(vno)
            current=f"{v[3:7]} - {party}" if not is_empty(party) else v[3:7]
        vtype=cell("Vch Type")
        if str(vtype).strip().lower()=="sales":
            continue
        item=cell("Particulars")
        if is_empty(item) or str(item).strip().upper() in excluded:
            continue
        qty=parse_qty(cell("Column1"))
        if qty is None:
            continue
        result.append({"Particulars":str(item).strip(), "Custom 1":current or "",
                       "Pairs":int(qty) if float(qty).is_integer() else qty})
    return result

def log_upload(kind, filename, count, note=""):
    db("INSERT INTO upload_history(kind,filename,uploaded_at,row_count,note) VALUES(:k,:f,:d,:r,:n)",
       {"k":kind,"f":filename,"d":datetime.utcnow().isoformat(timespec="seconds"),"r":count,"n":note})

def replace_rates(table, rows):
    # Validation happens before this transaction; old master remains if parsing fails.
    with engine.begin() as c:
        c.execute(text(f"DELETE FROM {table}"))
        for r in rows:
            c.execute(text(f"INSERT INTO {table}(product,rate,txn_date,row_order) VALUES(:p,:r,:d,:o)"),
                      {"p":r["product"],"r":r["rate"],"d":r["txn_date"],"o":r["row_order"]})

def mappings_dict():
    return {norm(x["sales_name"]):x["mapped_name"] for x in db("SELECT sales_name,mapped_name FROM mappings",fetch=True)}

def rate_dict(table):
    return {norm(x["product"]):x["rate"] for x in db(f"SELECT product,rate FROM {table}",fetch=True)}

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method=="POST":
        rows=db("SELECT * FROM users WHERE username=:u",{"u":request.form.get("username","")},True)
        if rows and check_password_hash(rows[0]["password_hash"],request.form.get("password","")):
            session["user"]=rows[0]["username"]; return redirect(url_for("index"))
        flash("Invalid username or password.","error")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear(); return redirect(url_for("login"))

@app.route("/")
@login_required
def index():
    counts={t:db(f"SELECT COUNT(*) n FROM {t}",fetch=True)[0]["n"] for t in ["mappings","sm_rates","danish_rates","sales_reports"]}
    history=db("SELECT * FROM upload_history ORDER BY id DESC LIMIT 8",fetch=True)
    return render_template("index.html",counts=counts,history=history)

@app.route("/rates",methods=["GET","POST"])
@login_required
def rates():
    if request.method=="POST":
        kind=request.form.get("kind")
        try:
            f=request.files.get("file")
            raw=open_workbook_rows(f)
            if kind=="sm":
                rows,total=extract_sm_rates(raw)
                table="sm_rates"; label="SM Sale"
            elif kind=="danish":
                rows,total=extract_danish_rates(raw)
                table="danish_rates"; label="Danish Purchase"
            else: raise ValueError("Unknown rate type.")
            if not rows:
                raise ValueError("No valid dated product rows with quantity and amount were found. Existing rates were not changed.")
            replace_rates(table,rows)
            log_upload(label,f.filename,len(rows),f"{total} valid product rows; latest date and last same-date occurrence retained.")
            flash(f"{label} rates replaced successfully: {len(rows)} products saved.","success")
        except Exception as e: flash(str(e),"error")
        return redirect(url_for("rates"))
    sm=db("SELECT product,rate,txn_date FROM sm_rates ORDER BY product",fetch=True)
    dan=db("SELECT product,rate,txn_date FROM danish_rates ORDER BY product",fetch=True)
    return render_template("rates.html",sm=sm,dan=dan)

@app.route("/mappings",methods=["GET","POST"])
@login_required
def mappings():
    if request.method=="POST":
        action=request.form.get("action")
        try:
            if action=="save":
                sales=(request.form.get("sales_name") or "").strip()
                mapped=(request.form.get("mapped_name") or "").strip()
                if not sales or not mapped: raise ValueError("Both product names are required.")
                db("""INSERT INTO mappings(sales_name,mapped_name,confirmed,updated_at)
                    VALUES(:s,:m,1,:d) ON CONFLICT(sales_name) DO UPDATE SET mapped_name=:m,confirmed=1,updated_at=:d""",
                   {"s":sales,"m":mapped,"d":datetime.utcnow().isoformat(timespec="seconds")})
                flash("Mapping saved.","success")
            elif action=="delete":
                db("DELETE FROM mappings WHERE id=:id",{"id":int(request.form["id"])})
                flash("Mapping deleted.","success")
            elif action=="add_stock":
                name=request.form.get("name","").strip()
                if name: db("INSERT INTO stock_master(name) VALUES(:n) ON CONFLICT(name) DO NOTHING",{"n":name})
            elif action=="bulk_auto":
                candidates=db("SELECT name FROM stock_master",fetch=True)
                existing=mappings_dict(); sales_names=db("SELECT DISTINCT product FROM sales_products_view",fetch=True) if False else []
                flash("Auto suggestions are generated when a Sales Register is processed.","success")
        except Exception as e: flash(str(e),"error")
        return redirect(url_for("mappings"))
    q=request.args.get("q","").strip()
    rows=db("SELECT * FROM mappings WHERE sales_name LIKE :q OR mapped_name LIKE :q ORDER BY sales_name LIMIT 1000",{"q":f"%{q}%"},True)
    stocks=db("SELECT name FROM stock_master ORDER BY name LIMIT 2000",fetch=True)
    return render_template("mappings.html",rows=rows,stocks=[x["name"] for x in stocks],q=q)

@app.route("/sales",methods=["GET","POST"])
@login_required
def sales():
    if request.method=="POST":
        try:
            f=request.files.get("file"); raw=open_workbook_rows(f)
            items=sales_extract(raw)
            if not items: raise ValueError("No sales product rows found.")
            maps=mappings_dict()
            stock=db("SELECT name FROM stock_master",fetch=True)
            stock_by_norm={norm(x["name"]):x["name"] for x in stock}
            sm=rate_dict("sm_rates"); dan=rate_dict("danish_rates")
            out=[]; suggestions=[]
            for x in items:
                source=x["Particulars"]
                key=norm(source)
                saved_mapping=maps.get(key)
                exact_sale_name=stock_by_norm.get(key)
                sale_name=exact_sale_name or "NA"

                # Exact stock matches are automatic. Non-exact matches are only
                # suggestions; they remain NA until the user saves a mapping.
                if not saved_mapping and not exact_sale_name:
                    ranked=sorted(
                        ((SequenceMatcher(None,key,norm(s)).ratio(),s) for s in stock_by_norm.values()),
                        reverse=True
                    )
                    if ranked and ranked[0][0] >= 0.72:
                        suggestions.append({
                            "sales": source,
                            "suggested": ranked[0][1],
                            "score": round(ranked[0][0]*100)
                        })

                final_name=saved_mapping or exact_sale_name or "NA"
                final_key=norm(final_name) if final_name != "NA" else ""
                # Danish Purchase Rate has priority. SM Sale is the fallback.
                selected_rate=dan.get(final_key)
                selected_source="Danish Purchase" if selected_rate is not None else None
                if selected_rate is None:
                    selected_rate=sm.get(final_key)
                    if selected_rate is not None:
                        selected_source="SM Sale"

                out.append({
                    "Particulars": source,
                    "Narration": x.get("Custom 1",""),
                    "Sale Name": sale_name,
                    "Pairs": x.get("Pairs",""),
                    "Final Name": final_name,
                    "SM rate or danish pur rate": selected_rate,
                    "Rate Source": selected_source or "Missing Rate"
                })
            report={"rows":out,"suggestions":suggestions}
            cur=db("INSERT INTO sales_reports(filename,created_at,report_json) VALUES(:f,:d,:j) RETURNING id",
                   {"f":f.filename,"d":datetime.utcnow().isoformat(timespec="seconds"),"j":json.dumps(report,ensure_ascii=False)})
            report_id=cur.scalar()
            session["last_report"]=report_id
            return redirect(url_for("report",report_id=report_id))
        except Exception as e: flash(str(e),"error")
        return redirect(url_for("sales"))
    return render_template("sales.html")

@app.route("/report/<int:report_id>")
@login_required
def report(report_id):
    rows=db("SELECT * FROM sales_reports WHERE id=:id",{"id":report_id},True)
    if not rows: return "Report not found",404
    data=json.loads(rows[0]["report_json"])
    return render_template("report.html",report=rows[0],rows=data["rows"],suggestions=data["suggestions"])

@app.route("/report/<int:report_id>/download")
@login_required
def download_report(report_id):
    rows=db("SELECT * FROM sales_reports WHERE id=:id",{"id":report_id},True)
    if not rows: return "Report not found",404
    data=json.loads(rows[0]["report_json"])
    export_columns=["Particulars","Narration","Sale Name","Pairs","Final Name","SM rate or danish pur rate"]
    buf=io.BytesIO()
    wb=Workbook()
    ws=wb.active
    ws.title="Sales Result"
    ws.append(export_columns)
    for item in data["rows"]:
        ws.append([item.get(column) for column in export_columns])
    suggest_ws=wb.create_sheet("Mapping Suggestions")
    suggest_ws.append(["Sales product","Suggested stock product","Similarity"])
    for item in data["suggestions"]:
        suggest_ws.append([item.get("sales"),item.get("suggested"),item.get("score")])
    wb.save(buf)
    buf.seek(0)
    return send_file(buf,as_attachment=True,download_name="Sales_Reconciliation.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/history")
@login_required
def history():
    rows=db("SELECT id,filename,created_at FROM sales_reports ORDER BY id DESC",fetch=True)
    return render_template("history.html",rows=rows)

@app.route("/api/mappings/<int:mid>",methods=["PUT","DELETE"])
@login_required
def mapping_api(mid):
    if request.method=="DELETE":
        db("DELETE FROM mappings WHERE id=:id",{"id":mid}); return jsonify(ok=True)
    data=request.get_json(force=True)
    s=(data.get("sales_name") or "").strip(); m=(data.get("mapped_name") or "").strip()
    if not s or not m: return jsonify(error="Both names required"),400
    db("UPDATE mappings SET sales_name=:s,mapped_name=:m,updated_at=:d WHERE id=:id",
       {"s":s,"m":m,"d":datetime.utcnow().isoformat(timespec="seconds"),"id":mid})
    return jsonify(ok=True)

@app.errorhandler(413)
def too_large(e): return "File exceeds 30 MB upload limit.",413

with app.app_context():
    init_db()

if __name__=="__main__":
    app.run(host="0.0.0.0",port=int(os.getenv("PORT",5000)),debug=False)
