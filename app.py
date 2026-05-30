"""
Transaction Admin — FastAPI backend v2
PostgreSQL + Brands CRUD + Settlement from template
"""

import copy as copy_module
import io
import json
import math
import os
import re
import time
from datetime import datetime
from io import BytesIO

import numpy as np
import pandas as pd
import psycopg2
import psycopg2.extras
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from openpyxl import load_workbook
from openpyxl.styles import Font

# ── optional Google Sheets ────────────────────────────────────────────────────
try:
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
    HAS_GOOGLE = True
except ImportError:
    HAS_GOOGLE = False

# =============================================================================
# CONFIG
# =============================================================================

SPREADSHEET_ID   = os.getenv("SPREADSHEET_ID", "")
RECON_SHEET_NAME = os.getenv("RECON_SHEET_NAME", "1. Reconciliation")
SCOPES           = ["https://www.googleapis.com/auth/spreadsheets"]
SHEETS_WRITE_CHUNK = 2000
HEADER_ROWS        = 2
DATA_START_ROW     = HEADER_ROWS + 1

EXCLUDED_MERCHANTS = {"sognolab", "cratecracker"}
CARD_PAY_TYPES     = {"visa", "mastercard", "maestro"}
OB_PAY_TYPES       = {"open-banking", "banks/germany"}
SHOW_STATUSES      = {"success", "decline", "processing"}

TEMPLATE_PATH = "template.xlsx"

COLUMN_MAP = {
    "provider payment id": ["acquirer_id / provider_payment_id", "provider_payment_id / acquirer_id", "provider payment id", "provider_payment_id"],
    "currency": ["currency", "currency / currency"],
    "transaction status": ["transaction status", "operation_status", "status"],
    "external id payment id": ["external_id / payment_id", "external_id", "payment_id"],
    "operation type": ["operation type", "operation_type", "type"],
    "amount": ["amount", "amount / amount"],
    "мой выгрузка": ["created_at / operation_created_at", "created_at", "operation_created_at"],
    "merchant name": ["merchant name", "merchant_name"],
    "payment_type_id": ["payment_type_id", "payment_type_id / payment_method_type", "payment_method_type"],
    "id / operation_id": ["id / operation_id", "operation_id", "id"],
    "customer email": ["customer email", "customer_email", "email"],
    "customer name": ["customer name"],
    "match": ["match", "Match"],
    "correct reconcile ?": ["correct reconcile ?", "correct_reconcile"],
    "fraud": ["fraud", "Fraud"],
    "fraud action ?": ["fraud action ?", "fraud_action"],
}

# =============================================================================
# DATABASE
# =============================================================================

def get_db():
    url = os.getenv("DATABASE_URL", "")
    if not url:
        raise RuntimeError("DATABASE_URL не задан")
    return psycopg2.connect(url, sslmode="require")

def init_db():
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS brands (
                id SERIAL PRIMARY KEY,
                name VARCHAR(255) UNIQUE NOT NULL,
                volume_fee_eur NUMERIC(10,4) DEFAULT 0,
                volume_fee_non_eur NUMERIC(10,4) DEFAULT 0,
                success_fee NUMERIC(10,4) DEFAULT 0,
                failed_fee NUMERIC(10,4) DEFAULT 0,
                settlement_fee_sepa NUMERIC(10,4) DEFAULT 0,
                settlement_fee_usdt NUMERIC(10,4) DEFAULT 0,
                chargeback_fee NUMERIC(10,4) DEFAULT 0,
                refund_fee NUMERIC(10,4) DEFAULT 0,
                rolling_reserve NUMERIC(10,4) DEFAULT 0,
                ecb_adjustment NUMERIC(10,4) DEFAULT 0,
                created_at TIMESTAMP DEFAULT NOW(),
                updated_at TIMESTAMP DEFAULT NOW()
            )
        """)
        conn.commit()
        cur.close()
        conn.close()
        print("DB initialized OK")
    except Exception as e:
        print(f"DB init error: {e}")

# =============================================================================
# APP
# =============================================================================

app = FastAPI(title="Transaction Admin API v2")

allowed = os.getenv("ALLOWED_ORIGIN", "*")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.on_event("startup")
async def startup():
    init_db()

@app.get("/")
def root():
    return {"status": "ok", "version": "2.0"}

# =============================================================================
# BRANDS CRUD
# =============================================================================

@app.get("/brands")
def get_brands():
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT * FROM brands ORDER BY name")
    rows = cur.fetchall()
    cur.close(); conn.close()
    return [dict(r) for r in rows]

@app.post("/brands")
def upsert_brand(data: dict):
    name = (data.get("name") or "").strip()
    if not name:
        raise HTTPException(400, "name обязателен")
    fields = ["volume_fee_eur","volume_fee_non_eur","success_fee","failed_fee",
              "settlement_fee_sepa","settlement_fee_usdt","chargeback_fee",
              "refund_fee","rolling_reserve","ecb_adjustment"]
    vals = {f: float(data.get(f, 0) or 0) for f in fields}
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO brands (name, volume_fee_eur, volume_fee_non_eur, success_fee, failed_fee,
            settlement_fee_sepa, settlement_fee_usdt, chargeback_fee, refund_fee,
            rolling_reserve, ecb_adjustment, updated_at)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,NOW())
        ON CONFLICT (name) DO UPDATE SET
            volume_fee_eur=EXCLUDED.volume_fee_eur,
            volume_fee_non_eur=EXCLUDED.volume_fee_non_eur,
            success_fee=EXCLUDED.success_fee,
            failed_fee=EXCLUDED.failed_fee,
            settlement_fee_sepa=EXCLUDED.settlement_fee_sepa,
            settlement_fee_usdt=EXCLUDED.settlement_fee_usdt,
            chargeback_fee=EXCLUDED.chargeback_fee,
            refund_fee=EXCLUDED.refund_fee,
            rolling_reserve=EXCLUDED.rolling_reserve,
            ecb_adjustment=EXCLUDED.ecb_adjustment,
            updated_at=NOW()
    """, (name, vals["volume_fee_eur"], vals["volume_fee_non_eur"],
          vals["success_fee"], vals["failed_fee"], vals["settlement_fee_sepa"],
          vals["settlement_fee_usdt"], vals["chargeback_fee"], vals["refund_fee"],
          vals["rolling_reserve"], vals["ecb_adjustment"]))
    conn.commit(); cur.close(); conn.close()
    return {"ok": True, "name": name}

@app.delete("/brands/{name}")
def delete_brand(name: str):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM brands WHERE name=%s", (name,))
    conn.commit(); cur.close(); conn.close()
    return {"ok": True}

# =============================================================================
# HELPERS
# =============================================================================

def norm_key(value):
    return re.sub(r"\s+", " ", (value or "").strip().lower())

def strip_apostrophe(value):
    s = str(value or "").strip()
    return s[1:].lstrip() if s.startswith("'") else s

def read_csv(data: bytes) -> pd.DataFrame:
    df = pd.read_csv(BytesIO(data), sep=";", dtype=str, keep_default_na=False)
    if df.shape[1] <= 1:
        df = pd.read_csv(BytesIO(data), sep=",", dtype=str, keep_default_na=False)
    return df

def to_float(x):
    s = str(x or "").strip().replace(" ", "").replace(",", ".")
    try: return float(s)
    except: return 0.0

def to_amount(x):
    s = strip_apostrophe(x).replace(" ", "")
    if not s: return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    else:
        s = s.replace(",", ".")
    try: return float(s)
    except: return 0.0

def to_safe_int(x, max_len=15):
    s = strip_apostrophe(x)
    if not s: return ""
    if s.isdigit() and len(s) <= max_len:
        try: return int(s)
        except: pass
    return s

def parse_date(x):
    s = strip_apostrophe(x)
    if not s: return None
    dt = pd.to_datetime(s, utc=True, errors="coerce")
    return None if pd.isna(dt) else dt.date()

def parse_date_ddmmyyyy(x):
    s = strip_apostrophe(x)
    if not s: return ""
    dt = pd.to_datetime(s, utc=True, errors="coerce")
    return "" if pd.isna(dt) else dt.strftime("%d.%m.%Y")

def fmt(x): return f"{x:,.2f}"
def trx_word(n): return "transaction" if n == 1 else "transactions"

def safe_filename(value):
    s = re.sub(r"[^\w\-\. ]+", "_", str(value).strip(), flags=re.UNICODE)
    return re.sub(r"\s+", "_", s) or "merchant"

def normalize_merchant(name):
    raw = (name or "").strip()
    if "temgo" in raw.lower():
        return "afribaba"
    return raw or "unknown"

def find_col(df, variants):
    index = {norm_key(col): col for col in df.columns}
    for v in variants:
        col = index.get(norm_key(v))
        if col: return col
    return None

def build_customer_name(df):
    first_col = find_col(df, ["first_name","customer_first_name","firstname","first name"])
    last_col  = find_col(df, ["last_name","customer_last_name","lastname","last name"])
    first = df[first_col].astype(str).map(strip_apostrophe).str.strip() if first_col else pd.Series("", index=df.index)
    last  = df[last_col].astype(str).map(strip_apostrophe).str.strip()  if last_col  else pd.Series("", index=df.index)
    return (first + " " + last).str.strip()

# =============================================================================
# SUMMARY
# =============================================================================

def _status_lines(grp, status):
    x = grp[grp["status"] == status]
    if x.empty: return 0, []
    cnt = int(x["cnt"].sum())
    if x["cur"].nunique() == 1:
        return cnt, [f"• {fmt(float(x['amt'].sum()))} {x['cur'].iloc[0].upper()}"]
    return cnt, [f"• {fmt(float(r['amt']))} {r['cur'].upper()}" for _, r in x.iterrows()]

def _merchant_block(grp, op_type, title):
    df = grp[grp["op_type"] == op_type].copy()
    succ_cnt, succ_lines = _status_lines(df, "success")
    decl_cnt, decl_lines = _status_lines(df, "decline")
    proc_cnt, proc_lines = _status_lines(df, "processing")
    total = succ_cnt + decl_cnt + proc_cnt
    if total == 0: return "", succ_cnt, decl_cnt, proc_cnt, 0.0
    conv  = succ_cnt / total * 100
    badge = "✅" if conv >= 50 else "⚠️"
    lines = [title]
    if succ_cnt: lines += [f"successful: {succ_cnt} {trx_word(succ_cnt)}"] + succ_lines
    if decl_cnt: lines += [f"declined: {decl_cnt} {trx_word(decl_cnt)}"] + decl_lines
    if proc_cnt: lines += [f"processing: {proc_cnt} {trx_word(proc_cnt)}"] + proc_lines
    lines.append(f"conversion: {conv:.2f}% successful {badge}")
    return "\n".join(lines), succ_cnt, decl_cnt, proc_cnt, conv

def _total_line(tot, status):
    x = tot[tot["status"] == status]
    if x.empty: return "0"
    if x["cur"].nunique() == 1:
        return f"{fmt(float(x['amt'].sum()))} {x['cur'].iloc[0].lower()}"
    return " / ".join(f"{fmt(float(r['amt']))} {r['cur'].lower()}" for _, r in x.iterrows())

def build_report(df, op_filter="deposit"):
    required = ["merchant_name","operation_status","operation_type","amount / amount","currency / currency"]
    payment_col = "payment_type_id / payment_method_type"
    has_pay_col = payment_col in df.columns
    if has_pay_col: required.append(payment_col)
    missing = [c for c in required if c not in df.columns]
    if missing:
        return "❌ Missing columns: " + ", ".join(missing)

    work = df.copy()
    work["merchant"] = work["merchant_name"].astype(str).apply(normalize_merchant).str.strip().str.lower()
    work["_low"]     = work["merchant_name"].astype(str).str.strip().str.lower()
    work             = work[~work["_low"].isin(EXCLUDED_MERCHANTS)].drop(columns=["_low"])
    work["status"]   = work["operation_status"].astype(str).str.strip().str.lower()
    work["op_type"]  = work["operation_type"].astype(str).str.strip().str.lower()
    work["cur"]      = work["currency / currency"].astype(str).str.strip().str.upper()
    work["amt"]      = work["amount / amount"].apply(to_float)
    work["pay_type"] = work[payment_col].astype(str).str.strip().str.lower() if has_pay_col else ""

    label   = "📥 Deposit" if op_filter == "deposit" else "📤 Payout"
    card_op = "sale" if op_filter == "deposit" else "payout"

    card = work[work["pay_type"].isin(CARD_PAY_TYPES) & (work["op_type"] == card_op) & work["status"].isin(SHOW_STATUSES)]
    card_msg = f"💳 card: {label} (file)\n\n"
    if card.empty:
        card_msg += f"No card transactions found."
    else:
        card_tot = card.groupby(["status","cur"], dropna=False).agg(cnt=("amt","size"), amt=("amt","sum")).reset_index()
        total_cnt = int(len(card)); succ_cnt = int(card_tot[card_tot["status"]=="success"]["cnt"].sum())
        decl_cnt = int(card_tot[card_tot["status"]=="decline"]["cnt"].sum())
        proc_cnt = int(card_tot[card_tot["status"]=="processing"]["cnt"].sum())
        card_msg += f"total: {total_cnt} {trx_word(total_cnt)}\n{succ_cnt} success for {_total_line(card_tot,'success')}\n{decl_cnt} decline for {_total_line(card_tot,'decline')}\n{proc_cnt} processing for {_total_line(card_tot,'processing')}\n\n"
        grp = card.groupby(["merchant","op_type","status","cur"], dropna=False).agg(cnt=("amt","size"), amt=("amt","sum")).reset_index()
        order = grp.groupby("merchant")["cnt"].sum().sort_values(ascending=False).index.tolist()
        blocks = []
        for i, merchant in enumerate(order, start=1):
            mm = grp[grp["merchant"]==merchant]
            block, *_ = _merchant_block(mm, card_op, "deposits:" if op_filter=="deposit" else "payouts:")
            if block: blocks.append(f"{i} - {merchant}\n\n{block}")
        card_msg += "\n\n".join(blocks) if blocks else "No blocks found."

    ob_base = work[work["pay_type"].isin(OB_PAY_TYPES)] if has_pay_col and work["pay_type"].isin(OB_PAY_TYPES).any() else work
    if op_filter == "deposit":
        ob = pd.concat([
            ob_base[(ob_base["op_type"]=="payment confirmation") & ob_base["status"].isin({"success","processing"})],
            ob_base[(ob_base["op_type"]=="sale") & (ob_base["status"]=="decline")],
        ], ignore_index=True)
    else:
        ob = ob_base[(ob_base["op_type"]=="payout") & ob_base["status"].isin(SHOW_STATUSES)].copy()

    ob_msg = f"💰 OB: {label} (file)\n\n"
    if ob.empty:
        ob_msg += f"No OB transactions found."
        return (card_msg + "\n\n" + ob_msg).strip()

    ob_tot = ob.groupby(["status","cur"], dropna=False).agg(cnt=("amt","size"), amt=("amt","sum")).reset_index()
    total_cnt = int(len(ob)); succ_cnt = int(ob_tot[ob_tot["status"]=="success"]["cnt"].sum())
    decl_cnt = int(ob_tot[ob_tot["status"]=="decline"]["cnt"].sum())
    proc_cnt = int(ob_tot[ob_tot["status"]=="processing"]["cnt"].sum())
    ob_msg += f"total: {total_cnt} {trx_word(total_cnt)}\n{succ_cnt} success for {_total_line(ob_tot,'success')}\n{decl_cnt} decline for {_total_line(ob_tot,'decline')}\n{proc_cnt} processing for {_total_line(ob_tot,'processing')}\n\n"

    grp2 = ob.groupby(["merchant","op_type","status","cur"], dropna=False).agg(cnt=("amt","size"), amt=("amt","sum")).reset_index()
    order2 = grp2.groupby("merchant")["cnt"].sum().sort_values(ascending=False).index.tolist()
    blocks2 = []
    ob_op = "payment confirmation" if op_filter == "deposit" else "payout"
    for i, merchant in enumerate(order2, start=1):
        mm = grp2[grp2["merchant"]==merchant]
        block, *_ = _merchant_block(mm, ob_op, "deposits:" if op_filter=="deposit" else "payouts:")
        if block: blocks2.append(f"{i} - {merchant}\n\n{block}")
    ob_msg += "\n\n".join(blocks2) if blocks2 else "No blocks found."
    return (card_msg + "\n\n" + ob_msg).strip()

# =============================================================================
# SETTLEMENT FROM TEMPLATE
# =============================================================================

def get_style(ws, row, col):
    c = ws.cell(row=row, column=col)
    return {
        'font': copy_module.copy(c.font),
        'fill': copy_module.copy(c.fill),
        'alignment': copy_module.copy(c.alignment),
        'border': copy_module.copy(c.border),
        'number_format': c.number_format,
    }

def build_settlement_from_template(df: pd.DataFrame, brand: dict, template_bytes: bytes) -> BytesIO:
    import shutil, tempfile, os as _os

    brand_name = brand["name"]
    vol_fee_eur     = float(brand.get("volume_fee_eur", 0))
    vol_fee_non_eur = float(brand.get("volume_fee_non_eur", 0))
    success_fee     = float(brand.get("success_fee", 0))
    failed_fee      = float(brand.get("failed_fee", 0))
    settle_fee_cf   = float(brand.get("settlement_fee_usdt", 0))
    rr_rate         = float(brand.get("rolling_reserve", 0))

    # Filter success only
    sb = df[df["operation_status"].str.strip().str.lower() == "success"].copy()
    if sb.empty:
        raise ValueError("Нет успешных транзакций для этого бренда")

    sb["_date"]   = sb["created_at / operation_created_at"].apply(parse_date)
    sb["_amount"] = sb["amount / amount"].apply(to_amount)
    sb["_cur"]    = sb["currency / currency"].str.strip().str.upper()

    # Daily aggregation
    daily = []
    for d, grp in sb.groupby("_date"):
        eur_g     = grp[grp["_cur"] == "EUR"]
        non_eur_g = grp[grp["_cur"] != "EUR"]
        eur_total     = float(eur_g["_amount"].sum())
        non_eur_total = float(non_eur_g["_amount"].sum())
        grand = eur_total + non_eur_total

        vol_fee = round(eur_total * vol_fee_eur + non_eur_total * vol_fee_non_eur, 6)
        s_fee   = round(len(eur_g) * success_fee, 6)
        f_fee   = round(len(grp) * failed_fee, 6)
        rr      = round(grand * rr_rate, 6)
        settlement = round(grand - vol_fee - s_fee - f_fee - rr, 6)
        cf_fee  = round(settlement * settle_fee_cf, 6)
        close   = round(settlement - cf_fee, 6)

        daily.append({
            "date": datetime.combine(d, datetime.min.time()),
            "eur_tx": len(eur_g), "eur_acq": eur_total, "eur_total": eur_total,
            "non_eur_tx": len(non_eur_g), "non_eur_acq": non_eur_total, "non_eur_total": non_eur_total,
            "grand": grand, "vol_fee": vol_fee, "success_fee": s_fee,
            "failed_fee": f_fee, "cb_refunds": 0, "corrections": 0,
            "rr_daily": rr, "settlement": settlement,
            "settle_fee_rr": 0, "settle_fee_cf": cf_fee, "close_balance": close,
        })

    # Load template
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.write(template_bytes); tmp.close()
    wb = load_workbook(tmp.name)
    _os.unlink(tmp.name)

    DATE_FMT   = "[$-809]dd\\ mmmm\\ yyyy;@"
    AMOUNT_FMT = '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-'
    INT_FMT    = '_-* #,##0_-;\\-* #,##0_-;_-* "-"??_-;_-@_-'

    # ── DAILY REPORT ──
    if "DAILY REPORT" in wb.sheetnames:
        ws_dr = wb["DAILY REPORT"]
        col_map = {
            "date":1,"eur_tx":2,"eur_acq":3,"eur_total":4,
            "non_eur_tx":5,"non_eur_acq":6,"non_eur_total":7,"grand":8,
            "vol_fee":12,"success_fee":13,"failed_fee":14,
            "cb_refunds":15,"corrections":16,"rr_daily":17,
            "settlement":18,"settle_fee_rr":19,"settle_fee_cf":20,"close_balance":21,
        }
        dr_styles = {key: get_style(ws_dr, 6, col) for key, col in col_map.items()}
        for row in ws_dr.iter_rows(min_row=6, max_row=ws_dr.max_row):
            for cell in row: cell.value = None
        for r_idx, d in enumerate(daily, start=6):
            for key, col in col_map.items():
                cell = ws_dr.cell(row=r_idx, column=col)
                cell.value = d[key]
                s = dr_styles[key]
                cell.font = copy_module.copy(s["font"])
                cell.fill = copy_module.copy(s["fill"])
                cell.alignment = copy_module.copy(s["alignment"])
                cell.border = copy_module.copy(s["border"])
                cell.number_format = s["number_format"]
            ws_dr.row_dimensions[r_idx].height = 13.0

    # ── TRANSACTIONS ──
    if "Transactions" in wb.sheetnames:
        ws_tx = wb["Transactions"]
        tx_styles = {i: get_style(ws_tx, 2, i) for i in range(1, 12)}
        for row in ws_tx.iter_rows(min_row=2, max_row=ws_tx.max_row):
            for cell in row: cell.value = None

        col_keys = ["merchant name","external id / payment id","customer name","status",
                    "Eurozone/Non-Eurozone","transaction created at","currency",
                    "Amount","ECB rate","Adjusted ECB","formula"]

        for r_idx, (_, r) in enumerate(sb.iterrows(), start=2):
            cur = r["_cur"]; amt = r["_amount"]
            first = str(r.get("customer_first_name","")).strip()
            last  = str(r.get("customer_last_name","")).strip()
            customer_name = (first + " " + last).strip() or None
            ext_id_raw = str(r.get("external_id / payment_id","")).strip()
            try: ext_id = int(ext_id_raw) if ext_id_raw.isdigit() else ext_id_raw
            except: ext_id = ext_id_raw

            row_vals = {
                "merchant name": brand_name,
                "external id / payment id": ext_id,
                "customer name": customer_name,
                "status": str(r.get("operation_status","")).strip(),
                "Eurozone/Non-Eurozone": "EUR" if cur == "EUR" else "NON-EUR",
                "transaction created at": parse_date(r.get("created_at / operation_created_at","")),
                "currency": cur,
                "Amount": amt,
                "ECB rate": 1,
                "Adjusted ECB": 1,
                "formula": f"=H{r_idx}*J{r_idx}",
            }
            for c_idx, key in enumerate(col_keys, start=1):
                cell = ws_tx.cell(row=r_idx, column=c_idx)
                cell.value = row_vals[key]
                s = tx_styles[c_idx]
                cell.font = copy_module.copy(s["font"])
                cell.fill = copy_module.copy(s["fill"])
                cell.alignment = copy_module.copy(s["alignment"])
                cell.border = copy_module.copy(s["border"])
                if key == "transaction created at": cell.number_format = DATE_FMT
                elif key in ("Amount","formula"): cell.number_format = AMOUNT_FMT
                elif key in ("ECB rate","Adjusted ECB"): cell.number_format = INT_FMT
                else: cell.number_format = s["number_format"]
            ws_tx.row_dimensions[r_idx].height = 13.0

    bio = BytesIO()
    wb.save(bio); bio.seek(0)
    return bio

# =============================================================================
# GOOGLE SHEETS (RECONCILIATION)
# =============================================================================

def _sheets_service():
    if not HAS_GOOGLE:
        raise RuntimeError("google-api-python-client не установлен")
    token_raw = os.getenv("GOOGLE_TOKEN_JSON", "")
    if not token_raw:
        raise RuntimeError("GOOGLE_TOKEN_JSON не задан")
    token_data = json.loads(token_raw)
    creds = Credentials.from_authorized_user_info(token_data, SCOPES)
    return build("sheets", "v4", credentials=creds)

def _get_recon_headers(service):
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{RECON_SHEET_NAME}'!A2:ZZ2",
    ).execute()
    rows = result.get("values", [])
    return rows[0] if rows else []

def _clear_data(service):
    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{RECON_SHEET_NAME}'!A{DATA_START_ROW}:ZZ",
    ).execute()

def _col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s or "A"

def _write_chunk(service, rng, chunk, retries=6):
    delay = 10
    for attempt in range(retries):
        try:
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID, range=rng,
                valueInputOption="USER_ENTERED", body={"values": chunk},
            ).execute(); return
        except Exception as e:
            if any(x in str(e) for x in ("429","RATE_LIMIT","Quota")) and attempt < retries-1:
                time.sleep(delay); delay = min(delay*2, 120)
            else: raise

def _write_all(service, values, start_row, col_count):
    if not values: return
    end_col = _col_letter(col_count)
    idx = 0
    while idx < len(values):
        chunk = values[idx: idx+SHEETS_WRITE_CHUNK]
        r1 = start_row + idx; r2 = r1 + len(chunk) - 1
        rng = f"'{RECON_SHEET_NAME}'!A{r1}:{end_col}{r2}"
        _write_chunk(service, rng, chunk)
        idx += len(chunk)
        if idx < len(values): time.sleep(1.5)

def _prepare_recon(df, headers):
    out = pd.DataFrame(index=df.index)
    first_col   = find_col(df, ["first_name","customer_first_name","firstname","first name"])
    last_col    = find_col(df, ["last_name","customer_last_name","lastname","last name"])
    created_col = find_col(df, ["created_at / operation_created_at","created_at","operation_created_at"])
    amount_col  = find_col(df, ["amount","amount / amount"])
    extid_col   = find_col(df, ["external_id / payment_id","external_id","payment_id"])
    opid_col    = find_col(df, ["id / operation_id","operation_id","id"])

    for h in headers:
        hk = norm_key(h)
        if hk == "customer name":
            first = df[first_col].astype(str).map(strip_apostrophe).str.strip() if first_col else pd.Series("", index=df.index)
            last  = df[last_col].astype(str).map(strip_apostrophe).str.strip()  if last_col  else pd.Series("", index=df.index)
            out[h] = (first + " " + last).str.strip()
        elif hk == norm_key("мой выгрузка"):
            out[h] = df[created_col].apply(parse_date_ddmmyyyy) if created_col else ""
        elif hk == "amount":
            out[h] = df[amount_col].apply(to_amount) if amount_col else 0.0
        elif hk == norm_key("external id payment id"):
            out[h] = df[extid_col].apply(lambda x: to_safe_int(x,15)) if extid_col else ""
        elif hk == norm_key("id / operation_id"):
            out[h] = df[opid_col].apply(lambda x: to_safe_int(x,15)) if opid_col else ""
        else:
            variants = COLUMN_MAP.get(hk, [h])
            col = find_col(df, variants)
            out[h] = df[col].astype(str).map(strip_apostrophe) if col else ""
    return out

def _df_to_values(df):
    df = df.replace({np.nan:"", np.inf:"", -np.inf:""})
    result = []
    for row in df.values.tolist():
        out = []
        for v in row:
            if isinstance(v, float) and (math.isnan(v) or math.isinf(v)): out.append("")
            else: out.append("" if v is None else v)
        result.append(out)
    return result

# =============================================================================
# API ENDPOINTS
# =============================================================================

@app.post("/summary")
async def api_summary(file: UploadFile = File(...), op_type: str = Form("deposit")):
    data = await file.read()
    try: df = read_csv(data)
    except Exception as e: raise HTTPException(400, f"Не смог прочитать CSV: {e}")
    report = build_report(df, op_filter=op_type)
    return {"report": report, "rows": len(df)}

@app.post("/recon")
async def api_recon(file: UploadFile = File(...)):
    if not SPREADSHEET_ID: raise HTTPException(400, "SPREADSHEET_ID не задан")
    data = await file.read()
    try: df = read_csv(data)
    except Exception as e: raise HTTPException(400, f"Не смог прочитать CSV: {e}")
    try:
        service = _sheets_service()
        headers = _get_recon_headers(service)
        prepared = _prepare_recon(df, headers)
        values = _df_to_values(prepared)
        _clear_data(service)
        _write_all(service, values, DATA_START_ROW, len(headers))
    except Exception as e: raise HTTPException(500, str(e))
    return {"ok": True, "rows": len(df)}

@app.post("/settlement-brand")
async def api_settlement_brand(
    file: UploadFile = File(...),
    template: UploadFile = File(...),
    brand_name: str = Form(...),
):
    # Get brand from DB
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT * FROM brands WHERE name=%s", (brand_name,))
    brand = cur.fetchone()
    cur.close(); conn.close()
    if not brand:
        raise HTTPException(404, f"Бренд '{brand_name}' не найден в базе")

    data = await file.read()
    tmpl = await template.read()

    try: df = read_csv(data)
    except Exception as e: raise HTTPException(400, f"Не смог прочитать CSV: {e}")

    # Filter by brand
    merchant_col = find_col(df, ["merchant_name","merchant name","merchant"])
    if not merchant_col: raise HTTPException(400, "Колонка merchant_name не найдена")
    df_brand = df[df[merchant_col].str.strip().str.lower() == brand_name.lower()].copy()
    if df_brand.empty: raise HTTPException(400, f"Транзакции бренда '{brand_name}' не найдены")

    try:
        excel = build_settlement_from_template(df_brand, dict(brand), tmpl)
    except Exception as e:
        raise HTTPException(500, str(e))

    filename = f"settlement_{safe_filename(brand_name)}.xlsx"
    return StreamingResponse(
        excel,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )

@app.post("/settlement-detect")
async def api_settlement_detect(file: UploadFile = File(...)):
    data = await file.read()
    try: df = read_csv(data)
    except Exception as e: raise HTTPException(400, f"Не смог прочитать CSV: {e}")
    merchant_col = find_col(df, ["merchant_name","merchant name","merchant"])
    if not merchant_col: raise HTTPException(400, "Колонка merchant_name не найдена")
    file_merchants = set(df[merchant_col].str.strip().str.lower().unique())
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT * FROM brands ORDER BY name")
    db_brands = cur.fetchall()
    cur.close(); conn.close()
    matched = []
    for b in db_brands:
        if b["name"].lower() in file_merchants:
            if "operation_status" in df.columns:
                mask = (df[merchant_col].str.strip().str.lower() == b["name"].lower()) &                        (df["operation_status"].str.strip().str.lower() == "success")
                count = int(mask.sum())
            else:
                count = int((df[merchant_col].str.strip().str.lower() == b["name"].lower()).sum())
            matched.append({"name": b["name"], "success_count": count})
    return {"brands": matched, "total_rows": len(df)}


@app.post("/settlement-bulk")
async def api_settlement_bulk(
    file: UploadFile = File(...),
    template: UploadFile = File(...),
    brands: str = Form(...),
):
    import zipfile, json as _json
    brand_names = _json.loads(brands)
    if not brand_names: raise HTTPException(400, "Не выбран ни один бренд")
    data = await file.read()
    tmpl = await template.read()
    try: df = read_csv(data)
    except Exception as e: raise HTTPException(400, f"Не смог прочитать CSV: {e}")
    merchant_col = find_col(df, ["merchant_name","merchant name","merchant"])
    if not merchant_col: raise HTTPException(400, "Колонка merchant_name не найдена")
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    placeholders = ",".join(["%s"] * len(brand_names))
    cur.execute(f"SELECT * FROM brands WHERE name IN ({placeholders})", brand_names)
    db_brands = {b["name"]: dict(b) for b in cur.fetchall()}
    cur.close(); conn.close()
    zip_buffer = BytesIO()
    errors = []
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for brand_name in brand_names:
            if brand_name not in db_brands:
                errors.append(f"{brand_name}: не найден в базе"); continue
            brand = db_brands[brand_name]
            df_brand = df[df[merchant_col].str.strip().str.lower() == brand_name.lower()].copy()
            if df_brand.empty:
                errors.append(f"{brand_name}: нет транзакций"); continue
            try:
                excel = build_settlement_from_template(df_brand, brand, tmpl)
                zf.writestr(f"settlement_{safe_filename(brand_name)}.xlsx", excel.read())
            except Exception as e:
                errors.append(f"{brand_name}: {str(e)}")
        if errors:
            zf.writestr("errors.txt", "\n".join(errors))
    zip_buffer.seek(0)
    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={"Content-Disposition": 'attachment; filename="settlement_reports.zip"'},
    )
