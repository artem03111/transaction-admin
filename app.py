"""
Transaction Admin — FastAPI backend
Deploy to Render.com (free tier) or any Python host.

ENV VARS needed on Render:
  SPREADSHEET_ID   — Google Sheet ID
  GOOGLE_TOKEN_JSON — contents of token.json (paste the JSON as a single-line string)
  ALLOWED_ORIGIN   — your frontend URL (or * for testing)
"""

import json
import math
import os
import re
import time
from io import BytesIO

import numpy as np
import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from openpyxl import Workbook
from openpyxl.styles import Font

# ── optional Google Sheets (only if GOOGLE_TOKEN_JSON is set) ─────────────────
try:
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
    HAS_GOOGLE = True
except ImportError:
    HAS_GOOGLE = False

# =============================================================================
# CONFIG  (override via env vars on Render)
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

SETTLEMENT_MERCHANTS = [
    "ORION TRANSACT INC",
    "Trueluck",
    "Undress",
    "Sultangames",
    "Nomadgames",
]

SETTLEMENT_HEADERS = [
    "merchant name", "external id / payment id", "customer name", "status",
    "Eurozone/Non-Eurozone", "transaction created at", "currency",
    "Amount", "ECB rate", "Adjusted ECB", "EUR amount",
]

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
# APP
# =============================================================================

app = FastAPI(title="Transaction Admin API")

allowed = os.getenv("ALLOWED_ORIGIN", "*")
app.add_middleware(
    CORSMiddleware,
    allow_origins=[allowed] if allowed != "*" else ["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def root():
    return {"status": "ok", "merchants": SETTLEMENT_MERCHANTS}

@app.get("/merchants")
def merchants():
    return SETTLEMENT_MERCHANTS

# =============================================================================
# HELPERS
# =============================================================================

def norm_key(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "").strip().lower())

def strip_apostrophe(value) -> str:
    s = str(value or "").strip()
    return s[1:].lstrip() if s.startswith("'") else s

def read_csv(data: bytes) -> pd.DataFrame:
    df = pd.read_csv(BytesIO(data), sep=";", dtype=str, keep_default_na=False)
    if df.shape[1] <= 1:
        df = pd.read_csv(BytesIO(data), sep=",", dtype=str, keep_default_na=False)
    return df

def to_float(x) -> float:
    s = str(x or "").strip().replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0

def to_amount(x) -> float:
    s = strip_apostrophe(x).replace(" ", "")
    if not s:
        return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0

def to_safe_int(x, max_len: int = 15):
    s = strip_apostrophe(x)
    if not s:
        return ""
    if s.isdigit() and len(s) <= max_len:
        try:
            return int(s)
        except Exception:
            pass
    return s

def parse_date(x):
    s = strip_apostrophe(x)
    if not s:
        return None
    dt = pd.to_datetime(s, utc=True, errors="coerce")
    return None if pd.isna(dt) else dt.date()

def parse_date_ddmmyyyy(x: str) -> str:
    s = strip_apostrophe(x)
    if not s:
        return ""
    dt = pd.to_datetime(s, utc=True, errors="coerce")
    return "" if pd.isna(dt) else dt.strftime("%d.%m.%Y")

def fmt(x: float) -> str:
    return f"{x:,.2f}"

def trx_word(n: int) -> str:
    return "transaction" if n == 1 else "transactions"

def safe_filename(value: str) -> str:
    s = re.sub(r"[^\w\-\. ]+", "_", str(value).strip(), flags=re.UNICODE)
    return re.sub(r"\s+", "_", s) or "merchant"

def normalize_merchant(name: str) -> str:
    raw = (name or "").strip()
    if "temgo" in raw.lower() and "groupe" in raw.lower():
        return "afribaba"
    return raw or "unknown"

def find_col(df: pd.DataFrame, variants: list) -> str | None:
    index = {norm_key(col): col for col in df.columns}
    for v in variants:
        col = index.get(norm_key(v))
        if col:
            return col
    return None

def find_merchant_col(df: pd.DataFrame) -> str | None:
    return find_col(df, ["merchant_name", "merchant name", "merchant"])

def build_customer_name(df: pd.DataFrame) -> pd.Series:
    first_col = find_col(df, ["first_name", "customer_first_name", "firstname", "first name"])
    last_col  = find_col(df, ["last_name",  "customer_last_name",  "lastname",  "last name"])
    first = df[first_col].astype(str).map(strip_apostrophe).str.strip() if first_col else pd.Series("", index=df.index)
    last  = df[last_col].astype(str).map(strip_apostrophe).str.strip()  if last_col  else pd.Series("", index=df.index)
    return (first + " " + last).str.strip()

def filter_by_merchant(df: pd.DataFrame, merchant_name: str) -> pd.DataFrame:
    col = find_merchant_col(df)
    if not col:
        raise ValueError(f"Колонка мерчанта не найдена. Колонки: {', '.join(df.columns[:30])}")
    mask = df[col].astype(str).str.strip().str.lower() == merchant_name.strip().lower()
    return df[mask].copy()

def filter_by_op_type(df: pd.DataFrame, op_type: str) -> pd.DataFrame:
    col = find_col(df, ["operation_type", "operation type", "type"])
    if not col:
        return df
    target = "sale" if op_type == "deposit" else "payout"
    return df[df[col].astype(str).str.strip().str.lower() == target].copy()

# =============================================================================
# SUMMARY
# =============================================================================

def _status_lines(grp, status):
    x = grp[grp["status"] == status]
    if x.empty:
        return 0, []
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
    if total == 0:
        return "", succ_cnt, decl_cnt, proc_cnt, 0.0
    conv  = succ_cnt / total * 100
    badge = "✅" if conv >= 50 else "⚠️"
    lines = [title]
    if succ_cnt: lines += [f"successful: {succ_cnt} {trx_word(succ_cnt)}"] + succ_lines
    if decl_cnt: lines += [f"declined: {decl_cnt} {trx_word(decl_cnt)}"]   + decl_lines
    if proc_cnt: lines += [f"processing: {proc_cnt} {trx_word(proc_cnt)}"] + proc_lines
    lines.append(f"conversion: {conv:.2f}% successful {badge}")
    return "\n".join(lines), succ_cnt, decl_cnt, proc_cnt, conv

def _total_line(tot, status):
    x = tot[tot["status"] == status]
    if x.empty:
        return "0"
    if x["cur"].nunique() == 1:
        return f"{fmt(float(x['amt'].sum()))} {x['cur'].iloc[0].lower()}"
    return " / ".join(f"{fmt(float(r['amt']))} {r['cur'].lower()}" for _, r in x.iterrows())

def build_report(df: pd.DataFrame, op_filter: str = "deposit") -> str:
    required = ["merchant_name", "operation_status", "operation_type", "amount / amount", "currency / currency"]
    payment_col = "payment_type_id / payment_method_type"
    has_pay_col = payment_col in df.columns
    if has_pay_col:
        required.append(payment_col)

    missing = [c for c in required if c not in df.columns]
    if missing:
        return "❌ Missing columns: " + ", ".join(missing) + "\n\n✅ Found columns:\n" + "\n".join(list(df.columns)[:60])

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
        card_msg += f"No card transactions found for operation_type '{card_op}' with statuses success/decline/processing."
    else:
        card_tot = (card.groupby(["status", "cur"], dropna=False).agg(cnt=("amt", "size"), amt=("amt", "sum")).reset_index().sort_values(["status", "cur"]))
        total_cnt = int(len(card))
        succ_cnt  = int(card_tot[card_tot["status"] == "success"]["cnt"].sum())
        decl_cnt  = int(card_tot[card_tot["status"] == "decline"]["cnt"].sum())
        proc_cnt  = int(card_tot[card_tot["status"] == "processing"]["cnt"].sum())
        card_msg += (f"total: {total_cnt} {trx_word(total_cnt)}\n{succ_cnt} success for {_total_line(card_tot, 'success')}\n{decl_cnt} decline for {_total_line(card_tot, 'decline')}\n{proc_cnt} processing for {_total_line(card_tot, 'processing')}\n\n")
        grp = (card.groupby(["merchant", "op_type", "status", "cur"], dropna=False).agg(cnt=("amt", "size"), amt=("amt", "sum")).reset_index())
        order = grp.groupby("merchant")["cnt"].sum().sort_values(ascending=False).index.tolist()
        blocks = []
        for i, merchant in enumerate(order, start=1):
            mm = grp[grp["merchant"] == merchant]
            title = "deposits:" if op_filter == "deposit" else "payouts:"
            block, *_ = _merchant_block(mm, card_op, title)
            if block:
                blocks.append(f"{i} - {merchant}\n\n{block}")
        card_msg += "\n\n".join(blocks) if blocks else f"No non-zero card {card_op} blocks found."

    ob_base = work[work["pay_type"].isin(OB_PAY_TYPES)] if has_pay_col and work["pay_type"].isin(OB_PAY_TYPES).any() else work
    if op_filter == "deposit":
        ob = pd.concat([
            ob_base[(ob_base["op_type"] == "payment confirmation") & ob_base["status"].isin({"success", "processing"})],
            ob_base[(ob_base["op_type"] == "sale") & (ob_base["status"] == "decline")],
        ], ignore_index=True)
    else:
        ob = ob_base[(ob_base["op_type"] == "payout") & ob_base["status"].isin(SHOW_STATUSES)].copy()

    ob_msg = f"💰 OB: {label} (file)\n\n"
    if ob.empty:
        ob_msg += f"No OB {op_filter} transactions found."
        return (card_msg + "\n\n" + ob_msg).strip()

    ob_tot = ob.groupby(["status", "cur"], dropna=False).agg(cnt=("amt", "size"), amt=("amt", "sum")).reset_index().sort_values(["status", "cur"])
    total_cnt = int(len(ob))
    succ_cnt  = int(ob_tot[ob_tot["status"] == "success"]["cnt"].sum())
    decl_cnt  = int(ob_tot[ob_tot["status"] == "decline"]["cnt"].sum())
    proc_cnt  = int(ob_tot[ob_tot["status"] == "processing"]["cnt"].sum())
    ob_msg += (f"total: {total_cnt} {trx_word(total_cnt)}\n{succ_cnt} success for {_total_line(ob_tot, 'success')}\n{decl_cnt} decline for {_total_line(ob_tot, 'decline')}\n{proc_cnt} processing for {_total_line(ob_tot, 'processing')}\n\n")

    grp2 = ob.groupby(["merchant", "op_type", "status", "cur"], dropna=False).agg(cnt=("amt", "size"), amt=("amt", "sum")).reset_index()
    order2 = grp2.groupby("merchant")["cnt"].sum().sort_values(ascending=False).index.tolist()
    blocks2 = []
    ob_op = "payment confirmation" if op_filter == "deposit" else "payout"
    for i, merchant in enumerate(order2, start=1):
        mm = grp2[grp2["merchant"] == merchant]
        title = "deposits:" if op_filter == "deposit" else "payouts:"
        block, *_ = _merchant_block(mm, ob_op, title)
        if block:
            blocks2.append(f"{i} - {merchant}\n\n{block}")
    ob_msg += "\n\n".join(blocks2) if blocks2 else f"No non-zero OB {op_filter} blocks found."

    return (card_msg + "\n\n" + ob_msg).strip()

# =============================================================================
# SETTLEMENT EXCEL
# =============================================================================

def _autosize(ws):
    for col in ws.columns:
        width = max((len(str(c.value or "")) for c in col), default=0)
        ws.column_dimensions[col[0].column_letter].width = min(width + 2, 30)

def _prepare_settlement(df: pd.DataFrame, merchant: str) -> pd.DataFrame:
    extid_col   = find_col(df, ["external_id / payment_id", "external_id", "payment_id"])
    status_col  = find_col(df, ["operation_status", "status", "transaction status"])
    created_col = find_col(df, ["created_at / operation_created_at", "created_at", "operation_created_at"])
    cur_col     = find_col(df, ["currency / currency", "currency"])
    amount_col  = find_col(df, ["amount / amount", "amount"])
    missing = [n for n, c in [("external_id / payment_id", extid_col), ("operation_status", status_col), ("created_at", created_col), ("currency", cur_col), ("amount", amount_col)] if not c]
    if missing:
        raise ValueError("Не хватает колонок: " + ", ".join(missing))

    out = pd.DataFrame(index=df.index)
    out["merchant name"]              = merchant
    out["external id / payment id"]   = df[extid_col].map(strip_apostrophe).astype(str).str.strip()
    out["customer name"]              = build_customer_name(df)
    out["status"]                     = df[status_col].map(strip_apostrophe).astype(str).str.strip()
    cur_s = df[cur_col].astype(str).str.strip().str.upper()
    out["Eurozone/Non-Eurozone"]      = np.where(cur_s == "", "", np.where(cur_s == "EUR", "EUR", "NON-EUR"))
    out["transaction created at"]     = df[created_col].apply(parse_date)
    out["currency"]                   = cur_s
    out["Amount"]                     = df[amount_col].apply(to_amount)
    out["ECB rate"]                   = 1
    out["Adjusted ECB"]               = 1
    out["EUR amount"]                 = out["Amount"] * out["Adjusted ECB"]
    before = len(out)
    out = out.drop_duplicates()
    out = out.drop_duplicates(subset=["external id / payment id", "status", "transaction created at", "currency", "Amount"], keep="first")
    out = out.reset_index(drop=True)
    return out[SETTLEMENT_HEADERS]

def build_settlement_excel(df: pd.DataFrame, merchant: str, op_type: str) -> BytesIO:
    prepared = _prepare_settlement(df, merchant)
    wb = Workbook()
    ws = wb.active
    ws.title = "Transactions"
    for i, h in enumerate(SETTLEMENT_HEADERS, start=1):
        cell = ws.cell(row=1, column=i, value=h)
        cell.font = Font(bold=True)
    for r, (_, row) in enumerate(prepared.iterrows(), start=2):
        ws.cell(r, 1).value = row["merchant name"]
        ws.cell(r, 2).value = row["external id / payment id"]
        ws.cell(r, 3).value = row["customer name"]
        ws.cell(r, 4).value = row["status"]
        ws.cell(r, 5).value = row["Eurozone/Non-Eurozone"]
        d = ws.cell(r, 6)
        d.value = row["transaction created at"]
        d.number_format = "DD.MM.YYYY"
        ws.cell(r, 7).value = row["currency"]
        for col, key, fmt_code in [(8, "Amount", "0.00"), (9, "ECB rate", "0.00"), (10, "Adjusted ECB", "0.00")]:
            c = ws.cell(r, col)
            c.value = row[key]
            c.number_format = fmt_code
        eur = ws.cell(r, 11)
        eur.value = f"=H{r}*J{r}"
        eur.number_format = "0.00"
    ws.freeze_panes = "A2"
    _autosize(ws)
    bio = BytesIO()
    suffix = "deposit" if op_type == "deposit" else "payout"
    bio.name = f"settlement_{safe_filename(merchant)}_{suffix}.xlsx"
    wb.save(bio)
    bio.seek(0)
    return bio

# =============================================================================
# GOOGLE SHEETS (RECONCILIATION)
# =============================================================================

def _sheets_service():
    if not HAS_GOOGLE:
        raise RuntimeError("google-api-python-client не установлен")
    token_raw = os.getenv("GOOGLE_TOKEN_JSON", "")
    if not token_raw:
        raise RuntimeError("GOOGLE_TOKEN_JSON не задан в переменных окружения Render")
    token_data = json.loads(token_raw)
    creds = Credentials.from_authorized_user_info(token_data, SCOPES)
    return build("sheets", "v4", credentials=creds)

def _get_recon_headers(service) -> list:
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{RECON_SHEET_NAME}'!A2:ZZ2",
    ).execute()
    rows = result.get("values", [])
    return rows[0] if rows else []

def _get_sheet_id(service) -> int:
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for s in meta.get("sheets", []):
        if s["properties"]["title"] == RECON_SHEET_NAME:
            return s["properties"]["sheetId"]
    raise ValueError(f"Лист '{RECON_SHEET_NAME}' не найден")

def _clear_data(service):
    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{RECON_SHEET_NAME}'!A{DATA_START_ROW}:ZZ",
    ).execute()

def _col_letter(n: int) -> str:
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
            ).execute()
            return
        except Exception as e:
            err = str(e)
            is_quota = any(x in err for x in ("429", "RATE_LIMIT_EXCEEDED", "Quota exceeded"))
            if is_quota and attempt < retries - 1:
                time.sleep(delay)
                delay = min(delay * 2, 120)
            else:
                raise

def _write_all(service, values, start_row, col_count):
    if not values:
        return
    end_col = _col_letter(col_count)
    idx = 0
    while idx < len(values):
        chunk = values[idx: idx + SHEETS_WRITE_CHUNK]
        r1 = start_row + idx
        r2 = r1 + len(chunk) - 1
        rng = f"'{RECON_SHEET_NAME}'!A{r1}:{end_col}{r2}"
        _write_chunk(service, rng, chunk)
        idx += len(chunk)
        if idx < len(values):
            time.sleep(1.5)

def _prepare_recon(df: pd.DataFrame, headers: list) -> pd.DataFrame:
    out = pd.DataFrame(index=df.index)
    first_col   = find_col(df, ["first_name", "customer_first_name", "firstname", "first name"])
    last_col    = find_col(df, ["last_name",  "customer_last_name",  "lastname",  "last name"])
    created_col = find_col(df, ["created_at / operation_created_at", "created_at", "operation_created_at"])
    amount_col  = find_col(df, ["amount", "amount / amount"])
    extid_col   = find_col(df, ["external_id / payment_id", "external_id", "payment_id"])
    opid_col    = find_col(df, ["id / operation_id", "operation_id", "id"])

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
            out[h] = df[extid_col].apply(lambda x: to_safe_int(x, 15)) if extid_col else ""
        elif hk == norm_key("id / operation_id"):
            out[h] = df[opid_col].apply(lambda x: to_safe_int(x, 15)) if opid_col else ""
        else:
            variants = COLUMN_MAP.get(hk, [h])
            col = find_col(df, variants)
            out[h] = df[col].astype(str).map(strip_apostrophe) if col else ""
    return out

def _df_to_values(df: pd.DataFrame) -> list:
    df = df.replace({np.nan: "", np.inf: "", -np.inf: ""})
    result = []
    for row in df.values.tolist():
        out = []
        for v in row:
            if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
                out.append("")
            else:
                out.append("" if v is None else v)
        result.append(out)
    return result

def reconcile_upload(df: pd.DataFrame):
    service  = _sheets_service()
    headers  = _get_recon_headers(service)
    prepared = _prepare_recon(df, headers)
    values   = _df_to_values(prepared)
    sheet_id = _get_sheet_id(service)
    _clear_data(service)
    _write_all(service, values, DATA_START_ROW, len(headers))

# =============================================================================
# API ENDPOINTS
# =============================================================================

@app.post("/summary")
async def api_summary(file: UploadFile = File(...), op_type: str = Form("deposit")):
    data = await file.read()
    try:
        df = read_csv(data)
    except Exception as e:
        raise HTTPException(400, f"Не смог прочитать CSV: {e}")
    report = build_report(df, op_filter=op_type)
    return {"report": report, "rows": len(df)}

@app.post("/recon")
async def api_recon(file: UploadFile = File(...)):
    if not SPREADSHEET_ID:
        raise HTTPException(400, "SPREADSHEET_ID не задан на сервере")
    data = await file.read()
    try:
        df = read_csv(data)
    except Exception as e:
        raise HTTPException(400, f"Не смог прочитать CSV: {e}")
    try:
        reconcile_upload(df)
    except Exception as e:
        raise HTTPException(500, str(e))
    return {"ok": True, "rows": len(df)}

@app.post("/settlement")
async def api_settlement(file: UploadFile = File(...), merchant: str = Form(...), op_type: str = Form("deposit")):
    if merchant not in SETTLEMENT_MERCHANTS:
        raise HTTPException(400, f"Мерчант '{merchant}' не в списке")
    data = await file.read()
    try:
        df = read_csv(data)
    except Exception as e:
        raise HTTPException(400, f"Не смог прочитать CSV: {e}")
    filtered = filter_by_merchant(df, merchant)
    if filtered.empty:
        raise HTTPException(400, f"Транзакции мерчанта '{merchant}' не найдены в файле")
    filtered = filter_by_op_type(filtered, op_type)
    if filtered.empty:
        raise HTTPException(400, f"После фильтра по {op_type} строк не осталось")
    try:
        excel = build_settlement_excel(filtered, merchant, op_type)
    except Exception as e:
        raise HTTPException(500, str(e))
    suffix = "deposit" if op_type == "deposit" else "payout"
    filename = f"settlement_{safe_filename(merchant)}_{suffix}.xlsx"
    return StreamingResponse(
        excel,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
