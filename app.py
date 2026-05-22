#!/usr/bin/env python3
"""
Typeform → Google Sheets 匯出服務
Railway 部署版（含已購買名單功能）
"""
import io, os, csv, re, time, requests
from datetime import datetime
from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

# ── 環境變數 ──
TYPEFORM_TOKEN  = os.getenv("TYPEFORM_TOKEN", "")
GDRIVE_FOLDER_ID = os.getenv("GDRIVE_FOLDER_ID", "")
GDRIVE_DRIVE_ID  = os.getenv("GDRIVE_DRIVE_ID", "")
API_SECRET       = os.getenv("API_SECRET", "")

PAGE_SIZE = 1000
MAX_PAGES = 30

_cache = {}          # Typeform 匯出快取
_purchased = {       # 已購買名單
    "loaded": False,
    "emails": set(),        # 小寫 email set，用於比對排除
    "phones": set(),        # 正規化後電話 set，用於比對排除
    "email_list": [],       # [(email,)] 用於寫入 sheet
    "phone_list": [],       # [(phone,)] 用於寫入 sheet
}


# ── Google 憑證 ──
def get_google_services():
    import json, tempfile
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    creds_json = os.getenv("GOOGLE_CREDENTIALS", "")
    if not creds_json:
        raise HTTPException(status_code=500, detail="未設定 GOOGLE_CREDENTIALS")

    with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
        f.write(creds_json)
        tmp_path = f.name

    creds = service_account.Credentials.from_service_account_file(
        tmp_path, scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ])
    os.unlink(tmp_path)
    sheets = build("sheets", "v4", credentials=creds)
    drive  = build("drive",  "v3", credentials=creds)
    return sheets, drive


# ── Typeform ──
def fetch_all_responses(form_id):
    all_responses, before_cursor, page_num = [], None, 0
    while page_num < MAX_PAGES:
        page_num += 1
        params = {"page_size": PAGE_SIZE}
        if before_cursor:
            params["before"] = before_cursor
        resp = requests.get(
            f"https://api.typeform.com/forms/{form_id}/responses",
            headers={"Authorization": f"Bearer {TYPEFORM_TOKEN}"},
            params=params, timeout=30)
        resp.raise_for_status()
        items = resp.json().get("items", [])
        all_responses.extend(items)
        print(f"  第 {page_num} 頁：{len(items)} 筆，累計 {len(all_responses)} 筆")
        if len(items) < PAGE_SIZE:
            break
        before_cursor = items[-1]["token"]
    return all_responses


def normalize_phone(raw: str) -> str:
    """將各種格式的電話號碼統一為 09XXXXXXXX 格式"""
    p = raw.strip()
    if p.startswith("+8860"): return "0" + p[5:]
    if p.startswith("+886"):  return "0" + p[4:]
    if p.startswith("886"):   return "0" + p[3:]
    return p


def clean_responses(raw):
    seen_email, seen_phone, cleaned = set(), set(), []
    for entry in raw:
        answers = entry.get("answers") or []
        email = next(
            (a.get("email", "") for a in answers if a.get("type") == "email"), ""
        ).lower()
        phone_raw = next(
            (a.get("phone_number", "") for a in answers if a.get("type") == "phone_number"), ""
        )
        phone = normalize_phone(phone_raw)

        if not email and not phone:
            continue

        is_email_dup = bool(email) and email in seen_email
        is_phone_dup = bool(phone) and phone in seen_phone
        if is_email_dup and is_phone_dup:
            continue

        if email: seen_email.add(email)
        if phone: seen_phone.add(phone)
        cleaned.append({
            "email": "" if is_email_dup else email,
            "phone": "" if is_phone_dup else phone,
        })
    return cleaned


def get_form_title(form_id):
    try:
        r = requests.get(
            f"https://api.typeform.com/forms/{form_id}",
            headers={"Authorization": f"Bearer {TYPEFORM_TOKEN}"}, timeout=10)
        title = r.json().get("title", form_id) if r.ok else form_id
    except:
        title = form_id
    return re.sub(r'[\\/:*?"<>|]', '', title).strip()[:40]


# ── 已購買名單解析 ──
def parse_purchased_csv(content: bytes) -> dict:
    """
    解析已購買名單 CSV。
    欄位：訂單編號, 退款申請狀態, 出貨狀態, 贊助人ID, 贊助人,
          電子信箱(5), 聯絡電話(6), ...
    """
    try:
        text = content.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = content.decode("big5", errors="replace")

    reader = csv.reader(io.StringIO(text))
    headers = []
    email_idx = phone_idx = None
    emails_set, phones_set = set(), set()
    email_list, phone_list = [], []

    for i, row in enumerate(reader):
        if i == 0:
            headers = [h.strip() for h in row]
            # 找欄位索引（容忍欄位名稱前後空白）
            for j, h in enumerate(headers):
                if "電子信箱" in h or "email" in h.lower():
                    email_idx = j
                if "聯絡電話" in h or "phone" in h.lower() or "電話" in h:
                    phone_idx = j
            if email_idx is None and len(headers) > 5:
                email_idx = 5   # fallback：第 6 欄
            if phone_idx is None and len(headers) > 6:
                phone_idx = 6   # fallback：第 7 欄
            continue

        email = row[email_idx].strip().lower() if (email_idx is not None and email_idx < len(row)) else ""
        phone_raw = row[phone_idx].strip() if (phone_idx is not None and phone_idx < len(row)) else ""
        phone = normalize_phone(phone_raw)

        if email and email not in emails_set:
            emails_set.add(email)
            email_list.append((email,))

        if phone and phone not in phones_set:
            phones_set.add(phone)
            phone_list.append((phone,))

    return {
        "emails_set": emails_set,
        "phones_set": phones_set,
        "email_list": email_list,
        "phone_list": phone_list,
    }


# ── Google Sheets 寫入 ──
def create_spreadsheet(drive_service, sheets_service, title):
    """建立試算表，含 4 個分頁：email / phone / 已購買email / 已購買phone"""
    file_metadata = {
        "name": title,
        "mimeType": "application/vnd.google-apps.spreadsheet",
    }
    if GDRIVE_FOLDER_ID:
        file_metadata["parents"] = [GDRIVE_FOLDER_ID]

    result = drive_service.files().create(
        body=file_metadata, supportsAllDrives=True, fields="id"
    ).execute()
    spreadsheet_id = result["id"]

    meta = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    default_sheet_id = meta["sheets"][0]["properties"]["sheetId"]

    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": [
            {"addSheet": {"properties": {"title": "email"}}},
            {"addSheet": {"properties": {"title": "phone"}}},
            {"addSheet": {"properties": {"title": "已購買 email"}}},
            {"addSheet": {"properties": {"title": "已購買 phone"}}},
            {"deleteSheet": {"sheetId": default_sheet_id}},
        ]},
    ).execute()

    return spreadsheet_id


def _batch_write(sheets_service, spreadsheet_id, sheet_title, rows, batch_size=900):
    """分批寫入單一分頁"""
    if not rows:
        return
    for i in range(0, len(rows), batch_size):
        chunk = rows[i:i + batch_size]
        range_name = f"'{sheet_title}'!A{i + 1}"
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption="RAW",
            body={"values": chunk},
        ).execute()
        if i + batch_size < len(rows):
            time.sleep(1)


def _expand_sheet(sheets_service, spreadsheet_id, sheet_map, sheet_title, row_count):
    """當資料超過 1000 列時先擴展格子數"""
    if row_count <= 1000:
        return
    if sheet_title not in sheet_map:
        return
    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": [{
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_map[sheet_title],
                    "gridProperties": {"rowCount": row_count + 100},
                },
                "fields": "gridProperties.rowCount",
            }
        }]},
    ).execute()


def write_to_sheets(sheets_service, spreadsheet_id, cleaned):
    """
    寫入四個分頁：
      email        → Typeform 全部 email（含標題列）
      phone        → Typeform phone，排除已購買名單的電話
      已購買 email  → 已購買名單的 email
      已購買 phone  → 已購買名單的 phone
    """
    purchased_phones = _purchased["phones"]
    purchased_emails = _purchased["emails"]

    # ── Typeform email 列 ──
    email_rows = [["email", "name"]] + [
        (r["email"], i + 1)
        for i, r in enumerate(cleaned) if r["email"]
    ]

    # ── Typeform phone 列（排除已購買電話）──
    phone_rows_all = [r["phone"] for r in cleaned if r["phone"]]
    phone_rows_filtered = [p for p in phone_rows_all if p not in purchased_phones]
    phone_rows = [["phone"]] + [(p,) for p in phone_rows_filtered]

    # ── 已購買名單 email / phone ──
    purchased_email_rows = [["email"]] + list(_purchased["email_list"])
    purchased_phone_rows = [["phone"]]  + list(_purchased["phone_list"])

    # 取得各分頁 sheetId，必要時擴展
    meta = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheet_map = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}

    for title, rows in [
        ("email",        email_rows),
        ("phone",        phone_rows),
        ("已購買 email", purchased_email_rows),
        ("已購買 phone", purchased_phone_rows),
    ]:
        _expand_sheet(sheets_service, spreadsheet_id, sheet_map, title, len(rows))

    # 決定 batch size
    total = len(cleaned)
    batch_size = 900 if total <= 2000 else (700 if total <= 4000 else 500)

    for title, rows in [
        ("email",        email_rows),
        ("phone",        phone_rows),
        ("已購買 email", purchased_email_rows),
        ("已購買 phone", purchased_phone_rows),
    ]:
        _batch_write(sheets_service, spreadsheet_id, title, rows, batch_size)

    return (
        len(email_rows) - 1,
        len(phone_rows) - 1,
        len(purchased_email_rows) - 1,
        len(purchased_phone_rows) - 1,
    )


# ── FastAPI ──
app = FastAPI(title="Typeform Exporter")
app.add_middleware(
    CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"]
)


class ExportRequest(BaseModel):
    form_id: str


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/forms")
def get_forms():
    if not TYPEFORM_TOKEN:
        raise HTTPException(status_code=500, detail="未設定 TYPEFORM_TOKEN")
    resp = requests.get(
        "https://api.typeform.com/forms",
        headers={"Authorization": f"Bearer {TYPEFORM_TOKEN}"},
        params={"page_size": 200}, timeout=15)
    resp.raise_for_status()
    items = resp.json().get("items", [])
    return {"forms": [
        {"id": f["id"], "title": f.get("title", f["id"]),
         "last_updated_at": f.get("last_updated_at", "")}
        for f in items
    ]}


@app.post("/upload-purchased")
async def upload_purchased(file: UploadFile = File(...)):
    """上傳已購買名單 CSV，解析後暫存供本次匯出使用"""
    if not file.filename.lower().endswith(".csv"):
        raise HTTPException(status_code=400, detail="請上傳 CSV 檔案")

    content = await file.read()
    try:
        result = parse_purchased_csv(content)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"CSV 解析失敗：{e}")

    _purchased["loaded"]     = True
    _purchased["emails"]     = result["emails_set"]
    _purchased["phones"]     = result["phones_set"]
    _purchased["email_list"] = result["email_list"]
    _purchased["phone_list"] = result["phone_list"]

    print(f"✅ 已購買名單載入：{len(result['emails_set'])} 封 email，{len(result['phones_set'])} 支電話")
    return {
        "loaded": True,
        "email_count": len(result["emails_set"]),
        "phone_count": len(result["phones_set"]),
    }


@app.get("/purchased-status")
def purchased_status():
    """回傳目前已載入的已購買名單狀態"""
    return {
        "loaded":      _purchased["loaded"],
        "email_count": len(_purchased["emails"]),
        "phone_count": len(_purchased["phones"]),
    }


@app.post("/export")
def export(req: ExportRequest):
    form_id = req.form_id.strip()
    start = time.time()

    try:
        raw     = fetch_all_responses(form_id)
        cleaned = clean_responses(raw)
    except requests.HTTPError as e:
        raise HTTPException(status_code=502, detail=f"Typeform API 錯誤：{e}")

    title = get_form_title(form_id)
    today = datetime.now().strftime("%m%d")
    sheet_name = f"{title}_{today}"

    try:
        sheets_service, drive_service = get_google_services()
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Google 憑證錯誤：{e}")

    try:
        spreadsheet_id = create_spreadsheet(drive_service, sheets_service, sheet_name)
        email_count, phone_count, p_email_count, p_phone_count = \
            write_to_sheets(sheets_service, spreadsheet_id, cleaned)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Google Sheets 寫入錯誤：{e}")

    elapsed = round(time.time() - start, 1)
    sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"
    print(f"🎉 完成！{elapsed}s email:{email_count} phone:{phone_count} "
          f"已購買email:{p_email_count} 已購買phone:{p_phone_count}")

    # 暫存供 CSV 下載
    email_data = [(r["email"], i + 1) for i, r in enumerate(cleaned) if r["email"]]
    purchased_phones_set = _purchased["phones"]
    phone_data = [
        (r["phone"],) for r in cleaned
        if r["phone"] and r["phone"] not in purchased_phones_set
    ]

    _cache[form_id] = {
        "email_rows":    email_data,
        "phone_rows":    phone_data,
        "p_email_rows":  list(_purchased["email_list"]),
        "p_phone_rows":  list(_purchased["phone_list"]),
        "timestamp":     today,
        "title":         title,
    }

    return {
        "form_id":            form_id,
        "form_title":         title,
        "sheet_url":          sheet_url,
        "email_count":        email_count,
        "phone_count":        phone_count,
        "purchased_email_count": p_email_count,
        "purchased_phone_count": p_phone_count,
        "elapsed_seconds":    elapsed,
    }


# ── CSV 下載 ──
def make_csv(rows, headers):
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(headers)
    w.writerows(rows)
    return buf.getvalue().encode("utf-8-sig")


def _csv_response(data: bytes, filename: str):
    return StreamingResponse(
        io.BytesIO(data),
        media_type="text/csv",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{requests.utils.quote(filename)}"},
    )


@app.get("/download/email")
def download_email(form_id: str):
    if form_id not in _cache:
        raise HTTPException(status_code=404, detail="請先執行匯出")
    c = _cache[form_id]
    return _csv_response(
        make_csv(c["email_rows"], ["email", "name"]),
        f"{c['title']}_email_{c['timestamp']}.csv",
    )


@app.get("/download/phone")
def download_phone(form_id: str):
    if form_id not in _cache:
        raise HTTPException(status_code=404, detail="請先執行匯出")
    c = _cache[form_id]
    return _csv_response(
        make_csv(c["phone_rows"], ["phone"]),
        f"{c['title']}_phone_{c['timestamp']}.csv",
    )


@app.get("/download/purchased-email")
def download_purchased_email(form_id: str):
    if form_id not in _cache:
        raise HTTPException(status_code=404, detail="請先執行匯出")
    c = _cache[form_id]
    return _csv_response(
        make_csv(c["p_email_rows"], ["email"]),
        f"{c['title']}_已購買email_{c['timestamp']}.csv",
    )


@app.get("/download/purchased-phone")
def download_purchased_phone(form_id: str):
    if form_id not in _cache:
        raise HTTPException(status_code=404, detail="請先執行匯出")
    c = _cache[form_id]
    return _csv_response(
        make_csv(c["p_phone_rows"], ["phone"]),
        f"{c['title']}_已購買phone_{c['timestamp']}.csv",
    )


# Static files（前端）— 放在所有 API route 之後
if os.path.exists("static"):
    app.mount("/", StaticFiles(directory="static", html=True), name="static")


if __name__ == "__main__":
    import uvicorn
    print("\n🚀 Server 啟動中...")
    print(f"  http://localhost:8080")
    print(f"  TYPEFORM_TOKEN：{'已設定 ✅' if TYPEFORM_TOKEN else '未設定 ❌'}")
    print(f"  GOOGLE_CREDENTIALS：{'已設定 ✅' if os.getenv('GOOGLE_CREDENTIALS') else '未設定 ❌'}")
    print(f"  GDRIVE_FOLDER_ID：{GDRIVE_FOLDER_ID or '未設定 ❌'}")
    uvicorn.run(app, host="0.0.0.0", port=8080)
