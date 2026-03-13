#!/usr/bin/env python3
"""
Typeform → Google Sheets 匯出服務
Railway 部署版
"""

import io, os, csv, re, time, requests
from datetime import datetime
from fastapi import FastAPI, HTTPException
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
_cache = {}

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
        print(f"   第 {page_num} 頁：{len(items)} 筆，累計 {len(all_responses)} 筆")
        if len(items) < PAGE_SIZE:
            break
        before_cursor = items[-1]["token"]
    return all_responses

def clean_responses(raw):
    seen_email, seen_phone, cleaned = set(), set(), []
    for entry in raw:
        answers = entry.get("answers", [])
        email = next((a.get("email","") for a in answers if a.get("type")=="email"), "").lower()
        phone_raw = next((a.get("phone_number","") for a in answers if a.get("type")=="phone_number"), "")
        if phone_raw.startswith("+8860"):   phone = "0" + phone_raw[5:]
        elif phone_raw.startswith("+886"):  phone = "0" + phone_raw[4:]
        elif phone_raw.startswith("886"):   phone = "0" + phone_raw[3:]
        else:                               phone = phone_raw
        if not email and not phone: continue
        is_email_dup = bool(email) and email in seen_email
        is_phone_dup = bool(phone) and phone in seen_phone
        if is_email_dup and is_phone_dup: continue
        if email: seen_email.add(email)
        if phone: seen_phone.add(phone)
        cleaned.append({"email": "" if is_email_dup else email, "phone": "" if is_phone_dup else phone})
    return cleaned

def get_form_title(form_id):
    try:
        r = requests.get(f"https://api.typeform.com/forms/{form_id}",
            headers={"Authorization": f"Bearer {TYPEFORM_TOKEN}"}, timeout=10)
        title = r.json().get("title", form_id) if r.ok else form_id
    except:
        title = form_id
    return re.sub(r'[\\/:*?"<>|]', '', title).strip()[:40]

# ── Google Sheets 寫入 ──
def create_spreadsheet(drive_service, sheets_service, title):
    # 透過 Drive API 直接在共用雲端硬碟建立試算表
    file_metadata = {
        'name': title,
        'mimeType': 'application/vnd.google-apps.spreadsheet',
    }
    if GDRIVE_FOLDER_ID:
        file_metadata['parents'] = [GDRIVE_FOLDER_ID]
    result = drive_service.files().create(
        body=file_metadata, supportsAllDrives=True, fields='id'
    ).execute()
    spreadsheet_id = result['id']

    # 新增 email / phone 分頁，刪除預設 Sheet1
    sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={
        "requests": [
            {"addSheet": {"properties": {"title": "email"}}},
            {"addSheet": {"properties": {"title": "phone"}}},
            {"deleteSheet": {"sheetId": 0}},
        ]
    }).execute()
    return spreadsheet_id

def write_to_sheets(sheets_service, spreadsheet_id, cleaned):
    email_rows = [["email", "name"]] + [(r["email"], i+1) for i, r in enumerate(cleaned) if r["email"]]
    phone_rows = [["phone"]] + [(r["phone"],) for r in cleaned if r["phone"]]

    # 先取得 sheet ID，再擴展行數以容納所有資料
    meta = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheet_map = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}

    expand_requests = []
    for sheet_title, rows in [("email", email_rows), ("phone", phone_rows)]:
        if len(rows) > 1000 and sheet_title in sheet_map:
            expand_requests.append({
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": sheet_map[sheet_title],
                        "gridProperties": {"rowCount": len(rows) + 100}
                    },
                    "fields": "gridProperties.rowCount"
                }
            })
    if expand_requests:
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={"requests": expand_requests}
        ).execute()

    total = len(cleaned)
    batch_size = 900 if total <= 2000 else (700 if total <= 4000 else 500)

    for sheet_title, rows in [("email", email_rows), ("phone", phone_rows)]:
        for i in range(0, len(rows), batch_size):
            chunk = rows[i:i+batch_size]
            range_name = f"{sheet_title}!A{i+1}"
            sheets_service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption="RAW",
                body={"values": chunk},
            ).execute()
            if i + batch_size < len(rows):
                time.sleep(1)

    return len(email_rows) - 1, len(phone_rows) - 1

# ── FastAPI ──
app = FastAPI(title="Typeform Exporter")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

class ExportRequest(BaseModel):
    form_id: str

@app.get("/health")
def health():
    return {"status": "ok"}

@app.get("/forms")
def get_forms():
    if not TYPEFORM_TOKEN:
        raise HTTPException(status_code=500, detail="未設定 TYPEFORM_TOKEN")
    resp = requests.get("https://api.typeform.com/forms",
        headers={"Authorization": f"Bearer {TYPEFORM_TOKEN}"},
        params={"page_size": 15}, timeout=15)
    resp.raise_for_status()
    items = resp.json().get("items", [])
    return {"forms": [{"id": f["id"], "title": f.get("title", f["id"]),
                       "last_updated_at": f.get("last_updated_at","")} for f in items]}

@app.post("/export")
def export(req: ExportRequest):
    form_id = req.form_id.strip()
    start = time.time()

    try:
        raw     = fetch_all_responses(form_id)
        cleaned = clean_responses(raw)
    except requests.HTTPError as e:
        raise HTTPException(status_code=502, detail=f"Typeform API 錯誤：{e}")

    title     = get_form_title(form_id)
    today     = datetime.now().strftime("%m%d")
    sheet_name = f"{title}_{today}"

    sheets_service, drive_service = get_google_services()
    spreadsheet_id = create_spreadsheet(drive_service, sheets_service, sheet_name)
    email_count, phone_count = write_to_sheets(sheets_service, spreadsheet_id, cleaned)

    elapsed   = round(time.time() - start, 1)
    sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"
    print(f"🎉 完成！{elapsed}s  email:{email_count}  phone:{phone_count}")

    # 暫存供 CSV 下載
    email_data = [(r["email"], i+1) for i, r in enumerate(cleaned) if r["email"]]
    phone_data = [(r["phone"],) for r in cleaned if r["phone"]]
    _cache[form_id] = {"email_rows": email_data, "phone_rows": phone_data, "timestamp": today, "title": title}

    return {"form_id": form_id, "form_title": title, "sheet_url": sheet_url,
            "email_count": email_count, "phone_count": phone_count,
            "elapsed_seconds": elapsed}

def make_csv(rows, headers):
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(headers)
    w.writerows(rows)
    return buf.getvalue().encode("utf-8-sig")

@app.get("/download/email")
def download_email(form_id: str):
    if form_id not in _cache: raise HTTPException(status_code=404, detail="請先執行匯出")
    c = _cache[form_id]
    fname = f"{c['title']}_email_{c['timestamp']}.csv"
    return StreamingResponse(io.BytesIO(make_csv(c["email_rows"], ["email","name"])),
        media_type="text/csv", headers={"Content-Disposition": f"attachment; filename*=UTF-8''{requests.utils.quote(fname)}"})

@app.get("/download/phone")
def download_phone(form_id: str):
    if form_id not in _cache: raise HTTPException(status_code=404, detail="請先執行匯出")
    c = _cache[form_id]
    fname = f"{c['title']}_phone_{c['timestamp']}.csv"
    return StreamingResponse(io.BytesIO(make_csv(c["phone_rows"], ["phone"])),
        media_type="text/csv", headers={"Content-Disposition": f"attachment; filename*=UTF-8''{requests.utils.quote(fname)}"})

# Static files（前端）— 放在所有 API route 之後
if os.path.exists("static"):
    app.mount("/", StaticFiles(directory="static", html=True), name="static")

if __name__ == "__main__":
    import uvicorn
    print("\n🚀 Server 啟動中...")
    print(f"   http://localhost:8080")
    print(f"   TYPEFORM_TOKEN：{'已設定 ✅' if TYPEFORM_TOKEN else '未設定 ❌'}")
    print(f"   GOOGLE_CREDENTIALS：{'已設定 ✅' if os.getenv('GOOGLE_CREDENTIALS') else '未設定 ❌'}")
    print(f"   GDRIVE_FOLDER_ID：{GDRIVE_FOLDER_ID or '未設定 ❌'}")
    uvicorn.run(app, host="0.0.0.0", port=8080)
