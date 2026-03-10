#!/usr/bin/env python3
import io, os, csv, time, re, requests
from datetime import datetime
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

TYPEFORM_TOKEN = os.getenv("TYPEFORM_TOKEN", "")
PAGE_SIZE = 1000
MAX_PAGES = 30
_cache = {}

app = FastAPI()
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

class ExportRequest(BaseModel):
    form_id: str

def fetch_all_responses(form_id):
    all_responses, before_cursor, page_num = [], None, 0
    while page_num < MAX_PAGES:
        page_num += 1
        params = {"page_size": PAGE_SIZE}
        if before_cursor:
            params["before"] = before_cursor
        resp = requests.get(f"https://api.typeform.com/forms/{form_id}/responses",
            headers={"Authorization": f"Bearer {TYPEFORM_TOKEN}"}, params=params, timeout=30)
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
        if phone_raw.startswith("+8860"): phone = "0" + phone_raw[5:]
        elif phone_raw.startswith("+886"): phone = "0" + phone_raw[4:]
        elif phone_raw.startswith("886"): phone = "0" + phone_raw[3:]
        else: phone = phone_raw
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

@app.get("/forms")
def get_forms():
    if not TYPEFORM_TOKEN:
        raise HTTPException(status_code=500, detail="未設定 TYPEFORM_TOKEN")
    resp = requests.get("https://api.typeform.com/forms",
        headers={"Authorization": f"Bearer {TYPEFORM_TOKEN}"},
        params={"page_size": 15}, timeout=15)
    resp.raise_for_status()
    items = resp.json().get("items", [])
    return {"forms": [{"id": f["id"], "title": f.get("title", f["id"]), "last_updated_at": f.get("last_updated_at","")} for f in items]}

@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/export")
def export(req: ExportRequest):
    form_id = req.form_id.strip()
    if not TYPEFORM_TOKEN:
        raise HTTPException(status_code=500, detail="未設定 TYPEFORM_TOKEN")
    start = time.time()
    try:
        raw = fetch_all_responses(form_id)
        cleaned = clean_responses(raw)
    except requests.HTTPError as e:
        raise HTTPException(status_code=502, detail=f"Typeform API 錯誤：{e}")
    today = datetime.now().strftime("%m%d")
    email_rows = [(row["email"], i+1) for i, row in enumerate(cleaned) if row["email"]]
    phone_rows = [(row["phone"],) for row in cleaned if row["phone"]]
    title = get_form_title(form_id)
    _cache[form_id] = {"email_rows": email_rows, "phone_rows": phone_rows, "timestamp": today, "title": title}
    elapsed = round(time.time() - start, 1)
    print(f"🎉 完成！{elapsed}s  email:{len(email_rows)}  phone:{len(phone_rows)}  title:{title}")
    return {"form_id": form_id, "sheet_url": f"http://localhost:8080/download/email?form_id={form_id}",
            "email_count": len(email_rows), "phone_count": len(phone_rows), "elapsed_seconds": elapsed}

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

from fastapi.staticfiles import StaticFiles
if os.path.exists("static"):
    app.mount("/", StaticFiles(directory="static", html=True), name="static")

if __name__ == "__main__":
    import uvicorn
    print("\n🚀 本機測試 Server 啟動中...")
    print(f"   API endpoint：http://localhost:8080")
    print(f"   TYPEFORM_TOKEN：{'已設定 ✅' if TYPEFORM_TOKEN else '未設定 ❌'}")
    uvicorn.run(app, host="0.0.0.0", port=8080)
