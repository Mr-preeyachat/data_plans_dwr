import os
import sys
import json
import time
import datetime
import threading
import urllib.request
import csv
import io
import webbrowser
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import parse_qs, urlparse

# Force UTF-8 stdout/stderr on Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

PORT = 8000
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
DATA_JSON_PATH = os.path.join(ASSETS_DIR, "data.json")

# Google Sheets Configuration
GOOGLE_SHEET_ID = "18bZNi3IlQNiUHtVhUvsZV65zA2EpqJRyl1ARgDRKsN0"
GOOGLE_SHEET_NAME = "DWR_007"
GOOGLE_SHEET_CSV_URL = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/gviz/tq?tqx=out:csv&sheet={GOOGLE_SHEET_NAME}"
ENABLE_GOOGLE_SHEETS = True

# Column Map: Excel Column (1-indexed) -> Field Key
COL_MAP = {
    4: "D",   # หน่วยงาน
    5: "E",   # ชื่อโครงการ
    6: "F",   # หมู่บ้าน
    7: "G",   # หมู่ที่
    8: "H",   # ตำบล
    9: "I",   # อำเภอ
    10: "J",  # จังหวัด
    11: "K",  # ภาค
    12: "L",  # ลุ่มน้ำหลัก
    13: "M",  # ลุ่มน้ำสาขา
    14: "lat",# ละติจูด
    15: "lon",# ลองจิจูด
    16: "P",  # งบประมาณ
    17: "Q",  # ประเภทงบประมาณ
    24: "X",  # สนับสนุน
    30: "AD", # กิจกรรม-ลักษณะงาน
    31: "AE", # ประเภทโครงการ
    33: "AG", # ปริมาณน้ำเพิ่มขึ้น
    34: "AH", # ครัวเรือนได้รับประโยชน์
    35: "AI", # พื้นที่การเกษตรได้รับประโยชน์
    36: "AJ", # ปริมาณน้ำที่กระจายได้
    37: "AK"  # ระยะทาง
}

NUMERIC_FIELDS = {"P", "AG", "AH", "AI", "AJ", "AK", "lat", "lon"}

cached_data = []
last_mtime = 0
last_filename = ""
data_source = ""
last_fetch_time = 0
CACHE_TTL_SECONDS = 30  # Re-fetch Google Sheets at most every 30 seconds

def clean_num(val):
    if val is None or val == "":
        return 0
    if isinstance(val, (int, float)):
        return float(val) if not str(val).endswith(".0") else float(val)
    s = str(val).replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0

def fetch_google_sheet():
    print(f"[INFO] Fetching live data from Google Sheets ({GOOGLE_SHEET_NAME}) ...")
    start_time = time.time()
    try:
        req = urllib.request.Request(
            GOOGLE_SHEET_CSV_URL,
            headers={"User-Agent": "Mozilla/5.0"}
        )
        with urllib.request.urlopen(req, timeout=10) as res:
            raw_bytes = res.read()
            content = raw_bytes.decode("utf-8", errors="replace")

        reader = csv.reader(io.StringIO(content))
        rows_data = []
        row_idx = 0
        for row in reader:
            row_idx += 1
            if row_idx < 6:  # Start row 6
                continue
            if not row or len(row) < 5 or not str(row[4]).strip():
                continue

            item = {}
            for col_1idx, key in COL_MAP.items():
                val_idx = col_1idx - 1
                val = row[val_idx] if val_idx < len(row) else ""
                if key in NUMERIC_FIELDS:
                    item[key] = clean_num(val)
                else:
                    s = str(val or "").replace("\n", " ").strip()
                    while "  " in s:
                        s = s.replace("  ", " ")
                    item[key] = s
            rows_data.append(item)

        print(f"[INFO] Successfully parsed {len(rows_data)} rows from Google Sheets in {time.time()-start_time:.2f}s")
        return rows_data
    except Exception as e:
        print(f"[WARNING] Failed to fetch Google Sheet: {e}")
        return None

def find_excel_file():
    if not os.path.exists(ASSETS_DIR):
        return None
    candidates = [
        "แบบฟอร์ม_DWR007_260820.xlsm",
        "DWR007.xlsx"
    ]
    for filename in candidates:
        filepath = os.path.join(ASSETS_DIR, filename)
        if os.path.exists(filepath):
            return filepath
    for filename in os.listdir(ASSETS_DIR):
        if (filename.endswith(".xlsm") or filename.endswith(".xlsx")) and not filename.startswith("~$"):
            return os.path.join(ASSETS_DIR, filename)
    return None

def parse_excel():
    global cached_data, last_mtime, last_filename, data_source
    excel_path = find_excel_file()
    if not excel_path:
        print("[WARNING] No Excel file found in assets/")
        return cached_data

    mtime = os.path.getmtime(excel_path)
    filename = os.path.basename(excel_path)

    print(f"[INFO] Reading local Excel file: {filename} ...")
    start_time = time.time()
    
    rows_data = []
    try:
        import openpyxl
        wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
        sheet_name = "DWR_007" if "DWR_007" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]
        
        row_idx = 0
        for row in ws.iter_rows(values_only=True):
            row_idx += 1
            if row_idx < 6:
                continue
            if not row or len(row) < 5 or not row[4]:
                continue
            
            pname = str(row[4]).strip()
            if not pname:
                continue

            item = {}
            for col_1idx, key in COL_MAP.items():
                val_idx = col_1idx - 1
                val = row[val_idx] if val_idx < len(row) else ""
                if key in NUMERIC_FIELDS:
                    item[key] = clean_num(val)
                else:
                    s = str(val or "").replace("\n", " ").strip()
                    while "  " in s:
                        s = s.replace("  ", " ")
                    item[key] = s
            
            rows_data.append(item)

        wb.close()
    except Exception as e:
        print(f"[ERROR] Failed reading Excel with openpyxl: {e}")
        try:
            import pandas as pd
            df = pd.read_excel(excel_path, sheet_name="DWR_007" if "DWR_007" else 0, header=None)
            df = df.iloc[5:]
            for _, r in df.iterrows():
                pname = str(r.iloc[4] if len(r) > 4 else "").strip()
                if not pname or pname.lower() == "nan":
                    continue
                item = {}
                for col_1idx, key in COL_MAP.items():
                    val_idx = col_1idx - 1
                    val = r.iloc[val_idx] if val_idx < len(r) else ""
                    if key in NUMERIC_FIELDS:
                        item[key] = clean_num(val)
                    else:
                        item[key] = "" if str(val).lower() == "nan" else str(val or "").strip()
                rows_data.append(item)
        except Exception as e2:
            print(f"[ERROR] Fallback reading Excel failed: {e2}")

    last_mtime = mtime
    last_filename = filename
    data_source = "Local Excel File"

    save_data_json(rows_data, filename, data_source, mtime)
    return rows_data

def save_data_json(rows_data, source_name, source_type, mtime=None):
    try:
        mtime_str = datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S") if mtime else datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        payload = {
            "status": "success",
            "source": source_type,
            "file": source_name,
            "updated_at": mtime_str,
            "total": len(rows_data),
            "data": rows_data
        }
        with open(DATA_JSON_PATH, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[WARNING] Could not save data.json: {e}")

def get_data(force_refresh=False):
    global cached_data, last_mtime, last_filename, data_source, last_fetch_time
    now = time.time()

    # Cache TTL check
    if not force_refresh and cached_data and (now - last_fetch_time < CACHE_TTL_SECONDS):
        return cached_data

    # Attempt 1: Fetch Google Sheets
    if ENABLE_GOOGLE_SHEETS:
        gs_data = fetch_google_sheet()
        if gs_data is not None and len(gs_data) > 0:
            cached_data = gs_data
            data_source = "Google Sheets (Live Sync)"
            last_filename = f"Google Sheet ({GOOGLE_SHEET_NAME})"
            last_mtime = now
            last_fetch_time = now
            save_data_json(gs_data, last_filename, data_source)
            return cached_data

    # Attempt 2: Fallback to local Excel file
    excel_data = parse_excel()
    if excel_data:
        cached_data = excel_data
        last_fetch_time = now
        return cached_data

    return cached_data

class DashboardRequestHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/data":
            query = parse_qs(parsed.query)
            force = "force" in query or "refresh" in query
            data = get_data(force_refresh=force)
            mtime_str = datetime.datetime.fromtimestamp(last_mtime).strftime("%Y-%m-%d %H:%M:%S") if last_mtime else ""
            payload = {
                "status": "success",
                "source": data_source,
                "file": last_filename,
                "updated_at": mtime_str,
                "total": len(data),
                "data": data
            }
            body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        if parsed.path == "/api/status":
            mtime_str = datetime.datetime.fromtimestamp(last_mtime).strftime("%Y-%m-%d %H:%M:%S") if last_mtime else ""
            payload = {
                "status": "ok",
                "source": data_source,
                "file": last_filename,
                "last_modified": mtime_str,
                "total": len(cached_data)
            }
            body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        super().do_GET()

def open_browser():
    time.sleep(1.2)
    print(f"[INFO] Opening browser at http://localhost:{PORT}")
    webbrowser.open(f"http://localhost:{PORT}")

def main():
    if "--test" in sys.argv:
        print("[TEST] Running Data parser test...")
        data = get_data(force_refresh=True)
        print(f"[TEST] Source: {data_source}")
        print(f"[TEST] Total rows parsed: {len(data)}")
        if data:
            print("[TEST] Row 1 Sample:", json.dumps(data[0], ensure_ascii=False, indent=2))
        return

    print("=" * 60)
    print(" DWR007 Project Dashboard Server")
    print(" Google Sheets Sync & Local Excel Fallback")
    print("=" * 60)
    
    # Pre-parse data on startup
    get_data(force_refresh=True)

    server = HTTPServer(("0.0.0.0", PORT), DashboardRequestHandler)
    print(f"[SUCCESS] Server running at http://localhost:{PORT}")
    print("[INFO] Press Ctrl+C to stop server")

    # Auto open browser in thread
    threading.Thread(target=open_browser, daemon=True).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n[INFO] Server stopped.")

if __name__ == "__main__":
    main()
