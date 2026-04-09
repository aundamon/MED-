"""
export_to_gdrive.py
-------------------
ดึงข้อมูล Working_Hour_Detail (Machine and Equipment Development) จาก Power BI Desktop
แล้ว upload ขึ้น Google Drive เป็น powerbi_export_latest.csv (overwrite ทุกวัน)

วิธีรัน:
    python export_to_gdrive.py              # ดึง 7 วันย้อนหลัง (default)
    python export_to_gdrive.py --all        # ดึงทั้งหมด ไม่ filter วันที่
    python export_to_gdrive.py --days 30    # ดึง 30 วันย้อนหลัง
"""

import argparse
import io
import os
import pickle
import subprocess
import sys
import tempfile
from datetime import datetime, timedelta

import pandas as pd
import psutil
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ── CONFIG (แก้ตรงนี้เท่านั้น) ──────────────────────────────────────────────
COST_CENTERS     = ["0740-15400", "0740-15410", "0740-19110"]
GDRIVE_FOLDER_ID = "1wprFYEbndjL7B5t4h3N8Zc5YytAqrq5G"
EXPORT_FILENAME  = "powerbi_export_latest.csv"
SCOPES           = ["https://www.googleapis.com/auth/drive"]
TOKEN_FILE       = os.path.join(os.path.dirname(__file__), "token.pickle")
CREDENTIALS_FILE = os.path.join(os.path.dirname(__file__), "credentials.json")
# ────────────────────────────────────────────────────────────────────────────

def find_pbi_port() -> int | None:
    """หา port ของ Power BI Desktop (msmdsrv.exe) แบบอัตโนมัติ"""
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            if "msmdsrv" in (proc.info["name"] or "").lower():
                for conn in proc.connections(kind="inet"):
                    if conn.status == "LISTEN":
                        return conn.laddr.port
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return None


def query_pbi(port: int, days_back: int | None) -> pd.DataFrame:
    """Query Power BI Desktop ผ่าน PowerShell + System.Data.OleDb (MSOLAP provider)"""
    today = datetime.now()
    if days_back is not None:
        cutoff = today - timedelta(days=days_back)
        date_filter = (
            f" && 'Working_Hour_Detail'[DATE_ENTRY] >= DATE({cutoff.year}, {cutoff.month}, {cutoff.day})"
            f" && 'Working_Hour_Detail'[DATE_ENTRY] <= DATE({today.year}, {today.month}, {today.day})"
        )
    else:
        date_filter = (
            f" && 'Working_Hour_Detail'[DATE_ENTRY] <= DATE({today.year}, {today.month}, {today.day})"
        )

    cc_list = ", ".join(f'"{cc}"' for cc in COST_CENTERS)
    dax = (
        f"EVALUATE "
        f"FILTER("
        f"'Working_Hour_Detail',"
        f"'Working_Hour_Detail'[INTERNAL_COSTCENTER] IN {{{cc_list}}}"
        f"{date_filter}"
        f") "
        f"ORDER BY 'Working_Hour_Detail'[DATE_ENTRY] DESC"
    )

    # เขียน CSV ลง temp file แล้วให้ Python อ่าน (หลีกเลี่ยงปัญหา encoding ใน stdout)
    tmp_csv = tempfile.NamedTemporaryFile(suffix=".csv", delete=False, mode="w")
    tmp_csv.close()
    tmp_path = tmp_csv.name

    ps_script = f"""
$ErrorActionPreference = 'Stop'
try {{
    $conn = New-Object System.Data.OleDb.OleDbConnection
    $conn.ConnectionString = "Provider=MSOLAP;Data Source=localhost:{port};"
    $conn.Open()

    $cmd = $conn.CreateCommand()
    $cmd.CommandText = '{dax.replace("'", "''")}'
    $cmd.CommandTimeout = 180

    $adapter = New-Object System.Data.OleDb.OleDbDataAdapter $cmd
    $table   = New-Object System.Data.DataTable
    $null    = $adapter.Fill($table)
    $conn.Close()

    $table | Export-Csv -Path '{tmp_path.replace(chr(92), "/")}' -NoTypeInformation -Encoding UTF8
    Write-Host "OK:$($table.Rows.Count)"
}} catch {{
    Write-Error $_.Exception.Message
    exit 1
}}
"""

    print(f"  Querying via PowerShell OleDb (MSOLAP) on port {port}...")
    result = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
        capture_output=True, text=True, timeout=200
    )

    if result.returncode != 0:
        err = result.stderr.strip() or result.stdout.strip()
        print(f"ERROR PowerShell:\n{err}")
        sys.exit(1)

    print(f"  {result.stdout.strip()}")

    df = pd.read_csv(tmp_path, encoding="utf-8-sig")
    os.unlink(tmp_path)

    # ตัด table prefix เช่น "Working_Hour_Detail[DATE_ENTRY]" → "DATE_ENTRY"
    df.columns = [c.split("[")[-1].rstrip("]") if "[" in c else c for c in df.columns]
    return df


def get_gdrive_service():
    """Authenticate Google Drive (เปิด browser ครั้งแรก หลังจากนั้น auto-refresh)"""
    if not os.path.exists(CREDENTIALS_FILE):
        print(f"ERROR: ไม่พบไฟล์ {CREDENTIALS_FILE}")
        print("  → ดาวน์โหลดจาก Google Cloud Console แล้ววางไว้ที่เดียวกับ script")
        sys.exit(1)

    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "rb") as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "wb") as f:
            pickle.dump(creds, f)

    return build("drive", "v3", credentials=creds)


def upload_to_drive(service, df: pd.DataFrame) -> None:
    """Upload หรือ overwrite ไฟล์ CSV ใน Google Drive folder ที่กำหนด"""
    csv_bytes = df.to_csv(index=False).encode("utf-8-sig")  # utf-8-sig รองรับ Excel ไทย
    media = MediaIoBaseUpload(io.BytesIO(csv_bytes), mimetype="text/csv", resumable=False)

    results = service.files().list(
        q=(f"name='{EXPORT_FILENAME}' "
           f"and '{GDRIVE_FOLDER_ID}' in parents "
           f"and trashed=false"),
        fields="files(id, name)"
    ).execute()
    existing = results.get("files", [])

    if existing:
        file_id = existing[0]["id"]
        service.files().update(fileId=file_id, media_body=media).execute()
        print(f"  Overwritten existing file (ID: {file_id})")
    else:
        meta = {"name": EXPORT_FILENAME, "parents": [GDRIVE_FOLDER_ID]}
        f = service.files().create(body=meta, media_body=media, fields="id").execute()
        print(f"  Uploaded new file (ID: {f['id']})")


def main():
    parser = argparse.ArgumentParser(description="Export Power BI → Google Drive")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--days", type=int, default=7,
                       help="จำนวนวันย้อนหลัง (default: 7)")
    group.add_argument("--all", dest="all_data", action="store_true",
                       help="ดึงข้อมูลทั้งหมด ไม่ filter วันที่")
    args = parser.parse_args()

    days_back = None if args.all_data else args.days
    label     = "ทั้งหมด" if days_back is None else f"{days_back} วันย้อนหลัง"

    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] Export เริ่มต้น — filter: {label}")
    print(f"  Cost centers: {', '.join(COST_CENTERS)}")

    # 1. หา port
    port = find_pbi_port()
    if not port:
        print("ERROR: ไม่พบ Power BI Desktop กรุณาเปิดไฟล์ ManpowerWorkload_Rev ก่อน")
        sys.exit(1)
    print(f"  พบ Power BI Desktop บน port {port}")

    # 2. Query ข้อมูล
    df = query_pbi(port, days_back)
    if df.empty:
        print(f"WARNING: ไม่มีข้อมูลสำหรับ filter ที่กำหนด — ไม่ได้ upload")
        sys.exit(0)
    print(f"  ได้ข้อมูล {len(df):,} rows, {len(df.columns)} columns")

    # 3. Upload
    print("  Uploading to Google Drive...")
    service = get_gdrive_service()
    upload_to_drive(service, df)

    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] เสร็จแล้ว ✓")


if __name__ == "__main__":
    main()
