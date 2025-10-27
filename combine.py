import pyodbc
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# =========================
# Config SQL Server
# =========================
SQL_SERVER = "10.9.227.118"
DATABASE   = "INT_MART"
USER       = "10000930"
PASSWORD   = "P@ssw0rd"
SQL_QUERY  = "SELECT TOP 1000 * FROM [INT_MART].[013_mart].[HR_Accum];"

# =========================
# Config CSV & Google Sheets
# =========================
CSV_PATH = r"D:\Users\10000930\Desktop\scripts\HR_Accum.csv"
SHEET_URL = "https://docs.google.com/spreadsheets/d/1kgT5e-h3ohTssdzE-CSLkaV44QrqQMjTJ2Qa8hHHgQY/edit#gid=0"
SERVICE_ACCOUNT_FILE = "service_account.json"

# =========================
# Step 1: Connect to SQL Server and fetch data
# =========================
conn_str = (
    f"DRIVER={{ODBC Driver 17 for SQL Server}};"
    f"SERVER={SQL_SERVER};"
    f"DATABASE={DATABASE};"
    f"UID={USER};"
    f"PWD={PASSWORD}"
)

try:
    conn = pyodbc.connect(conn_str)
    df = pd.read_sql(SQL_QUERY, conn)
    conn.close()
    print(f"Fetched {len(df)} rows from SQL Server.")
except Exception as e:
    print("SQL ERROR:", e)
    exit(1)

# =========================
# Step 2: Save CSV locally
# =========================
# ป้องกันปัญหาเลข scientific notation
for col in df.select_dtypes(include=['float', 'int']):
    df[col] = df[col].apply(lambda x: f"{x:.0f}" if pd.notnull(x) else "")

df.to_csv(CSV_PATH, index=False, encoding="utf-8-sig")
print(f"Saved CSV to {CSV_PATH}")

# =========================
# Step 3: Upload CSV to Google Sheets
# =========================
# Load CSV
df = pd.read_csv(CSV_PATH, dtype=str)  # อ่านทุกคอลัมน์เป็น string
# df = df.replace([float('inf'), float('-inf')], "")
df = df.fillna("")

# Connect to Google Sheets
scopes = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]
credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
gc = gspread.authorize(credentials)

# Open the sheet
sh = gc.open_by_url(SHEET_URL)
worksheet = sh.sheet1

# Clear existing content
worksheet.clear()

# Upload CSV content
worksheet.update([df.columns.values.tolist()] + df.values.tolist())

print("Upload to Google Sheets complete!")
