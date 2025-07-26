import pandas as pd
from pathlib import Path
import gspread
from google.oauth2.service_account import Credentials
import math

# === パス設定 ===
BASE_DIR = Path("/aggregation")
PKL_FILE = BASE_DIR / "data_store.pkl"
CREDS_FILE = BASE_DIR / "credentials.json"
SPREADSHEET_NAME = "日本株"
SHEET_NAME = "StockData"

# === 出力対象の項目（新項目を追加） ===
TARGET_COLUMNS = [
    "株価", "予想PER", "予想配当利回り", "PBR（実績）", "ROE（予想）", "株式益回り（予想）",
    "増収率", "経常増益率", "売上高経常利益率", "ROE", "ROA", "株主資本比率", "配当性向",
    "レーティング", "売上高予想", "経常利益予想", "規模", "割安度", "成長", "収益性",
    "安全性", "リスク", "リターン", "流動性", "トレンド", "為替", "テクニカル"
]

# === データ読み込み ===
df = pd.read_pickle(PKL_FILE)
df["証券コード"] = df["証券コード"].astype(str).str.strip()

# 最新4週を抽出（新しい順）
latest_dates = sorted(df["日付"].unique())[-4:][::-1]
df = df[df["日付"].isin(latest_dates)]

# === セクター・企業名を最新週から取得
latest_info = (
    df[df["日付"] == latest_dates[0]][["証券コード", "セクター", "企業名"]]
    .drop_duplicates(subset="証券コード")
    .set_index("証券コード")
)

# === ピボット（日付×項目）を横展開
pivoted = df.pivot(index="証券コード", columns="日付", values=TARGET_COLUMNS)
pivoted = pivoted.reorder_levels([1, 0], axis=1).sort_index(axis=1, ascending=False)

# === 整形用DataFrame作成
formatted = pd.DataFrame(index=pivoted.index)
for date in latest_dates:
    for col in TARGET_COLUMNS:
        colname = f"{col}_{date}"
        formatted[colname] = pivoted[(date, col)]

# === セクター・企業名を追加
final_df = pd.concat([latest_info, formatted], axis=1).reset_index()

# === Google Sheets 出力設定 ===
scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
creds = Credentials.from_service_account_file(str(CREDS_FILE), scopes=scope)
gc = gspread.authorize(creds)
sh = gc.open(SPREADSHEET_NAME)

try:
    worksheet = sh.worksheet(SHEET_NAME)
except gspread.exceptions.WorksheetNotFound:
    worksheet = sh.add_worksheet(title=SHEET_NAME, rows="1000", cols="100")

worksheet.clear()

# === NaN・Infinity 処理付きサニタイズ関数（%処理は削除） ===
def sanitize_value(val, colname=""):
    if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
        return ""
    return str(val)

# === 書き出し用データ整形
header = ["証券コード", "セクター", "企業名"] + list(formatted.columns)
rows = []

for _, row in final_df.iterrows():
    row_data = [
        sanitize_value(row["証券コード"]),
        sanitize_value(row.get("セクター", "")),
        sanitize_value(row.get("企業名", ""))
    ] + [
        sanitize_value(val, colname)
        for val, colname in zip(row.iloc[3:], formatted.columns)
    ]
    rows.append(row_data)

worksheet.update([header] + rows)

print("✅ 最新4週データをスプレッドシートに出力しました。")
