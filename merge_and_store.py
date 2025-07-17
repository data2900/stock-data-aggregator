import pandas as pd
import pickle
from pathlib import Path
import re
import os

base_dir = "/StockData"
csv_dir = os.path.join(base_dir, "data/csv")
os.makedirs(csv_dir, exist_ok=True)

# === パス設定 ===
# スクリプトのある場所を基準とする
BASE_DIR = Path(__file__).resolve().parent
CSV_DIR = BASE_DIR / "data" / "csv"
OUTPUT_DIR = BASE_DIR
PKL_FILE = OUTPUT_DIR / "data_store.pkl"
CSV_FILE = OUTPUT_DIR / "data_store.csv"

# === 既存データの読み込み ===
if PKL_FILE.exists():
    with open(PKL_FILE, "rb") as f:
        all_data = pickle.load(f)
else:
    all_data = pd.DataFrame()

# === ファイル一覧の取得 ===
nikkei_files = sorted(CSV_DIR.glob("nikkeireport_*.csv"))
money_files = sorted(CSV_DIR.glob("moneyworldreport_*.csv"))

# === 日付抽出関数 ===
def extract_date(filename):
    match = re.search(r"_(\d{4}-\d{2}-\d{2})_", filename)
    return match.group(1) if match else None

# === 日付ごとのマージ処理 ===
for nikkei_file in nikkei_files:
    date_str = extract_date(nikkei_file.name)
    if not date_str:
        continue

    # 対応する money ファイルを探す
    money_file = next((f for f in money_files if date_str in f.name), None)

    # CSV読み込み
    df_nikkei = pd.read_csv(nikkei_file).copy()
    df_nikkei["証券コード"] = df_nikkei["証券コード"].astype(str).str.replace(":", "").str.strip()

    if money_file:
        df_money = pd.read_csv(money_file).copy()
        df_money["証券コード"] = df_money["証券コード"].astype(str).str.replace(":", "").str.strip()
        df_merged = pd.merge(df_nikkei, df_money, on="証券コード", how="outer", suffixes=("_nikkei", "_money"))
    else:
        df_merged = df_nikkei.copy()

    df_merged["日付"] = date_str
    all_data = pd.concat([all_data, df_merged], ignore_index=True)

# === 重複排除（同一日付・証券コード）
all_data.drop_duplicates(subset=["日付", "証券コード"], keep="last", inplace=True)

# === 保存
all_data.to_csv(CSV_FILE, index=False)
with open(PKL_FILE, "wb") as f:
    pickle.dump(all_data, f)

print(f"✅ データ保存完了: {len(all_data)} 件")