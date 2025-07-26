import pandas as pd
import pickle
from pathlib import Path
import re
import os

# === ベースディレクトリの設定 ===
base_dir = "/aggregation"
csv_dir = os.path.join(base_dir, "data/csv")
os.makedirs(csv_dir, exist_ok=True)

# === 各種パス ===
CSV_DIR = Path(csv_dir)
OUTPUT_DIR = Path(base_dir)
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
sbi_files = sorted(CSV_DIR.glob("sbifinancedata_*.csv"))

# === 日付抽出関数 ===
def extract_date(filename):
    match = re.search(r"_(\d{4}-\d{2}-\d{2})_", filename)
    return match.group(1) if match else None

# === %をつける列（SBI）
sbi_percent_columns = [
    "増収率", "経常増益率", "売上高経常利益率", "ROE", "ROA", "株主資本比率", "配当性向"
]

# === %をつける列（NIKKEI）
nikkei_percent_columns = [
    "予想配当利回り", "ROE（予想）", "株式益回り（予想）"
]

# === 日付ごとのマージ処理 ===
for nikkei_file in nikkei_files:
    date_str = extract_date(nikkei_file.name)
    if not date_str:
        continue

    # money・sbi 対応ファイルの取得
    money_file = next((f for f in money_files if date_str in f.name), None)
    sbi_file = next((f for f in sbi_files if date_str in f.name), None)

    # === nikkei 読み込み ===
    df_nikkei = pd.read_csv(nikkei_file).copy()
    df_nikkei["証券コード"] = df_nikkei["証券コード"].astype(str).str.replace(":", "").str.strip()

    for col in nikkei_percent_columns:
        if col in df_nikkei.columns:
            df_nikkei[col] = df_nikkei[col].apply(
                lambda x: f"{x}%" if pd.notnull(x) and not str(x).strip().endswith('%') else x
            )

    # === sbi 読み込みとマージ ===
    if sbi_file:
        df_sbi = pd.read_csv(sbi_file).copy()
        df_sbi["証券コード"] = df_sbi["証券コード"].astype(str).str.replace(":", "").str.strip()

        for col in sbi_percent_columns:
            if col in df_sbi.columns:
                df_sbi[col] = df_sbi[col].apply(
                    lambda x: f"{x}%" if pd.notnull(x) and not str(x).strip().endswith('%') else x
                )

        df_merged = pd.merge(df_nikkei, df_sbi, on="証券コード", how="outer", suffixes=("_nikkei", "_sbi"))
    else:
        df_merged = df_nikkei.copy()

    # === money 読み込みとマージ（URL列を除外） ===
    if money_file:
        df_money = pd.read_csv(money_file).copy()
        df_money["証券コード"] = df_money["証券コード"].astype(str).str.replace(":", "").str.strip()

        if "URL" in df_money.columns:
            df_money = df_money.drop(columns=["URL"])

        df_merged = pd.merge(df_merged, df_money, on="証券コード", how="outer", suffixes=("", "_money"))

    # === 日付追加
    df_merged["日付"] = date_str

    # === 統合
    all_data = pd.concat([all_data, df_merged], ignore_index=True)

# === 重複排除（同一日付・証券コード）
all_data.drop_duplicates(subset=["日付", "証券コード"], keep="last", inplace=True)

# === 保存
all_data.to_csv(CSV_FILE, index=False)
with open(PKL_FILE, "wb") as f:
    pickle.dump(all_data, f)

print(f"✅ データ保存完了: {len(all_data)} 件")
