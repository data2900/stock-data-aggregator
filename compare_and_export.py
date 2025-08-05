import sqlite3
import pandas as pd
import argparse
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

import os
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "market_data.db")
CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.json")
SPREADSHEET_NAME = "日本株"
SHEET_NAME = "StockData_統合"

# 英語列名 → 日本語列名 マッピング
COLUMN_NAME_MAP = {
    "price": "株価",
    "per": "予想PER",
    "yield_rate": "予想配当利回り",
    "pbr": "PBR（実績）",
    "roe_n": "ROE（予想）",
    "earning_yield": "株式益回り（予想）",
    "sales_growth": "増収率",
    "op_profit_growth": "経常増益率",
    "op_margin": "売上高経常利益率",
    "roe_s": "ROE",
    "roa": "ROA",
    "equity_ratio": "株主資本比率",
    "dividend_payout": "配当性向",
    "rating": "レーティング",
    "sales": "売上高予想",
    "profit": "経常利益予想",
    "scale": "規模",
    "cheap": "割安度",
    "growth": "成長",
    "profitab": "収益性",
    "safety": "安全性",
    "risk": "リスク",
    "return_rate": "リターン",
    "liquidity": "流動性",
    "trend": "トレンド",
    "forex": "為替",
    "technical": "テクニカル",
}

def authorize_gspread():
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scope)
    gc = gspread.authorize(creds)
    return gc.open(SPREADSHEET_NAME)

def parse_date_range(date_str):
    if '-' in date_str:
        start_str, end_str = date_str.split('-')
        if len(end_str) == 4:
            end_str = start_str[:4] + end_str
        start = datetime.strptime(start_str, "%Y%m%d").date()
        end = datetime.strptime(end_str, "%Y%m%d").date()
        return start, end
    else:
        date = datetime.strptime(date_str, "%Y%m%d").date()
        return date, date

def get_available_dates_in_range(start_date, end_date):
    conn = sqlite3.connect(DB_PATH)
    query = '''
        SELECT DISTINCT target_date FROM consensus_url
        WHERE target_date BETWEEN ? AND ?
    '''
    rows = conn.execute(query, (start_date.strftime("%Y%m%d"), end_date.strftime("%Y%m%d"))).fetchall()
    conn.close()
    return sorted([row[0] for row in rows])

def fetch_merged_data(dates):
    conn = sqlite3.connect(DB_PATH)
    placeholders = ",".join(["?"] * len(dates))
    query = f'''
        SELECT
            c.target_date, c.code, c.name,
            n.sector, n.price, n.per, n.yield_rate, n.pbr, n.roe AS roe_n,
            n.earning_yield,
            s.sales_growth, s.op_profit_growth, s.op_margin, s.roe AS roe_s,
            s.roa, s.equity_ratio, s.dividend_payout,
            m.rating, m.sales, m.profit, m.scale, m.cheap, m.growth, m.profitab,
            m.safety, m.risk, m.return_rate, m.liquidity, m.trend, m.forex, m.technical,
            c.nikkeiurl, c.quickurl, c.sbiurl
        FROM consensus_url c
        LEFT JOIN nikkei_reports n ON c.target_date = n.target_date AND c.code = n.code
        LEFT JOIN moneyworld_reports m ON c.target_date = m.target_date AND c.code = m.code
        LEFT JOIN sbi_reports s ON c.target_date = s.target_date AND c.code = s.code
        WHERE c.target_date IN ({placeholders})
    '''
    df = pd.read_sql_query(query, conn, params=dates)
    conn.close()
    return df

def reshape_data(df):
    base_cols = ["target_date", "code", "name", "sector",
                 "price", "per", "yield_rate", "pbr", "roe_n", "earning_yield",
                 "sales_growth", "op_profit_growth", "op_margin", "roe_s", "roa", "equity_ratio", "dividend_payout",
                 "rating", "sales", "profit", "scale", "cheap", "growth", "profitab",
                 "safety", "risk", "return_rate", "liquidity", "trend", "forex", "technical"]

    url_cols = ["nikkeiurl", "quickurl", "sbiurl"]

    url_df = (
        df.sort_values("target_date", ascending=False)
          .drop_duplicates(subset=["code"])
          .set_index("code")[url_cols]
    )

    all_dates = sorted(df["target_date"].unique(), reverse=True)
    records = []

    for date in all_dates:
        day_df = df[df["target_date"] == date].copy()
        day_df = day_df.drop(columns=url_cols)
        rename_map = {
            col: f"{COLUMN_NAME_MAP.get(col, col)}_{date}" for col in base_cols[4:]
        }
        day_df = day_df.rename(columns=rename_map)
        records.append(day_df)

    merged = records[0][["code", "name", "sector"]].copy()
    for df_ in records:
        merged = pd.merge(merged, df_, on=["code", "name", "sector"], how="outer")

    final = merged.set_index("code")
    final = final.join(url_df, how="left")
    final.reset_index(inplace=True)

    ordered_cols = ["code", "name", "sector"]
    for date in all_dates:
        ordered_cols += [f"{COLUMN_NAME_MAP.get(col, col)}_{date}" for col in base_cols[4:]]
    ordered_cols += url_cols

    final = final[ordered_cols]
    final.insert(0, "日付", ",".join(all_dates))
    final = final.rename(columns={"code": "コード", "name": "企業名", "sector": "セクター"})

    return final

def sanitize(val):
    return "" if pd.isna(val) else str(val)

def export_to_gsheet(df):
    header = list(df.columns)
    rows = [list(map(sanitize, row)) for _, row in df.iterrows()]
    sh = authorize_gspread()
    try:
        worksheet = sh.worksheet(SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = sh.add_worksheet(title=SHEET_NAME, rows="1000", cols="100")
    worksheet.clear()
    worksheet.update([header] + rows)
    print(f"✅ {len(df)}件のデータをGoogle Sheetsに出力しました。")

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", required=True, help="例: 20250728 または 20250701-0731")
    args = parser.parse_args()

    start_date, end_date = parse_date_range(args.date)
    available_dates = get_available_dates_in_range(start_date, end_date)
    if not available_dates:
        print("⚠️ 指定された期間に利用可能なデータが見つかりませんでした。")
        return

    df = fetch_merged_data(available_dates)
    if df.empty:
        print("⚠️ データが見つかりませんでした。")
        return

    reshaped_df = reshape_data(df)
    export_to_gsheet(reshaped_df)

if __name__ == "__main__":
    main()
