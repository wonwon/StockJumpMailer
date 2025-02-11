import os

import pandas as pd
import requests
from bs4 import BeautifulSoup

# ベースURL
base_url = "https://kabureal.net/raterank/span/?d=2&page="

# ヘッダー情報（アクセスブロックを防ぐため）
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
}

# データ格納用リスト
data = []

# データ保存ディレクトリを作成
output_dir = "data"
os.makedirs(output_dir, exist_ok=True)

# ページ番号
page = 1
while True:
    url = base_url + str(page)
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, "html.parser")
    
    # ランキングのテーブルを探す
    table = soup.find("table")
    if not table:
        break  # テーブルがない場合、終了

    rows = table.find_all("tr")[1:]  # ヘッダー行をスキップ

    for row in rows:
        cols = row.find_all("td")
        if len(cols) >= 4:
            stock_code = cols[0].text.strip()  # 株価コード
            company_name = cols[1].text.strip()  # 会社名
            price_change = cols[3].text.strip()  # 値上がり幅（金額）
            
            data.append([stock_code, company_name, price_change])
    
    print(f"ページ {page} のデータを取得完了")
    page += 1

# pandasデータフレームに変換
df = pd.DataFrame(data, columns=["株価コード", "会社名", "値上がり幅（円）"])

# エクセルに保存
excel_filename = os.path.join(output_dir, "kabureal_ranking.xlsx")
df.to_excel(excel_filename, index=False)

print(f"データをエクセルファイル {excel_filename} に保存しました。")
