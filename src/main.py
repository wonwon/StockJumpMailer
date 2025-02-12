import os
import re

import pandas as pd
import requests
from bs4 import BeautifulSoup

# ベースURL
base_url = "https://kabureal.net/raterank/span/?page={}&d=1"

# ヘッダー情報（アクセスブロックを防ぐため）
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
}

# データ格納用リスト
data = []

# データ保存ディレクトリを作成
output_dir = "data"
os.makedirs(output_dir, exist_ok=True)

# 最初のページを取得して最後のページ番号を取得
response = requests.get(base_url.format(1), headers=headers)
soup = BeautifulSoup(response.text, "html.parser")

# ページネーションリンクを取得
pagination_links = soup.find_all("a", href=True)
max_page = 1
for link in pagination_links:
    if "page=" in link["href"]:
        try:
            page_num = int(link["href"].split("page=")[1].split("&")[0])
            max_page = max(max_page, page_num)
        except ValueError:
            continue

# ページごとにクロール
for page in range(1, max_page + 1):
    url = base_url.format(page)
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, "html.parser")
    
    # データ取得
    stock_sections = soup.find_all("div", class_="ovflow")
    
    for section in stock_sections:
        stock_info = section.find("div", class_="fs11")
        company_info = section.find("a", class_="fs15 fbold")
        price_info_container = section.find("span", class_="fcolor_12 mleft10")
        
        if stock_info and company_info and price_info_container:
            stock_code = stock_info.text.split("・")[0].strip()
            company_name = company_info.text.strip()
            
            # 直近3日間の株価値上がり率の文言から数値を抽出
            price_text = price_info_container.get_text(strip=True)

            # `（+1,234）` のような4桁カンマ入りの数値も考慮
            price_match = re.search(r"（\+?([\d,]+)）", price_text)
            
            # エラー回避: price_match が None の場合 "N/A" を設定
            if price_match:
                price_change = price_match.group(1).replace(",", "")  # カンマを削除して数値化
            else:
                # `+1234` の形式がある可能性も考慮
                price_match_alt = re.search(r"\+?([\d,]+)", price_text)
                price_change = price_match_alt.group(1).replace(",", "") if price_match_alt else "N/A"

            data.append([stock_code, company_name, price_change])
    
    print(f"ページ {page} のデータを取得完了")

# pandasデータフレームに変換
df = pd.DataFrame(data, columns=["株価コード", "会社名", "値上がり幅（円）"])

# エクセルに保存
excel_filename = os.path.join(output_dir, "kabureal_ranking.xlsx")
df.to_excel(excel_filename, index=False)

print(f"データをエクセルファイル {excel_filename} に保存しました。")
