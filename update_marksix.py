import requests
import pandas as pd
from bs4 import BeautifulSoup

SOURCE_URL = "https://lottery.hk/mark-six/results"
EXCEL_FILE = "marksix_latest_200.xlsx"
HTML_FILE = "index.html"

def fetch_latest_200():
    response = requests.get(SOURCE_URL, timeout=20)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")
    table = soup.select_one("table")

    rows = []
    for tr in table.select("tr")[1:]:
        tds = [td.get_text(strip=True) for td in tr.select("td")]
        if len(tds) >= 9:
            rows.append({
                "期數": tds[0],
                "日期": tds[1],
                "N1": int(tds[2]),
                "N2": int(tds[3]),
                "N3": int(tds[4]),
                "N4": int(tds[5]),
                "N5": int(tds[6]),
                "N6": int(tds[7]),
                "特別號": int(tds[8]),
            })

    return pd.DataFrame(rows).head(200)

def update_excel(df):
    df.to_excel(EXCEL_FILE, index=False)

def update_html(latest_issue, latest_date):
    with open(HTML_FILE, "r", encoding="utf-8") as f:
        html = f.read()

    marker = "數據更新至："
    start = html.find(marker)
    if start != -1:
        end = html.find("</p>", start)
        new_line = f"數據更新至：{latest_issue}（{latest_date}）"
        html = html[:start] + marker + new_line + html[end:]

    with open(HTML_FILE, "w", encoding="utf-8") as f:
        f.write(html)

def main():
    df = fetch_latest_200()
    update_excel(df)

    latest = df.iloc[0]
    update_html(latest["期數"], latest["日期"])

    print("✅ Mark Six data updated successfully")

if __name__ == "__main__":
    main()
