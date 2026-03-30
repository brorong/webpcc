import os
import time
import urllib.parse
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pandas as pd

def get_dynamic_url():
    # GitHub Actions 伺服器預設為 UTC 時間，校正為台灣時間 (UTC+8)
    tw_time = datetime.utcnow() + timedelta(hours=8)
    
    # 預設抓取最近 3 天的資料
    start_date = tw_time - timedelta(days=15)
    
    start_str = start_date.strftime("%Y/%m/%d")
    end_str = tw_time.strftime("%Y/%m/%d")
    
    # 將日期轉為 URL 編碼 (例如 / 變成 %2F)
    start_encoded = urllib.parse.quote(start_str, safe='')
    end_encoded = urllib.parse.quote(end_str, safe='')
    
    print(f"🔍 搜尋區間：{start_str} ~ {end_str}")
    
    url = f"https://web.pcc.gov.tw/prkms/tender/common/agent/readTenderAgent?pageSize=50&firstSearch=false&isQuery=&isBinding=N&isLogIn=N&orgName=&orgId=&tenderName=%E9%9B%BB%E5%8B%95%E8%BB%8A&tenderId=&tenderStatus=TENDER_STATUS_1&tenderWay=TENDER_WAY_ALL_DECLARATION&awardAnnounceStartDate={start_encoded}&awardAnnounceEndDate={end_encoded}&radProctrgCate=&tenderRange=TENDER_RANGE_ALL&minBudget=&maxBudget=&item=&gottenVendorName=&gottenVendorId=&submitVendorName=&submitVendorId=&execLocation=&priorityCate=&radReConstruct=&policyAdvocacy=&isCpp="
    return url

def scrape_pcc_tenders():
    url = get_dynamic_url()
    
    # 設定 Selenium 參數 (針對 GitHub Actions 環境優化)
    chrome_options = Options()
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")            
    chrome_options.add_argument("--disable-dev-shm-usage") 
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

    print("正在啟動虛擬瀏覽器...")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    base_url = "https://web.pcc.gov.tw"
    data = []

    try:
        driver.get(url)
        print("等待網頁 JavaScript 載入資料 (5秒)...")
        time.sleep(5) 
        
        soup = BeautifulSoup(driver.page_source, "html.parser")
        rows = soup.find_all("tr")

        for row in rows:
            cols = row.find_all(["td", "th"])
            
            if len(cols) >= 7: 
                # 略過表頭
                if cols[0].text.strip() == '項次' or not cols[0].text.strip().isdigit():
                    continue

                texts = [ele.text.strip().replace('\n', '').replace('\r', '').replace('\t', '') for ele in cols]
                
                # --- 1. 抓取案號 ---
                strings = list(cols[2].stripped_strings)
                tender_no = strings[0] if len(strings) > 0 else ""
                
                # --- 2. 精準抓取標案名稱 ---
                tender_name = ""
                a_tag = cols[2].find('a')
                if a_tag and a_tag.get('title'):
                    tender_name = a_tag.get('title').strip()
                elif a_tag:
                    tender_name = a_tag.text.strip()
                else:
                    span_tag = cols[2].find('span')
                    if span_tag:
                        tender_name = span_tag.text.strip()

                # 最後防線：無差別文字扣除法
                if not tender_name:
                    full_text = cols[2].get_text(separator=" ", strip=True)
                    tender_name = full_text.replace(tender_no, "", 1).strip()

                # --- 3. 抓取超連結 ---
                tender_link = ""
                if a_tag and a_tag.get('href'):
                    href = a_tag.get('href')
                    if "javascript" not in href.lower() and href != "#":
                        tender_link = urllib.parse.urljoin(base_url, href)

                # --- 4. 組合資料 ---
                data.append({
                    "項次": texts[0],
                    "機關名稱": texts[1],
                    "標案案號": tender_no,
                    "標案名稱": tender_name,
                    "標案連結": tender_link, # 之後會在信件中轉換成超連結
                    "招標方式": texts[3],
                    "標的分類": texts[4],
                    "公告日期": texts[5],
                    "決標金額": texts[6]
                })

    except Exception as e:
        print(f"執行過程中發生錯誤: {e}")
    finally:
        driver.quit()

    return data

def send_email(df):
    sender_email = os.environ.get('GMAIL_USER')
    sender_password = os.environ.get('GMAIL_APP_PASSWORD')
    receiver_email = os.environ.get('RECEIVER_EMAIL')

    if not sender_email or not sender_password:
        print("未設定 Email 環境變數，無法發送郵件 (請確認 GitHub Secrets 設定)。")
        return

    # 信件標題使用台灣時間
    tw_time = datetime.utcnow() + timedelta(hours=8)
    date_str = tw_time.strftime('%Y-%m-%d')

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = f"政府採購公告決標通知 ({date_str})"

    # 將網址與名稱結合成 HTML 點擊標籤
    if '標案連結' in df.columns and '標案名稱' in df.columns:
        df['標案名稱'] = df.apply(
            lambda x: f"<a href='{x['標案連結']}' target='_blank'>{x['標案名稱']}</a>" if x['標案連結'] else x['標案名稱'], 
            axis=1
        )
        # 刪除獨立的連結欄位，保持表格乾淨
        df = df.drop(columns=['標案連結'])

    # 轉換為 HTML 表格
    html_table = df.to_html(index=False, escape=False, border=0, classes="styled-table")

    # CSS 樣式與郵件內容設計
    html_content = f"""
    <html>
    <head>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            color: #333333;
            line-height: 1.6;
        }}
        .styled-table {{
            border-collapse: collapse;
            margin: 25px 0;
            font-size: 0.9em;
            min-width: 800px;
            box-shadow: 0 0 20px rgba(0, 0, 0, 0.1);
        }}
        .styled-table thead tr {{
            background-color: #007BFF;
            color: #ffffff;
            text-align: left;
        }}
        .styled-table th, .styled-table td {{
            padding: 12px 15px;
            border: 1px solid #dddddd;
        }}
        .styled-table tbody tr {{
            border-bottom: 1px solid #dddddd;
        }}
        .styled-table tbody tr:nth-of-type(even) {{
            background-color: #f9f9f9;
        }}
        .styled-table tbody tr:last-of-type {{
            border-bottom: 2px solid #007BFF;
        }}
        a {{
            color: #007BFF;
            text-decoration: none;
            font-weight: bold;
        }}
        a:hover {{
            text-decoration: underline;
            color: #0056b3;
        }}
    </style>
    </head>
    <body>
        <h3 style="color: #007BFF;">每日電動車決標速報 🚗</h3>
        <p>您好，以下為最近 3 天的相關標案資料。本日共為您抓取到 <b>{len(df)}</b> 筆最新資訊：</p>
        
        {html_table}
        
        <br>
        <p style="color: #999999; font-size: 0.85em;">本信件由 GitHub Actions 每日自動執行發送，請勿直接回覆。</p>
    </body>
    </html>
    """

    msg.attach(MIMEText(html_content, 'html', 'utf-8'))

    # 發送郵件
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()
        print("✅ HTML 郵件發送成功！")
    except Exception as e:
        print(f"❌ 郵件發送失敗: {e}")

# ==========================================
# 主程式執行區塊
# ==========================================
if __name__ == "__main__":
    print("--- 開始執行政府採購網自動化爬蟲 ---")
    results = scrape_pcc_tenders()
    
    # 防呆機制：確保資料存在且大於 0 筆
    if results and len(results) > 0:
        df = pd.DataFrame(results)
        
        # 雙重確認表格不是空的
        if not df.empty:
            print(f"成功抓取 {len(df)} 筆資料，準備寄發 HTML 信件...")
            send_email(df)
        else:
            print("表格內容為空，不發送郵件。")
    else:
        print("沒有查詢到新案件，郵件不用發出，程式自動結束。")
