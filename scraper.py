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
    start_date = tw_time - timedelta(days=3)
    
    start_str = start_date.strftime("%Y/%m/%d")
    end_str = tw_time.strftime("%Y/%m/%d")
    
    # 將日期轉為 URL 編碼
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
                    "標案連結": tender_link,
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
    # ==========================================
    # 恢復使用環境變數讀取 (請確保 GitHub Secrets 已設定這三個變數)
    # ==========================================
    sender_email = os.environ.get('GMAIL_USER')
    sender_password = os.environ.get('GMAIL_APP_PASSWORD')
    receiver_email = os.environ.get('RECEIVER_EMAIL')

    if not sender_email or not sender_password or not receiver_email:
        print("未設定完整的 Email 環境變數，無法發送郵件。")
        return

    tw_time = datetime.utcnow() + timedelta(hours=8)
    date_str = tw_time.strftime('%Y-%m-%d')

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = f"電動車標案通知 ({date_str})"

    # ==========================================
    # 將 DataFrame 轉為 Outlook 友善的「條列式 + 分隔線」HTML
    # ==========================================
    content_lines = []
    for index, row in df.iterrows():
        # 每筆資料的專屬區塊
        item_html = f"""
        <div style="margin-bottom: 20px;">
            <h4 style="color: #0056b3; margin-top: 0; margin-bottom: 10px; font-size: 16px;">
                【第 {index + 1} 筆】{row['標案名稱']}
            </h4>
            <p style="margin: 3px 0;"><b>機關名稱：</b>{row['機關名稱']}</p>
            <p style="margin: 3px 0;"><b>標案案號：</b>{row['標案案號']}</p>
            <p style="margin: 3px 0;"><b>招標方式：</b>{row['招標方式']}</p>
            <p style="margin: 3px 0;"><b>標的分類：</b>{row['標的分類']}</p>
            <p style="margin: 3px 0;"><b>公告日期：</b>{row['公告日期']}</p>
            <p style="margin: 3px 0;"><b>決標金額：</b>{row['決標金額']}</p>
            <p style="margin-top: 10px;">
                <a href="{row['標案連結']}" style="color: #ffffff; background-color: #007BFF; padding: 6px 12px; text-decoration: none; border-radius: 4px; font-size: 14px; display: inline-block;">
                    🔍 點擊查看詳細內容
                </a>
            </p>
        </div>
        <hr style="border: 0; border-top: 1px dashed #cccccc; margin: 20px 0;">
        """
        content_lines.append(item_html)
    
    # 將所有條列資料組合起來
    items_html_string = "".join(content_lines)

    # 主體架構
    html_content = f"""
    <html>
    <head></head>
    <body style="font-family: '微軟正黑體', 'Segoe UI', sans-serif; color: #333333; line-height: 1.5; padding: 10px;">
        <h3 style="color: #007BFF; border-bottom: 2px solid #007BFF; padding-bottom: 5px;">每日電動車標案速報 🚗</h3>
        <p style="font-size: 15px;">您好，以下為最近 3 天的相關標案資料。本日共為您抓取到 <b>{len(df)}</b> 筆最新資訊：</p>
        <br>
        
        {items_html_string}
        
        <p style="color: #999999; font-size: 12px; margin-top: 30px;">
            本信件由 GitHub Actions 每日自動執行發送，請勿直接回覆。
        </p>
    </body>
    </html>
    """

    msg.attach(MIMEText(html_content, 'html', 'utf-8'))

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
    
    if results and len(results) > 0:
        df = pd.DataFrame(results)
        
        if not df.empty:
            print(f"成功抓取 {len(df)} 筆資料，準備寄發 HTML 信件...")
            send_email(df)
        else:
            print("表格內容為空，不發送郵件。")
    else:
        print("沒有查詢到新案件，郵件不用發出，程式自動結束。")
