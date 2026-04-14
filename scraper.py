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

def get_dynamic_urls():
    # 預設抓取最近 3 天的資料 (校正為台灣時間 UTC+8)
    tw_time = datetime.utcnow() + timedelta(hours=8)
    start_date = tw_time - timedelta(days=3)
    
    start_str = start_date.strftime("%Y/%m/%d")
    end_str = tw_time.strftime("%Y/%m/%d")
    
    start_encoded = urllib.parse.quote(start_str, safe='')
    end_encoded = urllib.parse.quote(end_str, safe='')
    
    print(f"🔍 搜尋日期區間：{start_str} ~ {end_str}")
    
    # 您想要查詢的關鍵字清單 (可隨時在此擴充)
    keywords = ["電動車", "電動汽車", "電動 車"]
    urls = []
    
    for kw in keywords:
        kw_encoded = urllib.parse.quote(kw, safe='')
        url = f"https://web.pcc.gov.tw/prkms/tender/common/agent/readTenderAgent?pageSize=50&firstSearch=false&isQuery=&isBinding=N&isLogIn=N&orgName=&orgId=&tenderName={kw_encoded}&tenderId=&tenderStatus=TENDER_STATUS_1&tenderWay=TENDER_WAY_ALL_DECLARATION&awardAnnounceStartDate={start_encoded}&awardAnnounceEndDate={end_encoded}&radProctrgCate=&tenderRange=TENDER_RANGE_ALL&minBudget=&maxBudget=&item=&gottenVendorName=&gottenVendorId=&submitVendorName=&submitVendorId=&execLocation=&priorityCate=&radReConstruct=&policyAdvocacy=&isCpp="
        urls.append((kw, url))
        
    return urls

def scrape_pcc_tenders():
    urls_info = get_dynamic_urls()
    
    # 設定 Selenium 參數
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
    all_data = []

    try:
        # 輪流造訪每一個關鍵字的網址
        for kw, url in urls_info:
            print(f"▶ 正在搜尋關鍵字：【{kw}】...")
            driver.get(url)
            time.sleep(5) # 等待 JS 渲染
            
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
                    all_data.append({
                        "項次": texts[0],  # 這個項次稍後會因為去重而重排
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

    return all_data

def send_email(df):
    sender_email = os.environ.get('GMAIL_USER')
    sender_password = os.environ.get('GMAIL_APP_PASSWORD')
    receiver_emails_str = os.environ.get('RECEIVER_EMAIL')

    if not sender_email or not sender_password or not receiver_emails_str:
        print("未設定完整的 Email 環境變數，無法發送郵件。")
        return

    receivers_list = [email.strip() for email in receiver_emails_str.split(',')]

    tw_time = datetime.utcnow() + timedelta(hours=8)
    date_str = tw_time.strftime('%Y-%m-%d')

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ", ".join(receivers_list)
    msg['Subject'] = f"電動車標案決標通知 ({date_str})"

    content_lines = []
    # 使用 iterrows 時，利用迴圈的 index 重新編排顯示的項次
    for display_index, (index, row) in enumerate(df.iterrows(), start=1):
        item_html = f"""
        <div style="margin-bottom: 20px;">
            <h4 style="color: #0056b3; margin-top: 0; margin-bottom: 10px; font-size: 16px;">
                【第 {display_index} 筆】{row['標案名稱']}
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
    
    items_html_string = "".join(content_lines)

    html_content = f"""
    <html>
    <head></head>
    <body style="font-family: '微軟正黑體', 'Segoe UI', sans-serif; color: #333333; line-height: 1.5; padding: 10px;">
        <h3 style="color: #007BFF; border-bottom: 2px solid #007BFF; padding-bottom: 5px;">每日電動車決標案件速報 🚗</h3>
        <p style="font-size: 15px;">您好，以下為最近 3 天的相關決標案件。本日共為您抓取並過濾出 <b>{len(df)}</b> 筆最新不重複資訊：</p>
        <br>
        {items_html_string}
        <p style="color: #999999; font-size: 12px; margin-top: 30px;">
            本信件由 AutoMail 每日自動執行發送，請勿直接回覆。
        </p>
    </body>
    </html>
    """

    msg.attach(MIMEText(html_content, 'html', 'utf-8'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receivers_list, msg.as_string())
        server.quit()
        print(f"✅ HTML 郵件成功發送給 {len(receivers_list)} 位收件者！")
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
        
        # --- 關鍵防呆：依據「標案案號」去除重複的資料 ---
        original_count = len(df)
        df = df.drop_duplicates(subset=['標案案號'], keep='first').reset_index(drop=True)
        final_count = len(df)
        
        if original_count > final_count:
            print(f"🔄 自動過濾了 {original_count - final_count} 筆重複符合多個關鍵字的標案。")
        
        if not df.empty:
            print(f"成功整理出 {final_count} 筆最新資料，準備寄發 HTML 信件...")
            #send_email(df)
        else:
            print("表格內容去重後為空，不發送郵件。")
    else:
        print("沒有查詢到任何新案件，郵件不用發出，程式自動結束。")
