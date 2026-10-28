import os
import sys
import time
import urllib.parse
from datetime import datetime, timedelta, timezone
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pandas as pd

ZHIJIE_LABEL = "智捷得標"
PAGE_LOAD_TIMEOUT = 60   # 單次載入最多等幾秒（需小於 Selenium 內部的 120 秒上限）
MAX_ATTEMPTS = 3         # 每個網址最多重試幾次

def save_debug_page(html):
    try:
        with open("debug_page.html", "w", encoding="utf-8") as f:
            f.write(html or "")
        print("   [selenium] 已將收到的網頁存成 debug_page.html", flush=True)
    except Exception as e:
        print(f"   [selenium] 存 debug_page.html 失敗：{e}", flush=True)

def load_page(driver, url):
    """載入網頁；逾時就停止載入，若結果表格已出現就直接使用，否則重試。"""
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            driver.get(url)
            return True
        except TimeoutException:
            print(f"   第 {attempt} 次載入逾時（{PAGE_LOAD_TIMEOUT} 秒）", flush=True)
            try:
                driver.execute_script("window.stop();")
                if 'id="atm"' in driver.page_source:
                    print("   結果表格已載入，直接使用目前內容", flush=True)
                    return True
            except Exception:
                pass
            if attempt < MAX_ATTEMPTS:
                time.sleep(15)
    return False

def describe_page(label, html):
    # 【診斷】印出網頁摘要，判斷是否拿到正常的查詢結果頁
    soup = BeautifulSoup(html, "html.parser")
    title = soup.title.get_text(strip=True) if soup.title else "(無 title)"
    has_table = soup.find("table", id="atm") is not None
    print(f"   [{label}] 長度={len(html)}，title={title}，結果表格(atm)={'有' if has_table else '無'}", flush=True)
    if not has_table:
        snippet = soup.get_text(" ", strip=True)[:500]
        print(f"   [{label}] 頁面文字前 500 字：{snippet}", flush=True)
    return has_table

def get_dynamic_urls():
    # 預設抓取最近 3 天的資料 (校正為台灣時間 UTC+8)
    tw_time = datetime.now(timezone.utc) + timedelta(hours=8)
    start_date = tw_time - timedelta(days=3)
    
    start_str = start_date.strftime("%Y/%m/%d")
    end_str = tw_time.strftime("%Y/%m/%d")
    
    start_encoded = urllib.parse.quote(start_str, safe='')
    end_encoded = urllib.parse.quote(end_str, safe='')
    
    print(f"🔍 搜尋日期區間：{start_str} ~ {end_str}", flush=True)
    
    # 【階段一：寬鬆抓取】只搜尋「電動」大類/
    keywords = ["電動"]
    urls = []
    
    def build_url(kw, gotten_vendor=""):
        kw_encoded = urllib.parse.quote(kw, safe='')
        vendor_encoded = urllib.parse.quote(gotten_vendor, safe='')
        return f"https://web.pcc.gov.tw/prkms/tender/common/agent/readTenderAgent?pageSize=50&firstSearch=false&isQuery=&isBinding=N&isLogIn=N&orgName=&orgId=&tenderName={kw_encoded}&tenderId=&tenderStatus=TENDER_STATUS_1&tenderWay=TENDER_WAY_ALL_DECLARATION&awardAnnounceStartDate={start_encoded}&awardAnnounceEndDate={end_encoded}&radProctrgCate=&tenderRange=TENDER_RANGE_ALL&minBudget=&maxBudget=&item=&gottenVendorName={vendor_encoded}&gottenVendorId=&submitVendorName=&submitVendorId=&execLocation=&priorityCate=&radReConstruct=&policyAdvocacy=&isCpp="

    for kw in keywords:
        urls.append((kw, build_url(kw)))

    # 【智捷得標】同區間、同關鍵字，但限定得標廠商名稱含「智捷」，用來在信件中標註
    urls.append((ZHIJIE_LABEL, build_url("電動", gotten_vendor="智捷")))

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
    # 只要 DOM 載好就繼續，不等圖片、外部統計碼等資源（避免被慢資源卡住）
    chrome_options.page_load_strategy = "eager"

    print("正在啟動虛擬瀏覽器...", flush=True)
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)

    # 【關鍵修復】採購網防火牆會阻擋舊版 Chrome 與 HeadlessChrome 的 User-Agent，
    # 因此改用本機實際安裝的 Chrome 版本號，並移除 "Headless" 字樣
    chrome_major = driver.capabilities.get("browserVersion", "141").split(".")[0]
    user_agent = f"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/{chrome_major}.0.0.0 Safari/537.36"
    driver.execute_cdp_cmd("Network.setUserAgentOverride", {"userAgent": user_agent})
    print(f"   使用 User-Agent：{user_agent}", flush=True)

    base_url = "https://web.pcc.gov.tw"
    all_data = []
    page_ok = True

    try:
        # 輪流造訪每一個關鍵字的網址
        for kw, url in urls_info:
            print(f"▶ 正在網站上搜尋【{kw}】的案件...", flush=True)
            if not load_page(driver, url):
                print(f"   ❌ 【{kw}】重試 {MAX_ATTEMPTS} 次仍無法載入（採購網可能限制海外連線）", flush=True)
            time.sleep(5) # 等待 JS 渲染

            html = driver.page_source
            print(f"   [selenium] 最終網址={driver.current_url}", flush=True)
            if not describe_page("selenium", html):
                page_ok = False
                save_debug_page(html)

            soup = BeautifulSoup(html, "html.parser")
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
                        "查詢": kw,
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
        print(f"執行過程中發生錯誤: {e}", flush=True)
        page_ok = False
        try:
            save_debug_page(driver.page_source)
        except Exception:
            pass
    finally:
        driver.quit()

    return all_data, page_ok

def send_email(df):
    print("▶ 進入發信模組，開始讀取環境變數...", flush=True)
    sender_email = os.environ.get('GMAIL_USER')
    sender_password = os.environ.get('GMAIL_APP_PASSWORD')
    receiver_emails_str = os.environ.get('RECEIVER_EMAIL')

    # 印出變數讀取狀態 (故意遮蔽密碼以策安全)
    pwd_status = "✅ 已成功讀取" if sender_password else "❌ 未讀取到 (空的)"
    print(f"   - 發信帳號 (GMAIL_USER): {sender_email}", flush=True)
    print(f"   - 應用密碼 (GMAIL_APP_PASSWORD): {pwd_status}", flush=True)
    print(f"   - 收件信箱 (RECEIVER_EMAIL): {receiver_emails_str}", flush=True)

    if not sender_email or not sender_password or not receiver_emails_str:
        print("❌ 致命錯誤：環境變數有缺漏，程式強制取消寄信！請檢查 GitHub Secrets。", flush=True)
        return

    receivers_list = [email.strip() for email in receiver_emails_str.split(',')]

    tw_time = datetime.now(timezone.utc) + timedelta(hours=8)
    date_str = tw_time.strftime('%Y-%m-%d')

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ", ".join(receivers_list)
    zhijie_count = int(df['智捷得標'].sum())
    zhijie_prefix = f"【智捷得標 {zhijie_count} 件】" if zhijie_count else ""
    msg['Subject'] = f"{zhijie_prefix}電動車/相關載具標案決標通知 ({date_str})"

    content_lines = []
    # 使用 iterrows 重新編排顯示的項次
    for display_index, (index, row) in enumerate(df.iterrows(), start=1):
        if row['智捷得標']:
            box_style = "margin-bottom: 20px; background-color: #FFF4E5; border-left: 5px solid #FF8C00; padding: 10px 12px;"
            badge = '<span style="background-color: #FF8C00; color: #ffffff; padding: 2px 8px; border-radius: 4px; font-size: 13px; margin-right: 6px;">🏆 智捷得標</span>'
        else:
            box_style = "margin-bottom: 20px;"
            badge = ""
        item_html = f"""
        <div style="{box_style}">
            <h4 style="color: #0056b3; margin-top: 0; margin-bottom: 10px; font-size: 16px;">
                {badge}【第 {display_index} 筆】{row['標案名稱']}
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

    if zhijie_count:
        zhijie_list = "".join(f"<li>{r['機關名稱']}｜{r['標案名稱']}｜決標金額 {r['決標金額']}</li>" for _, r in df[df['智捷得標']].iterrows())
        zhijie_summary = f"""
        <div style="background-color: #FFF4E5; border: 1px solid #FF8C00; border-radius: 6px; padding: 10px 14px; margin-bottom: 10px;">
            <p style="margin: 0 0 6px 0; font-size: 15px; color: #C05600;"><b>🏆 本期「智捷」得標案件：{zhijie_count} 件</b></p>
            <ul style="margin: 0; padding-left: 20px;">{zhijie_list}</ul>
        </div>"""
    else:
        zhijie_summary = '<p style="font-size: 14px; color: #666666;">本期沒有「智捷」得標的案件。</p>'

    html_content = f"""
    <html>
    <head></head>
    <body style="font-family: '微軟正黑體', 'Segoe UI', sans-serif; color: #333333; line-height: 1.5; padding: 10px;">
        <h3 style="color: #007BFF; border-bottom: 2px solid #007BFF; padding-bottom: 5px;">每日電動車/汽車決標案件速報 🚗</h3>
        <p style="font-size: 15px;">您好，以下為最近 3 天的相關決標案件。本日共為您抓取並篩選出 <b>{len(df)}</b> 筆包含「電動」與「車」的最新資訊：</p>
        {zhijie_summary}
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
        print("▶ 正在連線至 Gmail SMTP 伺服器...", flush=True)
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        
        print("▶ 正在嘗試登入 Gmail...", flush=True)
        server.login(sender_email, sender_password)
        
        print(f"▶ 登入成功！準備發送郵件給 {len(receivers_list)} 位收件者...", flush=True)
        server.sendmail(sender_email, receivers_list, msg.as_string())
        server.quit()
        print(f"✅ HTML 郵件成功發送給 {len(receivers_list)} 位收件者！", flush=True)
    except smtplib.SMTPAuthenticationError as e:
        print(f"❌ 登入失敗！請檢查您的應用程式密碼是否正確。詳細錯誤: {e}", flush=True)
    except Exception as e:
        print(f"❌ 郵件發送發生未知錯誤: {e}", flush=True)

# ==========================================
# 主程式執行區塊
# ==========================================
if __name__ == "__main__":
    print("--- 開始執行政府採購網自動化爬蟲 ---", flush=True)
    results, page_ok = scrape_pcc_tenders()

    if not page_ok:
        # 沒拿到正常的查詢結果頁 → 讓 GitHub Actions 顯示失敗，而不是假裝「沒有新案件」
        print("❌ 沒有取得正常的查詢結果頁（可能被網站阻擋），請查看上方診斷訊息。", flush=True)
        sys.exit(1)

    if any(r['查詢'] != ZHIJIE_LABEL for r in results):
        df = pd.DataFrame(results)

        # 智捷得標的案號 (來自「智捷得標」查詢)
        zhijie_ids = set(df.loc[df['查詢'] == ZHIJIE_LABEL, '標案案號'])
        print(f"🏆 其中「智捷」得標的案件：{len(zhijie_ids)} 筆", flush=True)

        # 以「電動」查詢結果為主，並標註是否為智捷得標
        df = df[df['查詢'] != ZHIJIE_LABEL].copy()
        df['智捷得標'] = df['標案案號'].isin(zhijie_ids)

        # 先去除重複的資料 (以防網頁有重複案號)
        df = df.drop_duplicates(subset=['標案案號'], keep='first')
        
        total_electric = len(df)
        print(f"✅ 第一階段：成功抓取到 {total_electric} 筆包含「電動」的案件。", flush=True)
        
        # ==========================================
        # 【階段二：本地精準過濾】只保留標案名稱中有「車」的案件
        # ==========================================
        print("▶ 第二階段：正在本地過濾名稱中包含「車」的案件...", flush=True)
        df = df[df['標案名稱'].str.contains('車', na=False) | df['智捷得標']]
        # 智捷得標的案件排在最前面
        df = df.sort_values('智捷得標', ascending=False, kind='stable').reset_index(drop=True)
        
        final_count = len(df)
        print(f"✅ 過濾完成：剃除了 {total_electric - final_count} 筆非車輛案件，最終保留 {final_count} 筆。", flush=True)
        
        # 階段三：判斷並發送郵件
        if not df.empty:
            print(f"準備將這 {final_count} 筆資料寄發 HTML 信件...", flush=True)
            send_email(df)
        else:
            print("經過「車」字過濾後，沒有符合的案件，不發送郵件。", flush=True)
    else:
        print("網站上沒有查詢到任何「電動」相關新案件，程式自動結束。", flush=True)
