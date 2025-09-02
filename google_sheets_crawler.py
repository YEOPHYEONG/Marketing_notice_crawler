import requests
from bs4 import BeautifulSoup
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import os
import time
from datetime import datetime, timezone, timedelta
import json

# Google Sheets 연동을 위한 라이브러리
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- 1. 설정 및 전역 변수 ---
PROCESSED_LINKS_FILE = 'processed_links.txt'

# --- 2. Google Sheets 연동 함수들 ---
def get_gspread_client():
    """gspread 클라이언트를 인증하고 반환합니다."""
    try:
        creds_json_str = os.environ.get('GOOGLE_API_CREDENTIALS')
        if not creds_json_str:
            print("❌ GOOGLE_API_CREDENTIALS Secret이 설정되지 않았습니다.")
            return None
        creds_dict = json.loads(creds_json_str)
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        print(f"❌ Google Sheets 클라이언트 인증 실패: {e}")
        return None

def load_settings_from_sheets(client, sheet_name):
    """'Settings' 시트에서 설정을 불러옵니다."""
    try:
        sheet = client.open(sheet_name).worksheet("Settings")
        settings_raw = sheet.get_all_records()
        settings = {item['Setting']: item['Value'] for item in settings_raw}
        
        keywords = [k.strip() for k in settings.get('Keywords (comma-separated)', '').split(',')]
        receiver_email = settings.get('Receiver Email')

        if not receiver_email or not keywords:
            print("❌ 'Settings' 시트에 'Receiver Email' 또는 'Keywords' 설정이 없습니다.")
            return None, None

        print("✅ 'Settings' 시트 로드 성공.")
        return keywords, receiver_email
    except gspread.exceptions.WorksheetNotFound:
        print(f"❌ '{sheet_name}' 파일에 'Settings' 시트가 없습니다.")
        return None, None
    except Exception as e:
        print(f"❌ 'Settings' 시트 로드 중 오류 발생: {e}")
        return None, None

def load_targets_from_sheets(client, sheet_name):
    """'Crawl_Targets' 시트에서 크롤링 대상을 불러옵니다."""
    try:
        sheet = client.open(sheet_name).worksheet("Crawl_Targets")
        records = sheet.get_all_records()
        print(f"✅ 'Crawl_Targets' 시트에서 {len(records)}개의 대상을 불러왔습니다.")
        return records
    except gspread.exceptions.WorksheetNotFound:
        print(f"❌ '{sheet_name}' 파일에 'Crawl_Targets' 시트가 없습니다.")
        return []
    except Exception as e:
        print(f"❌ 'Crawl_Targets' 시트 로드 중 오류 발생: {e}")
        return []

def save_announcements_to_sheet(client, sheet_name, announcements):
    """'Collected_Announcements' 시트에 새로운 공고를 저장합니다."""
    if not announcements:
        return
    print(f"\n--- Google Sheets에 {len(announcements)}개의 신규 공고 저장 시도 ---")
    try:
        sheet = client.open(sheet_name).worksheet("Collected_Announcements")
        rows_to_add = []
        
        kst = timezone(timedelta(hours=9))
        collected_time_kst = datetime.now(kst).strftime('%Y-%m-%d %H:%M:%S')

        for ann in announcements:
            # [수정] 수집일, 회사, 제목, 공고일, 링크 순서로 리스트 생성
            row = [
                collected_time_kst,
                ann['company'],
                ann['title'],
                ann.get('date', 'N/A'),  # 공고일 추가
                ann['href']
            ]
            rows_to_add.append(row)
        
        sheet.append_rows(rows_to_add)
        print("✅ Google Sheets에 신규 공고 저장 완료.")
    except gspread.exceptions.WorksheetNotFound:
        print(f"❌ '{sheet_name}' 파일에 'Collected_Announcements' 시트가 없습니다.")
    except Exception as e:
        print(f"❌ Google Sheets 저장 중 오류 발생: {e}")

# --- 3. 크롤러 핵심 함수들 ---
def send_email(subject, body, receiver_email):
    print("\n--- 이메일 발송 시도 ---")
    try:
        smtp_user = os.environ.get('GMAIL_USER')
        smtp_password = os.environ.get('GMAIL_PASSWORD')
        if not smtp_user or not smtp_password:
            print("❌ GMAIL_USER 또는 GMAIL_PASSWORD Secret이 설정되지 않았습니다.")
            return
    except Exception as e:
        print(f"❌ GitHub Secrets 로드 실패: {e}")
        return
    msg = MIMEText(body, 'html', 'utf-8')
    msg['Subject'] = Header(subject, 'utf-8')
    msg['From'] = smtp_user
    msg['To'] = receiver_email
    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.sendmail(msg['From'], [msg['To']], msg.as_string())
        print(f"✅ 이메일 발송 성공: {subject}")
    except Exception as e:
        print(f"❌ 이메일 발송 실패: {e}")

def load_processed_links():
    if not os.path.exists(PROCESSED_LINKS_FILE): return set()
    with open(PROCESSED_LINKS_FILE, 'r', encoding='utf-8') as f: return set(line.strip() for line in f)

def save_processed_link(link):
    with open(PROCESSED_LINKS_FILE, 'a', encoding='utf-8') as f: f.write(link + '\n')

def generate_summary_email_body(announcements):
    kst = timezone(timedelta(hours=9))
    # [수정] 이메일 테이블 헤더에 '공고일' 추가
    html = """<head><style>body{font-family:sans-serif}.container{border:1px solid #ddd;padding:20px;margin:20px;border-radius:8px}h2{color:#005aab}table{width:100%;border-collapse:collapse}th,td{border:1px solid #ddd;padding:12px;text-align:left}th{background-color:#f2f2f2}a{color:#005aab;text-decoration:none}a:hover{text-decoration:underline}.footer{margin-top:20px;font-size:12px;color:#888}</style></head><body><div class="container"><h2>📢 신규 입찰 공고 요약</h2><p><strong>""" + datetime.now(kst).strftime('%Y년 %m월 %d일') + """</strong>에 발견된 신규 공고 목록입니다.</p><table><thead><tr><th>회사명</th><th>공고일</th><th>공고 제목</th></tr></thead><tbody>"""
    for ann in announcements:
        # [수정] 이메일 테이블 행에 공고일 데이터 추가
        html += f"""<tr><td>{ann['company']}</td><td>{ann.get('date', 'N/A')}</td><td><a href="{ann['href']}">{ann['title']}</a></td></tr>"""
    html += """</tbody></table><p class="footer">본 메일은 자동화된 스크립트에 의해 발송되었습니다.</p></div></body>"""
    return html

def crawl_site(target, keywords, processed_links):
    company, url, selector, base_url = target.get('company','N/A'), target.get('url'), target.get('selector'), target.get('base_url','')
    new_announcements = []
    if not all([url, selector]): print(f"🟡 경고: '{company}'의 url 또는 selector가 비어있어 건너뜁니다."); return new_announcements
    print(f"\n--- '{company}' 사이트 크롤링 시작 ---")
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
    except requests.RequestException as e: print(f"❌ '{company}' 사이트 접속 실패: {e}"); return new_announcements
    soup = BeautifulSoup(response.text, 'html.parser')
    links = soup.select(selector)
    if not links: print(f"🟡 경고: '{company}'에서 '{selector}' 선택자에 해당하는 링크를 찾지 못했습니다."); return new_announcements
    for link in links:
        title = link.get_text(strip=True)
        href = link.get('href', '')
        
        # [수정] 공고일(post_date)을 크롤링하는 로직 추가
        post_date = "N/A"
        try:
            # 링크(a) 태그의 부모인 tr 태그를 찾고, 그 안에서 class가 'date'인 td를 찾습니다.
            parent_row = link.find_parent('tr')
            if parent_row:
                date_cell = parent_row.find('td', class_='date')
                if date_cell:
                    post_date = date_cell.get_text(strip=True)
        except Exception:
            pass # 날짜를 찾지 못해도 오류 없이 진행

        if href and not href.startswith('http'): href = base_url.rstrip('/') + '/' + href.lstrip('/')
        if any(keyword.lower() in title.lower() for keyword in keywords) and href and href not in processed_links:
            # [수정] 로그에 공고일 추가
            print(f"🚀 새로운 공고 발견: [{company}] {title} (공고일: {post_date})")
            # [수정] 수집 데이터에 공고일 추가
            new_announcements.append({"company": company, "title": title, "href": href, "date": post_date})
            save_processed_link(href)
            processed_links.add(href)
            
    if not new_announcements: print(f"ℹ️ '{company}'에서 키워드에 맞는 새로운 공고를 찾지 못했습니다.")
    return new_announcements

# --- 4. 메인 실행 로직 ---
def main():
    print("="*50 + "\nGoogle Sheets 연동 입찰 공고 크롤러 (v2)를 시작합니다.\n" + "="*50)
    
    google_sheet_filename = "마케팅 공고 크롤러"

    client = get_gspread_client()
    if not client:
        return

    keywords_to_find, email_to_receive = load_settings_from_sheets(client, google_sheet_filename)
    targets = load_targets_from_sheets(client, google_sheet_filename)
    
    if not targets or not keywords_to_find or not email_to_receive:
        print("크롤링에 필요한 설정 정보가 부족하여 작업을 종료합니다.")
        return
        
    processed_links = load_processed_links()
    all_new_announcements = []

    for target in targets:
        new_finds = crawl_site(target, keywords_to_find, processed_links)
        if new_finds:
            all_new_announcements.extend(new_finds)
        time.sleep(1)

    print("\n--- 모든 사이트 크롤링 완료 ---")

    if all_new_announcements:
        save_announcements_to_sheet(client, google_sheet_filename, all_new_announcements)
        count = len(all_new_announcements)
        print(f"\n총 {count}개의 신규 공고를 발견하여 요약 이메일을 발송합니다.")
        subject = f"[입찰 공고] {count}개의 신규 공고가 있습니다."
        body = generate_summary_email_body(all_new_announcements)
        send_email(subject, body, email_to_receive)
    else:
        print("\nℹ️ 모든 사이트에서 새로운 공고를 찾지 못했습니다.")

if __name__ == '__main__':
    main()

