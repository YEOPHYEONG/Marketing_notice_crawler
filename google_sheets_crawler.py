import requests
from bs4 import BeautifulSoup
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import os
import time
from datetime import datetime
import json

# Google Sheets 연동을 위한 라이브러리
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- 1. 설정 및 전역 변수 ---
PROCESSED_LINKS_FILE = 'processed_links.txt'

# --- 2. Google Sheets 인증 및 데이터 로드 ---
def load_targets_from_sheets():
    """Google Sheets에서 크롤링 대상을 불러옵니다."""
    print("--- Google Sheets에서 크롤링 대상 로드 시작 ---")
    try:
        # GitHub Secret에서 JSON 인증 정보를 가져옵니다.
        creds_json_str = os.environ.get('GOOGLE_API_CREDENTIALS')
        if not creds_json_str:
            print("❌ GOOGLE_API_CREDENTIALS Secret이 설정되지 않았습니다.")
            return []
            
        creds_dict = json.loads(creds_json_str)
        
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        
        # [★★ 중요 ★★] 여기에 본인의 Google Sheet 파일 이름을 정확히 입력하세요.
        sheet_name = "나의 크롤러 설정 시트" 
        sheet = client.open(sheet_name).sheet1
        
        records = sheet.get_all_records()
        print(f"✅ Google Sheets에서 {len(records)}개의 크롤링 대상을 불러왔습니다.")
        return records

    except Exception as e:
        print(f"❌ Google Sheets 연동 실패: {e}")
        print("   (API 권한, 시트 공유, 시트 이름 등을 확인해주세요.)")
        return []

# --- 3. 크롤러 핵심 함수들 ---

def send_email(subject, body, receiver_email):
    """요약된 이메일을 발송하는 함수."""
    print("\n--- 이메일 발송 시도 ---")
    try:
        # GitHub Secrets에서 이메일 정보를 가져옵니다.
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
    """이미 처리된 링크 목록을 파일에서 불러옵니다."""
    if not os.path.exists(PROCESSED_LINKS_FILE):
        return set()
    with open(PROCESSED_LINKS_FILE, 'r', encoding='utf-8') as f:
        return set(line.strip() for line in f)

def save_processed_link(link):
    """새롭게 처리된 링크를 파일에 추가합니다."""
    with open(PROCESSED_LINKS_FILE, 'a', encoding='utf-8') as f:
        f.write(link + '\n')

def generate_summary_email_body(announcements):
    """공고 리스트를 받아 HTML 이메일 본문을 생성합니다."""
    html = """
    <head>
        <style>
            body { font-family: 'Malgun Gothic', sans-serif; } .container { border: 1px solid #ddd; padding: 20px; margin: 20px; border-radius: 8px; } h2 { color: #005AAB; } table { width: 100%; border-collapse: collapse; } th, td { border: 1px solid #ddd; padding: 12px; text-align: left; } th { background-color: #f2f2f2; } a { color: #005AAB; text-decoration: none; } a:hover { text-decoration: underline; } .footer { margin-top: 20px; font-size: 12px; color: #888; }
        </style>
    </head>
    <body>
        <div class="container">
            <h2>📢 신규 입찰 공고 요약</h2>
            <p><strong>""" + datetime.now().strftime('%Y년 %m월 %d일') + """</strong>에 발견된 신규 공고 목록입니다.</p>
            <table><thead><tr><th>회사명</th><th>공고 제목</th></tr></thead><tbody>
    """
    for ann in announcements:
        html += f"""<tr><td>{ann['company']}</td><td><a href="{ann['href']}">{ann['title']}</a></td></tr>"""
    html += """
            </tbody></table>
            <p class="footer">본 메일은 자동화된 스크립트에 의해 발송되었습니다.</p>
        </div>
    </body>
    """
    return html

def crawl_site(target, keywords, processed_links):
    """사이트를 크롤링하여 새로운 공고 리스트를 반환합니다."""
    company = target.get('company', 'N/A')
    url = target.get('url')
    selector = target.get('selector')
    base_url = target.get('base_url', '')
    new_announcements = []

    if not all([url, selector]):
        print(f"🟡 경고: '{company}'의 url 또는 selector가 비어있어 건너뜁니다.")
        return new_announcements
        
    print(f"\n--- '{company}' 사이트 크롤링 시작 ---")
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
    except requests.RequestException as e:
        print(f"❌ '{company}' 사이트 접속 실패: {e}")
        return new_announcements

    soup = BeautifulSoup(response.text, 'html.parser')
    links = soup.select(selector)

    if not links:
        print(f"🟡 경고: '{company}'에서 '{selector}' 선택자에 해당하는 링크를 찾지 못했습니다.")
        return new_announcements

    for link in links:
        title = link.get_text(strip=True)
        href = link.get('href', '')

        if href and not href.startswith('http'):
            href = base_url.rstrip('/') + '/' + href.lstrip('/')

        if any(keyword.lower() in title.lower() for keyword in keywords) and href and href not in processed_links:
            print(f"🚀 새로운 공고 발견: [{company}] {title}")
            new_announcements.append({"company": company, "title": title, "href": href})
            save_processed_link(href)
            processed_links.add(href)
    
    if not new_announcements:
        print(f"ℹ️ '{company}'에서 키워드에 맞는 새로운 공고를 찾지 못했습니다.")
    return new_announcements

# --- 4. 메인 실행 로직 ---
def main():
    """스크립트의 메인 실행 함수입니다."""
    print("="*50)
    print("Google Sheets 연동 입찰 공고 크롤러를 시작합니다.")
    print("="*50)
    
    targets = load_targets_from_sheets()
    if not targets:
        print("크롤링 대상이 없어 작업을 종료합니다.")
        return

    # [★★ 중요 ★★] 아래 키워드와 이메일 주소를 원하는 값으로 수정하세요.
    keywords_to_find = ["대행사", "입찰", "선정", "공고", "모집", "마케팅"]
    email_to_receive = "gooodong3@gmail.com"
    
    processed_links = load_processed_links()
    all_new_announcements = []

    for target in targets:
        new_finds = crawl_site(target, keywords_to_find, processed_links)
        if new_finds:
            all_new_announcements.extend(new_finds)
        time.sleep(1) # 사이트 부하를 줄이기 위한 지연

    print("\n--- 모든 사이트 크롤링 완료 ---")

    if all_new_announcements:
        count = len(all_new_announcements)
        print(f"\n총 {count}개의 신규 공고를 발견하여 요약 이메일을 발송합니다.")
        subject = f"[입찰 공고] {count}개의 신규 공고가 있습니다."
        body = generate_summary_email_body(all_new_announcements)
        send_email(subject, body, email_to_receive)
    else:
        print("\nℹ️ 모든 사이트에서 새로운 공고를 찾지 못했습니다.")

if __name__ == '__main__':
    main()

