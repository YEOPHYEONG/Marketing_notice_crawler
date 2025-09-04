import requests
from bs4 import BeautifulSoup
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import os
import time
from datetime import datetime, timezone, timedelta
import json
import msal  # Microsoft 인증 라이브러리

# --- 1. 설정 및 전역 변수 ---
PROCESSED_LINKS_FILE = 'processed_links.txt'

# --- 2. Microsoft Graph API 연동 함수들 ---
def get_ms_graph_access_token():
    """Azure AD에서 MS Graph API 접근 토큰을 발급받습니다."""
    tenant_id = os.environ.get('MS_TENANT_ID')
    client_id = os.environ.get('MS_CLIENT_ID')
    client_secret = os.environ.get('MS_CLIENT_SECRET')

    if not all([tenant_id, client_id, client_secret]):
        print("❌ MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET Secret이 설정되지 않았습니다.")
        return None

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id, authority=authority, client_credential=client_secret
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

    if "access_token" in result:
        print("✅ MS Graph API 토큰 발급 성공.")
        return result['access_token']
    else:
        print("❌ MS Graph API 토큰 발급 실패:", result.get("error_description"))
        return None

def get_excel_data(access_token, sheet_name):
    """MS Graph API를 통해 Excel 시트의 데이터를 불러옵니다."""
    user_principal_name = os.environ.get('MS_USER_PRINCIPAL_NAME')
    excel_file_path = os.environ.get('MS_EXCEL_FILE_PATH')

    if not all([user_principal_name, excel_file_path]):
        print("❌ MS_USER_PRINCIPAL_NAME 또는 MS_EXCEL_FILE_PATH Secret이 설정되지 않았습니다.")
        return []
    
    graph_url = f"https://graph.microsoft.com/v1.0/users/{user_principal_name}/drive/root:/{excel_file_path}:/workbook/tables('{sheet_name}')/rows"
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    
    try:
        response = requests.get(graph_url, headers=headers)
        response.raise_for_status()
        rows_data = response.json().get('value', [])
        
        header_url = f"https://graph.microsoft.com/v1.0/users/{user_principal_name}/drive/root:/{excel_file_path}:/workbook/tables('{sheet_name}')/headerRowRange"
        header_response = requests.get(header_url, headers=headers)
        header_response.raise_for_status()
        header = header_response.json()['values'][0]

        records = [dict(zip(header, row['values'][0])) for row in rows_data]
        print(f"✅ Excel '{sheet_name}' 시트에서 {len(records)}개의 행을 로드했습니다.")
        return records

    except requests.exceptions.HTTPError as e:
        print(f"❌ Excel '{sheet_name}' 시트 로드 실패 (HTTP {e.response.status_code}): {e.response.text}")
        print("   (Excel 파일 경로, 시트/표 이름, API 권한을 확인해주세요.)")
        return []
    except Exception as e:
        print(f"❌ Excel '{sheet_name}' 시트 처리 중 오류: {e}")
        return []

def save_announcements_to_excel(access_token, announcements):
    """[수정] 새로운 공고를 Excel 테이블의 맨 위에 삽입합니다."""
    if not announcements: return
    user_principal_name, excel_file_path = os.environ.get('MS_USER_PRINCIPAL_NAME'), os.environ.get('MS_EXCEL_FILE_PATH')
    sheet_name = "Collected_Announcements"
    graph_url = f"https://graph.microsoft.com/v1.0/users/{user_principal_name}/drive/root:/{excel_file_path}:/workbook/tables('{sheet_name}')/rows/add"
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    kst = timezone(timedelta(hours=9))
    collected_time_kst = datetime.now(kst).strftime('%Y-%m-%d %H:%M:%S')
    rows_to_add = [[collected_time_kst, ann['company'], ann['title'], ann.get('date', 'N/A'), ann['href']] for ann in announcements]
    
    # [수정] index: 0을 추가하여 테이블의 맨 위에 행을 삽입합니다.
    payload = {"values": rows_to_add, "index": 0}
    
    try:
        response = requests.post(graph_url, headers=headers, json=payload)
        response.raise_for_status()
        print("✅ Excel에 신규 공고 저장 완료.")
    except requests.exceptions.HTTPError as e:
        print(f"❌ Excel 저장 중 오류 발생 (HTTP {e.response.status_code}): {e.response.text}")
    except Exception as e:
        print(f"❌ Excel 저장 중 오류 발생: {e}")

# --- 3. 크롤러 핵심 함수들 ---
def send_email(subject, body, receiver_email):
    print("\n--- 이메일 발송 시도 ---")
    try:
        smtp_user, smtp_password = os.environ.get('GMAIL_USER'), os.environ.get('GMAIL_PASSWORD')
        if not all([smtp_user, smtp_password]): print("❌ GMAIL_USER 또는 GMAIL_PASSWORD Secret이 설정되지 않았습니다."); return
    except Exception as e: print(f"❌ GitHub Secrets 로드 실패: {e}"); return
    msg = MIMEText(body, 'html', 'utf-8')
    msg['Subject'], msg['From'], msg['To'] = Header(subject, 'utf-8'), smtp_user, receiver_email
    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls(); server.login(smtp_user, smtp_password)
            server.sendmail(msg['From'], [msg['To']], msg.as_string())
        print(f"✅ 이메일 발송 성공: {subject}")
    except Exception as e: print(f"❌ 이메일 발송 실패: {e}")

def load_processed_links():
    if not os.path.exists(PROCESSED_LINKS_FILE): return set()
    with open(PROCESSED_LINKS_FILE, 'r', encoding='utf-8') as f: return set(line.strip() for line in f)

def save_processed_link(link):
    with open(PROCESSED_LINKS_FILE, 'a', encoding='utf-8') as f: f.write(link + '\n')

def generate_summary_email_body(announcements):
    kst = timezone(timedelta(hours=9))
    html = """<head><style>body{font-family:sans-serif}.container{border:1px solid #ddd;padding:20px;margin:20px;border-radius:8px}h2{color:#005aab}table{width:100%;border-collapse:collapse}th,td{border:1px solid #ddd;padding:12px;text-align:left}th{background-color:#f2f2f2}a{color:#005aab;text-decoration:none}a:hover{text-decoration:underline}.footer{margin-top:20px;font-size:12px;color:#888}</style></head><body><div class="container"><h2>📢 신규 공고 요약</h2><p><strong>""" + datetime.now(kst).strftime('%Y년 %m월 %d일') + """</strong>에 발견된 신규 공고 목록입니다.</p><table><thead><tr><th>회사명</th><th>공고일</th><th>공고 제목</th></tr></thead><tbody>"""
    for ann in announcements:
        html += f"""<tr><td>{ann['company']}</td><td>{ann.get('date', 'N/A')}</td><td><a href="{ann['href']}">{ann['title']}</a></td></tr>"""
    html += """</tbody></table><p class="footer">본 메일은 자동화된 스크립트에 의해 발송되었습니다.</p></div></body>"""
    return html

def crawl_site(target, processed_links):
    company, url, base_url = target.get('company','N/A'), target.get('url'), target.get('base_url','')
    item_selector, title_link_selector, date_selector = target.get('item_selector'), target.get('title_link_selector'), target.get('date_selector')
    new_announcements = []
    if not all([url, item_selector, title_link_selector]):
        print(f"🟡 경고: '{company}'의 url, item_selector 또는 title_link_selector가 비어있어 건너뜁니다.")
        return new_announcements
    print(f"\n--- '{company}' 사이트 크롤링 시작 ---")
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
    except requests.RequestException as e:
        print(f"❌ '{company}' 사이트 접속 실패: {e}"); return new_announcements
    soup = BeautifulSoup(response.text, 'html.parser')
    items = soup.select(item_selector)
    if not items:
        print(f"🟡 경고: '{company}'에서 '{item_selector}' 선택자에 해당하는 항목을 찾지 못했습니다.")
        return new_announcements
        
    for item in items:
        title_link_element = item.select_one(title_link_selector)
        if not title_link_element: continue
        title, href = title_link_element.get_text(strip=True), title_link_element.get('href', '')
        post_date = "N/A"
        if date_selector:
            date_element = item.select_one(date_selector)
            if date_element: post_date = date_element.get_text(strip=True)
        if href and not href.startswith('http'):
            href = base_url.rstrip('/') + '/' + href.lstrip('/')
        
        if href and href not in processed_links:
            print(f"🚀 새로운 공고 발견: [{company}] {title} (공고일: {post_date})")
            new_announcements.append({"company": company, "title": title, "href": href, "date": post_date})
            save_processed_link(href)
            processed_links.add(href)
            
    if not new_announcements: print(f"ℹ️ '{company}'에서 새로운 공고를 찾지 못했습니다.")
    
    # [수정] reverse() 로직을 제거하여, 웹사이트에 보이는 순서 (보통 최신순) 그대로 리스트를 반환합니다.
    return new_announcements

# --- 4. 메인 실행 로직 ---
def main():
    print("="*50 + "\nMS Excel 연동 입찰 공고 크롤러 (v3 - 최신글 상단)를 시작합니다.\n" + "="*50)
    
    access_token = get_ms_graph_access_token()
    if not access_token: return

    settings_data = get_excel_data(access_token, "Settings")
    settings = {item['Setting']: item['Value'] for item in settings_data if 'Setting' in item and 'Value' in item}
    email_to_receive = settings.get('Receiver Email')

    targets = get_excel_data(access_token, "Crawl_Targets")
    
    if not targets or not email_to_receive:
        print("크롤링에 필요한 설정 정보(대상, 수신 이메일)가 부족하여 작업을 종료합니다.")
        return
        
    processed_links = load_processed_links()
    all_new_announcements = []

    for target in targets:
        if not target.get('company'): continue
        new_finds = crawl_site(target, processed_links)
        if new_finds:
            all_new_announcements.extend(new_finds)
        time.sleep(1)

    print("\n--- 모든 사이트 크롤링 완료 ---")

    if all_new_announcements:
        save_announcements_to_excel(access_token, all_new_announcements)
        count = len(all_new_announcements)
        subject = f"[신규 공고 알림] {count}개의 새로운 공고가 수집되었습니다."
        body = generate_summary_email_body(all_new_announcements)
        send_email(subject, body, email_to_receive)
    else:
        print("\nℹ️ 모든 사이트에서 새로운 공고를 찾지 못했습니다.")

if __name__ == '__main__':
    main()

