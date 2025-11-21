import requests
from requests_html import HTMLSession # 동적 컨텐츠 렌더링을 위해 requests_html 사용
from bs4 import BeautifulSoup
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import os
import time
from datetime import datetime, timezone, timedelta
import json
import msal
import re
from dateutil.parser import parse as date_parse # 날짜 형식 표준화를 위해 dateutil 라이브러리 사용

# --- 1. 설정 및 전역 변수 ---
PROCESSED_LINKS_FILE = 'processed_links.txt'

# --- 2. Microsoft Graph API 연동 함수들 (기존과 동일) ---
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

        # 빈 행('')을 None으로 변환하여 일관성 유지
        records = [dict(zip(header, [val if val != '' else None for val in row['values'][0]])) for row in rows_data]
        print(f"✅ Excel '{sheet_name}' 시트에서 {len(records)}개의 행을 로드했습니다.")
        return records

    except requests.exceptions.HTTPError as e:
        print(f"❌ Excel '{sheet_name}' 시트 로드 실패 (HTTP {e.response.status_code}): {e.response.text}")
        print("     (Excel 파일 경로, 시트/표 이름, API 권한을 확인해주세요.)")
        return []
    except Exception as e:
        print(f"❌ Excel '{sheet_name}' 시트 처리 중 오류: {e}")
        return []

def save_announcements_to_excel(access_token, announcements):
    """새로운 공고를 Excel 테이블의 맨 위에 삽입합니다."""
    if not announcements: return
    user_principal_name, excel_file_path = os.environ.get('MS_USER_PRINCIPAL_NAME'), os.environ.get('MS_EXCEL_FILE_PATH')
    sheet_name = "Collected_Announcements"
    graph_url = f"https://graph.microsoft.com/v1.0/users/{user_principal_name}/drive/root:/{excel_file_path}:/workbook/tables('{sheet_name}')/rows/add"
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    kst = timezone(timedelta(hours=9))
    collected_time_kst = datetime.now(kst).strftime('%Y-%m-%d %H:%M:%S')
    rows_to_add = [[collected_time_kst, ann['company'], ann['title'], ann.get('date', 'N/A'), ann['href']] for ann in announcements]
    
    payload = {"values": rows_to_add, "index": 0}
    
    try:
        response = requests.post(graph_url, headers=headers, json=payload)
        response.raise_for_status()
        print("✅ Excel에 신규 공고 저장 완료.")
    except requests.exceptions.HTTPError as e:
        print(f"❌ Excel 저장 중 오류 발생 (HTTP {e.response.status_code}): {e.response.text}")
    except Exception as e:
        print(f"❌ Excel 저장 중 오류 발생: {e}")

# --- 3. 크롤러 유틸리티 함수들 ---
def send_email(subject, body, receiver_emails):
    """지정된 수신자 목록에게 이메일을 발송합니다."""
    if not receiver_emails:
        print("🟡 경고: 수신자 이메일 주소가 없어 이메일을 발송하지 않습니다.")
        return
        
    print(f"\n--- 이메일 발송 시도 ({', '.join(receiver_emails)}) ---")
    try:
        smtp_user, smtp_password = os.environ.get('GMAIL_USER'), os.environ.get('GMAIL_PASSWORD')
        if not all([smtp_user, smtp_password]):
            print("❌ GMAIL_USER 또는 GMAIL_PASSWORD Secret이 설정되지 않았습니다.")
            return
    except Exception as e:
        print(f"❌ GitHub Secrets 로드 실패: {e}")
        return
        
    msg = MIMEText(body, 'html', 'utf-8')
    msg['Subject'] = Header(subject, 'utf-8')
    msg['From'] = smtp_user
    msg['To'] = ", ".join(receiver_emails)
    
    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.sendmail(msg['From'], receiver_emails, msg.as_string())
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
    html = """<head><style>body{font-family:sans-serif}.container{border:1px solid #ddd;padding:20px;margin:20px;border-radius:8px}h2{color:#005aab}table{width:100%;border-collapse:collapse}th,td{border:1px solid #ddd;padding:12px;text-align:left}th{background-color:#f2f2f2}a{color:#005aab;text-decoration:none}a:hover{text-decoration:underline}.footer{margin-top:20px;font-size:12px;color:#888}</style></head><body><div class="container"><h2>📢 신규 공고 요약</h2><p><strong>""" + datetime.now(kst).strftime('%Y년 %m월 %d일') + """</strong>에 발견된 신규 공고 목록입니다.</p><table><thead><tr><th>회사명</th><th>공고일</th><th>공고 제목</th></tr></thead><tbody>"""
    for ann in announcements:
        html += f"""<tr><td>{ann['company']}</td><td>{ann.get('date', 'N/A')}</td><td><a href="{ann['href']}">{ann['title']}</a></td></tr>"""
    html += """</tbody></table><p class="footer">본 메일은 자동화된 스크립트에 의해 발송되었습니다.</p></div></body>"""
    return html

# --- [추가된 함수] ---
def generate_no_new_announcements_email_body():
    """신규 공고가 없을 때 발송할 이메일 본문을 생성합니다."""
    kst = timezone(timedelta(hours=9))
    html = """<head><style>body{font-family:sans-serif}.container{border:1px solid #ddd;padding:20px;margin:20px;border-radius:8px}h2{color:#005aab}.footer{margin-top:20px;font-size:12px;color:#888}</style></head><body><div class="container"><h2>📝 금일 신규 입찰 공고 없음</h2><p><strong>""" + datetime.now(kst).strftime('%Y년 %m월 %d일') + """</strong> 기준, 모니터링 중인 사이트에서 새로운 입찰 공고를 찾지 못했습니다.</p><p class="footer">본 메일은 자동화된 스크립트에 의해 발송되었습니다.</p></div></body>"""
    return html

def standardize_date(date_str):
    """다양한 형식의 날짜 문자열을 YYYY-MM-DD 형식으로 변환합니다."""
    if not date_str or not isinstance(date_str, str):
        return "N/A"
    try:
        # 정규식으로 'YYYY.MM.DD' 또는 'YYYY-MM-DD' 등의 기본 형식만 추출
        match = re.search(r'\d{4}[-.]\d{1,2}[-.]\d{1,2}', date_str)
        if match:
            return date_parse(match.group()).strftime('%Y-%m-%d')
        return date_str # 매칭되는 형식이 없으면 원본 반환
    except Exception:
        return date_str # 파싱 실패 시 원본 반환

# --- 4. 크롤링 전략별 핸들러 ---
def handle_css_crawl(target, session):
    """CSS 선택자 기반의 일반적인 웹사이트 크롤링을 처리합니다."""
    url = target.get('url')
    base_url = target.get('base_url', '')
    item_selector = target.get('item_selector')
    title_link_selector = target.get('title_link_selector')
    date_selector = target.get('date_selector')
    js_render = (target.get('js_render') or '').upper() == 'Y'

    company = target.get('company', 'N/A')

    if not all([url, item_selector, title_link_selector]):
        print(f"🟡 경고: '{company}'의 url, item_selector 또는 title_link_selector가 비어있어 건너뜁니다.")
        return []

    try:
        headers = {
            'User-Agent': (
                'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                'AppleWebKit/537.36 (KHTML, like Gecko) '
                'Chrome/91.0.4472.124 Safari/537.36'
            )
        }
        
        response = session.get(url, headers=headers, timeout=20)
        response.raise_for_status()

        # --- 인코딩 보정 ---
        if 'heungkuklife' in url:
            response.encoding = 'EUC-KR'
            print(f"ℹ️ '{company}' 사이트의 인코딩을 EUC-KR로 설정했습니다.")
        elif 'pikk.co.kr' in url:
            response.encoding = 'utf-8'
            print(f"ℹ️ '{company}' 사이트의 인코딩을 UTF-8로 설정했습니다.")

        if js_render:
            print(f"ℹ️ '{company}' 사이트는 JavaScript 렌더링을 사용합니다.")
            response.html.render(sleep=3, timeout=20)

        if js_render and hasattr(response, "html") and getattr(response.html, "html", None):
            html_source = response.html.html
        else:
            html_source = response.text

        soup = BeautifulSoup(html_source, 'html.parser')
        items = soup.select(item_selector)

        if not items:
            print(f"🟡 경고: '{company}'에서 '{item_selector}' 선택자에 해당하는 항목을 찾지 못했습니다.")
            return []

        announcements = []
        for item in items:
            # 1차 시도: 정의된 title_link_selector로 찾기
            title_element = None
            if title_link_selector:
                title_element = item.select_one(title_link_selector)

            # 2차 fallback: item 자체가 <a href="..."> 인 경우
            if not title_element:
                if item.name == 'a' and item.get('href'):
                    title_element = item
                else:
                    # 3차 fallback: item 내부의 첫 번째 <a href=...> 사용
                    link_tag = item.find('a', href=True)
                    if link_tag:
                        title_element = link_tag

            if not title_element:
                continue

            href = (title_element.get('href') or '').strip()

            # --- 제목 추출 ---
            if 'pikk.co.kr' in url:
                title_tag = item.find('h3')
                if title_tag:
                    title = title_tag.get_text(strip=True)
                else:
                    title = title_element.get_text(strip=True)
            else:
                title = title_element.get_text(strip=True)

            # --- [수정된 부분] 링크 추출 로직 완성형 (data-key, href=javascript, onclick 모두 지원) ---
            # href가 없거나, javascript, 또는 # 링크인 경우 대체 속성 확인
            if not href or 'javascript' in href.lower() or href == '#':
                link_format = target.get('link_format')
                
                # 1. data-key 속성 확인 (신한라이프 등)
                data_key = title_element.attrs.get('data-key')
                
                if data_key and link_format:
                    href = link_format.replace('{id}', str(data_key).strip())
                
                # 2. 자바스크립트(onclick 또는 href)에서 ID 추출 (삼양그룹, 미래에셋 등)
                else:
                    # onclick 값을 먼저 가져오고, 없으면 href 값이 'javascript:'로 시작하는지 확인
                    js_code = (title_element.get('onclick') or '').strip()
                    if not js_code and href.lower().startswith('javascript:'):
                        js_code = href
                    
                    # 정규식으로 괄호 안의 숫자나 문자열 추출 (예: goView(11453) -> 11453)
                    if js_code:
                        match = re.search(r"[(']([^()']+)[')]", js_code)
                        if match:
                            link_part = match.group(1)
                            if link_format:
                                href = link_format.replace('{id}', link_part)

            # 날짜 파싱
            post_date = "N/A"
            if date_selector:
                date_element = item.select_one(date_selector)
                if date_element:
                    post_date = standardize_date(date_element.get_text(strip=True))

            # 상대경로 링크를 절대경로로 변환
            if href and not href.startswith('http') and not href.startswith('javascript'):
                href = (base_url or url).rstrip('/') + '/' + href.lstrip('/')

            if href and title:
                announcements.append({
                    "title": title,
                    "href": href,
                    "date": post_date
                })

        return announcements

    except requests.exceptions.Timeout:
        print(f"❌ '{company}' 사이트 접속 시간 초과.")
        return []
    except requests.RequestException as e:
        print(f"❌ '{company}' 사이트 접속 실패: {e}")
        return []
    except Exception as e:
        print(f"❌ '{company}' 처리 중 알 수 없는 오류: {e}")
        return []

def handle_api_crawl(target, session):
    """JSON API 기반의 크롤링을 처리합니다."""
    api_url = target.get('api_url')
    method = (target.get('api_method') or 'GET').upper()
    
    def get_path(path_str):
        return path_str.split('.') if path_str else []

    item_path = get_path(target.get('json_item_path'))
    title_path = get_path(target.get('json_title_path'))
    link_id_path = get_path(target.get('json_link_id_path'))
    date_path = get_path(target.get('json_date_path'))
    link_format = target.get('link_format')

    if not all([api_url, item_path, title_path, link_id_path, link_format]):
        print(f"🟡 경고: '{target.get('company')}'의 API 설정이 부족하여 건너뜁니다.")
        return []

    try:
        payload_str = target.get('api_payload')
        payload = json.loads(payload_str) if payload_str else None

        if method == 'POST':
            response = session.post(api_url, json=payload, data=None if payload else target.get('api_form_data'))
        else:
            response = session.get(api_url, params=payload)
        
        response.raise_for_status()
        data = response.json()

        items = data
        for key in item_path:
            items = items.get(key, [])
            if not isinstance(items, list):
                print(f"🟡 경고: '{target.get('company')}'의 json_item_path '{'.'.join(item_path)}'가 리스트가 아닙니다.")
                return []
        
        announcements = []
        for item in items:
            title = item
            for key in title_path: title = title.get(key)
            
            link_id = item
            for key in link_id_path: link_id = link_id.get(key)
            
            post_date = "N/A"
            if date_path:
                date_val = item
                for key in date_path: date_val = date_val.get(key)
                if isinstance(date_val, int) and date_val > 10000000000:
                    post_date = datetime.fromtimestamp(date_val / 1000).strftime('%Y-%m-%d')
                else:
                    post_date = standardize_date(str(date_val))

            if title and link_id:
                href = link_format.replace('{id}', str(link_id))
                announcements.append({"title": str(title), "href": href, "date": post_date})
        
        return announcements

    except requests.RequestException as e:
        print(f"❌ '{target.get('company')}' API 접속 실패: {e}")
        return []
    except json.JSONDecodeError:
        print(f"❌ '{target.get('company')}' API 응답이 JSON 형식이 아닙니다.")
        return []
    except Exception as e:
        print(f"❌ '{target.get('company')}' API 처리 중 오류 발생: {e}")
        return []

# --- 5. 메인 실행 로직 ---
def crawl_site(target, processed_links, session):
    """크롤링 대상을 분기하여 실행하고 신규 공고를 반환합니다."""
    company = target.get('company', 'N/A')
    crawl_type = (target.get('crawl_type') or 'CSS').upper()

    print(f"\n--- '{company}' ({crawl_type}) 사이트 크롤링 시작 ---")
    
    new_announcements = []
    if crawl_type == 'CSS':
        results = handle_css_crawl(target, session)
    elif crawl_type == 'API':
        results = handle_api_crawl(target, session)
    else:
        print(f"🟡 경고: '{company}'의 crawl_type '{crawl_type}'은 지원되지 않는 형식입니다.")
        results = []

    if results:
        for ann in results:
            ann['company'] = company
            if ann['href'] and ann['href'] not in processed_links:
                print(f"🚀 새로운 공고 발견: [{company}] {ann['title']} (공고일: {ann['date']})")
                new_announcements.append(ann)
                save_processed_link(ann['href'])
                processed_links.add(ann['href'])
    
    if not new_announcements:
        print(f"ℹ️ '{company}'에서 새로운 공고를 찾지 못했습니다.")
        
    return new_announcements

def main():
    print("="*60 + f"\n입찰 공고 크롤러 (v4.2 - 공고 없을 시에도 메일 발송)를 시작합니다.\n" + "="*60)
    
    access_token = get_ms_graph_access_token()
    if not access_token: return

    settings_data = get_excel_data(access_token, "Settings")
    settings = {item['Setting']: item['Value'] for item in settings_data if item.get('Setting') and item.get('Value')}
    
    # 워크플로우 타입에 따라 수신자 이메일 목록 결정
    workflow_type = os.environ.get('WORKFLOW_TYPE', 'DEFAULT')
    receiver_emails = []
    
    developer_email = settings.get('Developer Email')
    receiver_email = settings.get('Receiver Email')

    if workflow_type == 'TEST':
        if developer_email:
            receiver_emails.append(developer_email)
        print("ℹ️ 테스트 모드로 실행. 개발자에게만 이메일이 발송됩니다.")
    else: # DEFAULT (일반 스케줄 실행)
        if receiver_email:
            receiver_emails.append(receiver_email)
        if developer_email:
            receiver_emails.append(developer_email)
        print("ℹ️ 일반 모드로 실행. 모든 수신자에게 이메일이 발송됩니다.")
            
    targets = get_excel_data(access_token, "Crawl_Targets")
    
    if not targets or not receiver_emails:
        print("❌ 크롤링에 필요한 설정 정보(대상 또는 수신 이메일)가 부족하여 작업을 종료합니다.")
        return
        
    processed_links = load_processed_links()
    all_new_announcements = []
    
    session = HTMLSession()

    for target in targets:
        if not target.get('company'): 
            continue
        try:
            new_finds = crawl_site(target, processed_links, session)
            if new_finds:
                all_new_announcements.extend(new_finds)
        except Exception as e:
            print(f"🚨 '{target.get('company')}' 크롤링 중 치명적 오류 발생: {e}")
        time.sleep(1)

    print("\n" + "="*25 + " 모든 사이트 크롤링 완료 " + "="*25)

    # --- [수정된 부분] ---
    if all_new_announcements:
        all_new_announcements.sort(key=lambda x: (x.get('date', '0000-00-00'), x.get('company')), reverse=True)
        
        save_announcements_to_excel(access_token, all_new_announcements)
        count = len(all_new_announcements)
        subject = f"[신규 공고 알림] {count}개의 새로운 공고가 수집되었습니다."
        body = generate_summary_email_body(all_new_announcements)
        send_email(subject, body, receiver_emails)
    else:
        # 신규 공고가 없을 때도 이메일을 발송하도록 변경
        print("\nℹ️ 모든 사이트에서 새로운 공고를 찾지 못했습니다. 결과 이메일을 발송합니다.")
        kst = timezone(timedelta(hours=9))
        today_str = datetime.now(kst).strftime('%Y-%m-%d')
        subject = f"[입찰 공고 알림] {today_str} 신규 공고 없음"
        body = generate_no_new_announcements_email_body() # 새로 추가한 함수 호출
        send_email(subject, body, receiver_emails)
        
    print("\n" + "="*30 + " 작업 종료 " + "="*30)

if __name__ == '__main__':
    main()
