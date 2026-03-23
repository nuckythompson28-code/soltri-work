"""
폐수배출시설 운영일지 - 로컬 서버
- Google Sheets 자동 연동 (서비스 계정)
- 정적 파일 서빙 (HTML, template.docx)
"""
import http.server, json, os, sys, traceback
from urllib.parse import urlparse, parse_qs

PORT = 8789
HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(HERE, 'config.json')
TEMPLATE_PATH = os.path.join(HERE, '폐수배출시설_운영일지_template.docx')
CREDS_PATH = "G:/내 드라이브/work/warehouse/gen-lang-client-0766779209-fd365fc3ce58.json"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
SHEET_TAB = '폐수배출'
HEADER = ['날짜','요일','쉬는날','1호기용수','2호기용수','1호기계량기','2호기계량기','필터교체','필터교체량','위탁량','확인서번호','처리업소명']

# ── Google Sheets ──
_sheet = None

def get_sheet():
    global _sheet
    if _sheet:
        return _sheet
    import gspread
    from google.oauth2.service_account import Credentials
    creds = Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES)
    gc = gspread.authorize(creds)

    cfg = load_config()
    sheet_id = cfg.get('sheet_id')

    # 기존 스프레드시트 사용 (없으면 frequency 프로젝트 시트 활용)
    if not sheet_id:
        sheet_id = '1MsmVKtz5NTxIIoj3efXYPLEhL3GaONW5LAlRNjKk7s0'

    sh = gc.open_by_key(sheet_id)
    save_config({'sheet_id': sheet_id, 'sheet_url': sh.url})
    print(f"[Sheet] 시트 연결: {sh.title}")

    try:
        ws = sh.worksheet(SHEET_TAB)
    except Exception:
        ws = sh.add_worksheet(SHEET_TAB, rows=400, cols=len(HEADER))
        ws.update([HEADER], 'A1')
        print(f"[Sheet] '{SHEET_TAB}' 탭 생성")

    _sheet = ws
    return ws


def load_config():
    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}

def save_config(updates):
    cfg = load_config()
    cfg.update(updates)
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def read_sheet_data(year, month):
    """시트에서 해당 월 데이터 읽기"""
    ws = get_sheet()
    all_data = ws.get_all_values()
    if len(all_data) <= 1:
        return []

    header = all_data[0]
    yy = year % 100
    # 여러 날짜 형식 지원: 26/02/01, 2026-02-01, 26-02-01
    prefixes = [
        f"{yy}/{month:02d}/",
        f"{year}-{month:02d}-",
        f"{yy}-{month:02d}-",
    ]
    rows = []
    for r in all_data[1:]:
        obj = {}
        for i, h in enumerate(header):
            obj[h] = r[i] if i < len(r) else ''
        ds = obj.get('날짜', '').strip()
        if any(ds.startswith(p) for p in prefixes):
            rows.append(obj)
    return rows


def _normalize_date(ds):
    """날짜 문자열을 YY/MM/DD 형식으로 통일"""
    ds = str(ds).strip()
    # 2026-01-01 → 26/01/01
    if '-' in ds:
        parts = ds.split('-')
        y = parts[0][-2:]  # 마지막 2자리
        return f"{y}/{parts[1].zfill(2)}/{parts[2].zfill(2)}"
    # 26/1/1 → 26/01/01
    if '/' in ds:
        parts = ds.split('/')
        return f"{parts[0]}/{parts[1].zfill(2)}/{parts[2].zfill(2)}"
    return ds


def write_sheet_data(rows):
    """시트에 데이터 쓰기 (날짜 기준 upsert, 정규화 매칭)"""
    ws = get_sheet()
    all_data = ws.get_all_values()

    # 기존 날짜(정규화) → 행번호 매핑
    date_map = {}
    if len(all_data) > 1:
        for i, r in enumerate(all_data[1:], start=2):
            if r and r[0]:
                date_map[_normalize_date(r[0])] = i
    elif len(all_data) == 0:
        ws.update([HEADER], 'A1')

    updated, added = 0, 0
    for row in rows:
        ds = row.get('날짜', '')
        if not ds:
            continue
        norm = _normalize_date(ds)
        values = [row.get(h, '') for h in HEADER]
        if norm in date_map:
            rn = date_map[norm]
            ws.update([values], f'A{rn}:L{rn}')
            updated += 1
        else:
            ws.append_row(values, value_input_option='RAW')
            date_map[norm] = len(all_data) + added + 1
            added += 1
    return updated, added


# ── HTTP Server ──
class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=HERE, **kw)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        qs = parse_qs(parsed.query)

        if path == '/':
            self.path = '/폐수운영일지_생성기.html'
            return super().do_GET()
        elif path == '/api/data':
            self._handle_read(qs)
        elif path == '/api/template':
            self._serve_template()
        elif path == '/api/config':
            self._json(load_config())
        elif path == '/favicon.ico':
            self.send_response(204); self.end_headers()
        else:
            super().do_GET()

    def do_POST(self):
        if urlparse(self.path).path == '/api/data':
            self._handle_write()
        else:
            self.send_error(404)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def _handle_read(self, qs):
        try:
            y = int(qs.get('year', [2026])[0])
            m = int(qs.get('month', [1])[0])
            data = read_sheet_data(y, m)
            print(f"[API] 읽기: {y}-{m:02d} → {len(data)}행")
            self._json({'success': True, 'data': data})
        except Exception as e:
            traceback.print_exc()
            self._json({'success': False, 'error': str(e)}, 500)

    def _handle_write(self):
        try:
            length = int(self.headers.get('Content-Length', 0))
            body = json.loads(self.rfile.read(length).decode('utf-8'))
            rows = body.get('rows', [])
            updated, added = write_sheet_data(rows)
            print(f"[API] 쓰기: {updated}건 수정, {added}건 추가")
            self._json({'success': True, 'updated': updated, 'added': added})
        except Exception as e:
            traceback.print_exc()
            self._json({'success': False, 'error': str(e)}, 500)

    def _serve_template(self):
        try:
            with open(TEMPLATE_PATH, 'rb') as f:
                data = f.read()
            self.send_response(200)
            self.send_header('Content-Type', 'application/octet-stream')
            self.send_header('Content-Length', str(len(data)))
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(data)
        except Exception as e:
            self._json({'error': str(e)}, 500)

    def _json(self, data, code=200):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):
        if '/api/' in str(args[0]):
            print(f"  {args[0]}")


if __name__ == '__main__':
    print("=" * 50)
    print("  폐수배출시설 운영일지 생성기 서버")
    print(f"  http://localhost:{PORT}")
    print("=" * 50)

    # 초기 시트 연결
    try:
        ws = get_sheet()
        cfg = load_config()
        print(f"  시트: {cfg.get('sheet_url', '?')}")
    except Exception as e:
        print(f"  시트 연결 실패: {e}")

    print("=" * 50)

    # 브라우저 자동 열기
    import webbrowser, threading
    threading.Timer(1.0, lambda: webbrowser.open(f'http://localhost:{PORT}')).start()

    server = http.server.HTTPServer(('', PORT), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n서버 종료")
        server.shutdown()
