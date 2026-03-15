"""
SOLTRI-1 자동화 스크립트 (이미지 인식 방식)
실행 순서:
  1. capture_buttons.py 로 버튼 이미지 캡처 (최초 1회, 또는 이미지캡처.bat)
  2. soltri_config.json 에 비밀번호/경로 입력
  3. 이 파일 실행 (또는 자동실행.bat)

좌표 방식(soltri_coords.json)이 있으면 fallback으로 사용
"""
import pyautogui, pyperclip, time, json, os, sys
import pygetwindow as gw
import win32com.client
import threading, http.server, webbrowser
from datetime import date, datetime

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.15

# ── 경로 설정 ──────────────────────────────────────────────
BASE = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.join(BASE, 'btn_images')
COORDS_FILE = os.path.join(BASE, 'soltri_coords.json')
CONFIG_FILE = os.path.join(BASE, 'soltri_config.json')

# ── 설정 로드 ──────────────────────────────────────────────
if not os.path.exists(CONFIG_FILE):
    print("❌ soltri_config.json 이 없습니다.")
    sys.exit(1)

with open(CONFIG_FILE, encoding='utf-8') as f:
    CFG = json.load(f)

# 좌표 fallback 로드
COORDS = {}
if os.path.exists(COORDS_FILE):
    with open(COORDS_FILE, encoding='utf-8') as f:
        COORDS = json.load(f)

USE_IMG = os.path.isdir(IMG_DIR) and len(os.listdir(IMG_DIR)) >= 6
if USE_IMG:
    print("🖼️  이미지 인식 모드")
else:
    if not COORDS:
        print("❌ btn_images 폴더도 없고 soltri_coords.json도 없습니다.")
        print("   이미지캡처.bat 또는 좌표세팅.bat을 먼저 실행하세요.")
        sys.exit(1)
    print("📌 좌표 모드 (이미지 없음, fallback)")

USER_PW   = CFG.get('pw', '')
EXE_PATH  = CFG.get('exe_path', '')
SAVE_DIR  = CFG.get('save_folder', BASE)
DATE_FROM = CFG.get('date_start', '2026-01-01')
DATE_TO   = date.today().strftime('%Y-%m-%d')

_now = datetime.now().strftime('%H%M')
SAVE_FILENAME = f"CNC_생산현황_{DATE_FROM.replace('-','')}_{DATE_TO.replace('-','')}_{_now}.xlsx"
SAVE_PATH     = os.path.normpath(os.path.join(SAVE_DIR, SAVE_FILENAME))

# ── 유틸 함수 ──────────────────────────────────────────────
def find_and_click(key, delay=0.5, timeout=10, confidence=0.8):
    """이미지 인식으로 버튼 찾아서 클릭. 실패시 좌표 fallback."""
    if USE_IMG:
        img_path = os.path.join(IMG_DIR, f'{key}.png')
        if os.path.exists(img_path):
            start = time.time()
            while time.time() - start < timeout:
                try:
                    loc = pyautogui.locateCenterOnScreen(img_path, confidence=confidence)
                    if loc:
                        pyautogui.click(loc)
                        time.sleep(delay)
                        return True
                except pyautogui.ImageNotFoundException:
                    pass
                except Exception:
                    pass
                time.sleep(0.3)
            print(f"    ⚠️ 이미지 못찾음: {key} → 좌표 fallback 시도")

    # 좌표 fallback
    if key in COORDS:
        x, y = COORDS[key]
        pyautogui.click(x, y)
        time.sleep(delay)
        return True

    print(f"    ❌ {key}: 이미지도 좌표도 없음!")
    return False

def type_text(text, clear=True):
    if clear:
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.1)
    pyperclip.copy(str(text))
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.2)

def wait(msg, sec=1.0):
    print(f"  ⏳ {msg} ({sec}초 대기)")
    time.sleep(sec)

# ── 메인 자동화 ────────────────────────────────────────────
print("=" * 50)
print("SOLTRI-1 자동화 시작")
print(f"  기간: {DATE_FROM} ~ {DATE_TO}")
print(f"  저장: {SAVE_PATH}")
print("=" * 50)
print("⚠️  마우스를 화면 왼쪽 상단으로 이동하면 즉시 중단됩니다")
print()

for i in range(3, 0, -1):
    print(f"  {i}초 후 시작...")
    time.sleep(1)
print()

# ── STEP 1: 프로그램 실행 ──
print("[1/7] SOLTRI-1 실행")
already_open = '(주)쏠트리' in gw.getAllTitles()

if already_open:
    print("  ℹ️  이미 실행 중")
elif EXE_PATH and os.path.exists(EXE_PATH):
    os.startfile(EXE_PATH)
    wait("프로그램 로딩 대기", 5.0)
else:
    print("  ⚠️  exe_path 미설정")

# ── STEP 2: 로그인 ──
print("[2/7] 로그인")
# 로그인 버튼이 보일 때까지 대기 (화면 준비 확인)
if USE_IMG:
    img_path = os.path.join(IMG_DIR, 'login_btn.png')
    if os.path.exists(img_path):
        print("  ⏳ 로그인 화면 대기 중...", end='', flush=True)
        for _ in range(30):
            try:
                loc = pyautogui.locateCenterOnScreen(img_path, confidence=0.8)
                if loc:
                    print(" 확인!")
                    break
            except Exception:
                pass
            time.sleep(0.5)
            print('.', end='', flush=True)
        else:
            print("\n  ⚠️ 로그인 화면 감지 못함, 계속 진행")
        time.sleep(0.3)
# 프로그램 실행 시 비밀번호칸에 커서가 이미 있음
type_text(USER_PW, clear=False)
find_and_click('login_btn', 1.0)
wait("로그인 처리", 2.0)

# ── STEP 3: 보조기능 메뉴 ──
print("[3/7] 보조기능 메뉴 클릭")
find_and_click('menu_bojo', 0.5)
wait("메뉴 열림 대기", 1.0)

# ── STEP 4: 생산완료 클릭 ──
print("[4/7] 생산완료 클릭")
find_and_click('menu_prod', 0.5)
wait("화면 로딩", 10.0)

# ── STEP 5: 날짜 입력 ──
print(f"[5/7] 날짜 입력  {DATE_FROM} ~ {DATE_TO}")
find_and_click('date_start', 0.3)
type_text(DATE_FROM)
pyautogui.press('tab')
time.sleep(0.2)

find_and_click('date_end', 0.3)
type_text(DATE_TO)
pyautogui.press('tab')
time.sleep(0.2)

# ── STEP 6: 검색 ──
print("[6/7] 검색")
find_and_click('btn_search', 0.5)
wait("검색 결과 로딩", 3.0)

# ── STEP 7: 엑셀 내보내기 ──
print("[7/7] 엑셀로 내보내기")
find_and_click('btn_excel', 1.0)

# Excel이 열릴 때까지 대기 (최대 15초)
print("  ⏳ Excel 로딩 대기 중...", end='', flush=True)
for _ in range(30):
    time.sleep(0.5)
    wins = gw.getAllTitles()
    if any('Excel' in w or 'excel' in w for w in wins):
        print(" 확인!")
        break
    print('.', end='', flush=True)
else:
    print("\n  ⚠️  Excel 창을 찾지 못했지만 계속 진행합니다")
time.sleep(1.5)

# Excel 창 앞으로 가져오기
excel_wins = [w for w in gw.getAllWindows() if 'Excel' in w.title or 'excel' in w.title]
if excel_wins:
    try:
        excel_wins[0].activate()
        time.sleep(0.8)
    except Exception:
        pass

# Excel COM으로 직접 xlsx 저장 (다이얼로그 없이)
print("  Excel: COM으로 xlsx 저장 중...")
try:
    xl = win32com.client.GetActiveObject("Excel.Application")
    wb = xl.ActiveWorkbook
    xl.DisplayAlerts = False
    wb.SaveAs(Filename=str(SAVE_PATH), FileFormat=51)
    wb.Close(False)
    xl.DisplayAlerts = True
    print("  ✅ xlsx 저장 완료")
except Exception as e:
    print(f"  ⚠️  COM 저장 실패: {e}")
    print("  → Excel 창을 수동으로 닫아주세요")

print()
print("=" * 50)
if os.path.exists(SAVE_PATH):
    size = os.path.getsize(SAVE_PATH)
    print(f"✅ 완료!  저장됨: {SAVE_PATH}  ({size:,} bytes)")

    # ── 로컬 서버 시작 + 브라우저 오픈 ──
    PORT = 18765
    HTML_NAME = 'CNC_일별생산현황_v3.html'

    class SilentHandler(http.server.SimpleHTTPRequestHandler):
        def log_message(self, *a): pass
        def end_headers(self):
            self.send_header('Cache-Control', 'no-cache')
            super().end_headers()

    os.chdir(SAVE_DIR)
    server = http.server.HTTPServer(('localhost', PORT), SilentHandler)
    t = threading.Thread(target=server.serve_forever, daemon=True)
    t.start()
    print(f"🌐 로컬 서버 시작: http://localhost:{PORT}")

    for key in ('ctrl', 'alt', 'shift', 'win'):
        pyautogui.keyUp(key)
    time.sleep(0.5)

    url = f"http://localhost:{PORT}/{HTML_NAME}?autoload={SAVE_FILENAME}"
    webbrowser.open(url)
    print(f"🔗 브라우저 오픈: {url}")
    print("(브라우저 닫아도 이 창은 계속 실행 중 — 창 닫으면 서버도 종료)")
    input("\n브라우저 확인 후 Enter 누르면 종료...")
    server.shutdown()
else:
    print(f"⚠️  저장 확인 불가: {SAVE_PATH}")
print("=" * 50)
