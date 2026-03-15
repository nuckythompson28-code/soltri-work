"""
SOLTRI-1 버튼 이미지 캡처 도구
각 버튼 위에 마우스를 올리고 F9를 누르면 해당 영역을 캡처합니다.
"""
import pyautogui, keyboard, json, os, sys
from PIL import ImageGrab

BASE = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.join(BASE, 'btn_images')
os.makedirs(IMG_DIR, exist_ok=True)

ITEMS = [
    ('login_pw',    '비밀번호 입력칸'),
    ('login_btn',   '로그인 버튼'),
    ('menu_bojo',   '보조기능 메뉴'),
    ('menu_prod',   '생산완료 버튼'),
    ('date_start',  '시작날짜 입력칸'),
    ('date_end',    '끝날짜 입력칸'),
    ('btn_search',  '검색 버튼'),
    ('btn_excel',   '엑셀내보내기 버튼'),
]

# 캡처 영역 크기 (마우스 중심 기준 좌우상하)
W_HALF = 60
H_HALF = 18

print("=" * 50)
print("SOLTRI-1 버튼 이미지 캡처")
print(f"  저장 위치: {IMG_DIR}")
print(f"  캡처 크기: {W_HALF*2} x {H_HALF*2} px (마우스 중심)")
print("=" * 50)
print()
print("각 항목에서 해당 버튼/칸 위에 마우스를 올리고 F9를 누르세요.")
print("ESC = 종료")
print()

for i, (key, desc) in enumerate(ITEMS, 1):
    print(f"[{i}/{len(ITEMS)}] {desc} ({key}) — 마우스를 올리고 F9", flush=True)

    while True:
        event = keyboard.read_event(suppress=False)
        if event.event_type != 'down':
            continue
        if event.name == 'esc':
            print("중단됨.")
            sys.exit(0)
        if event.name == 'f9':
            break

    x, y = pyautogui.position()
    left = max(0, x - W_HALF)
    top = max(0, y - H_HALF)
    right = x + W_HALF
    bottom = y + H_HALF

    img = ImageGrab.grab(bbox=(left, top, right, bottom))
    path = os.path.join(IMG_DIR, f'{key}.png')
    img.save(path)
    print(f"  ✅ 저장: {path}  (중심: {x},{y}  영역: {left},{top}~{right},{bottom})")
    print()

print("=" * 50)
print("✅ 모든 버튼 이미지 캡처 완료!")
print(f"   {IMG_DIR} 폴더를 확인하세요.")
print("=" * 50)
