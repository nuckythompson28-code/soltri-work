"""
SOLTRI-1 좌표 찾기 도우미
실행 후 마우스를 버튼 위에 올리고 F9 누르면 좌표 저장됨
ESC 누르면 종료
"""
import pyautogui, keyboard, time, json, os

coords = {}
labels = [
    ("login_pw",       "비밀번호 입력칸 위에 마우스 올리고 F9"),
    ("login_btn",      "로그인 버튼 위에 마우스 올리고 F9"),
    ("menu_bojo",      "보조기능 메뉴 위에 마우스 올리고 F9"),
    ("menu_prod",      "생산완료 항목 위에 마우스 올리고 F9"),
    ("date_start",     "시작날짜 입력칸 위에 마우스 올리고 F9"),
    ("date_end",       "끝날짜 입력칸 위에 마우스 올리고 F9"),
    ("btn_search",     "검색 버튼 위에 마우스 올리고 F9"),
    ("btn_excel",      "엑셀로 내보내기 버튼 위에 마우스 올리고 F9"),
]

print("=" * 50)
print("SOLTRI-1 좌표 세팅 도우미")
print("F9: 현재 마우스 위치 저장 / ESC: 종료")
print("=" * 50)

stop = False
for key, msg in labels:
    print(f"\n▶ {msg}")
    while True:
        if keyboard.is_pressed('f9'):
            x, y = pyautogui.position()
            coords[key] = [x, y]
            print(f"   저장됨: ({x}, {y})")
            time.sleep(0.5)
            break
        if keyboard.is_pressed('esc'):
            print("중단됨")
            stop = True
            break
        time.sleep(0.05)
    if stop:
        break

save_path = os.path.join(os.path.dirname(__file__), 'soltri_coords.json')
with open(save_path, 'w', encoding='utf-8') as f:
    json.dump(coords, f, ensure_ascii=False, indent=2)

print(f"\n✅ 좌표 저장 완료: {save_path}")
print(json.dumps(coords, ensure_ascii=False, indent=2))
