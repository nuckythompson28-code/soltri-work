# SOLTRI-1 (주)쏠트리 창의 컨트롤 구조를 출력
from pywinauto.application import Application
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# WinForms 앱이므로 win32 backend로 연결
try:
    app = Application(backend='win32').connect(class_name='WindowsForms10.Window.8.app.0.1a0e24_r7_ad1')
    dlg = app.window(class_name='WindowsForms10.Window.8.app.0.1a0e24_r7_ad1')
    print(f"창 제목: {dlg.window_text()}")
    print(f"창 클래스: {dlg.class_name()}")
    print("\n=== win32 backend 컨트롤 트리 ===")
    dlg.print_control_identifiers(depth=5)
except Exception as e:
    print(f"win32 실패: {e}")

print("\n" + "="*60)

# UIA backend도 시도
try:
    app2 = Application(backend='uia').connect(title='(주)쏠트리')
    dlg2 = app2.window(title='(주)쏠트리')
    print(f"\n=== uia backend 컨트롤 트리 ===")
    dlg2.print_control_identifiers(depth=5)
except Exception as e:
    print(f"uia 실패: {e}")
