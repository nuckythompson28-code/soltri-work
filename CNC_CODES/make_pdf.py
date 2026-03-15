from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm, cm
from reportlab.lib.colors import HexColor, white, black
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, KeepTogether, HRFlowable
)
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

# ─── 폰트 등록 ───
pdfmetrics.registerFont(TTFont('Malgun', 'C:/Windows/Fonts/malgun.ttf'))
pdfmetrics.registerFont(TTFont('MalgunBd', 'C:/Windows/Fonts/malgunbd.ttf'))

# ─── 색상 ───
NAVY      = HexColor('#1B2A4A')
DARK_BLUE = HexColor('#2C3E6B')
ACCENT    = HexColor('#3B82F6')
TEAL      = HexColor('#0D9488')
RED       = HexColor('#EF4444')
ORANGE    = HexColor('#F97316')
GREEN     = HexColor('#10B981')
GRAY      = HexColor('#64748B')
LIGHT_BG  = HexColor('#F0F4F8')
BLUE_BG   = HexColor('#E3F2FD')
RED_BG    = HexColor('#FFEBEE')
GREEN_BG  = HexColor('#E8F5E9')
YELLOW_BG = HexColor('#FFFACD')

# ─── 스타일 (큰 글씨 — 노안 배려) ───
styles = {
    'title': ParagraphStyle('title', fontName='MalgunBd', fontSize=30, leading=40, textColor=NAVY, spaceAfter=5*mm),
    'subtitle': ParagraphStyle('subtitle', fontName='Malgun', fontSize=16, leading=22, textColor=GRAY, spaceAfter=8*mm),
    'h1': ParagraphStyle('h1', fontName='MalgunBd', fontSize=24, leading=32, textColor=NAVY, spaceBefore=10*mm, spaceAfter=5*mm),
    'h2': ParagraphStyle('h2', fontName='MalgunBd', fontSize=19, leading=26, textColor=DARK_BLUE, spaceBefore=7*mm, spaceAfter=4*mm),
    'h3': ParagraphStyle('h3', fontName='MalgunBd', fontSize=16, leading=22, textColor=ACCENT, spaceBefore=5*mm, spaceAfter=3*mm),
    'body': ParagraphStyle('body', fontName='Malgun', fontSize=14, leading=21, textColor=black, spaceAfter=3*mm),
    'body_bold': ParagraphStyle('body_bold', fontName='MalgunBd', fontSize=14, leading=21, textColor=black, spaceAfter=3*mm),
    'code': ParagraphStyle('code', fontName='Courier', fontSize=12, leading=17, textColor=HexColor('#1E293B'), spaceAfter=2*mm, leftIndent=10*mm),
    'note': ParagraphStyle('note', fontName='Malgun', fontSize=13, leading=19, textColor=DARK_BLUE, spaceAfter=3*mm, leftIndent=5*mm),
    'warn': ParagraphStyle('warn', fontName='MalgunBd', fontSize=14, leading=20, textColor=RED, spaceAfter=3*mm, leftIndent=5*mm),
    'small': ParagraphStyle('small', fontName='Malgun', fontSize=12, leading=17, textColor=GRAY, spaceAfter=2*mm),
    'toc': ParagraphStyle('toc', fontName='Malgun', fontSize=16, leading=26, textColor=black, leftIndent=8*mm, spaceAfter=2*mm),
    'center': ParagraphStyle('center', fontName='MalgunBd', fontSize=15, leading=22, textColor=DARK_BLUE, alignment=TA_CENTER, spaceAfter=4*mm),
    'center_big': ParagraphStyle('center_big', fontName='MalgunBd', fontSize=18, leading=26, textColor=NAVY, alignment=TA_CENTER, spaceAfter=5*mm),
}

def B(text):
    return f'<b>{text}</b>'

def C(text, color):
    return f'<font color="{color}">{text}</font>'

def make_table(data, col_widths=None, header_color=NAVY):
    """Create a styled table (큰 글씨)"""
    t = Table(data, colWidths=col_widths, repeatRows=1)
    style_cmds = [
        ('FONTNAME', (0,0), (-1,0), 'MalgunBd'),
        ('FONTNAME', (0,1), (-1,-1), 'Malgun'),
        ('FONTSIZE', (0,0), (-1,-1), 12),
        ('LEADING', (0,0), (-1,-1), 17),
        ('BACKGROUND', (0,0), (-1,0), header_color),
        ('TEXTCOLOR', (0,0), (-1,0), white),
        ('TEXTCOLOR', (0,1), (-1,-1), black),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.7, HexColor('#CBD5E1')),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]
    for i in range(1, len(data)):
        if i % 2 == 0:
            style_cmds.append(('BACKGROUND', (0,i), (-1,i), LIGHT_BG))
        else:
            style_cmds.append(('BACKGROUND', (0,i), (-1,i), white))
    t.setStyle(TableStyle(style_cmds))
    return t

def info_box(text, bg_color=BLUE_BG, text_color=DARK_BLUE):
    """Create an info box (큰 글씨)"""
    p = Paragraph(text, ParagraphStyle('box', fontName='MalgunBd', fontSize=14, leading=21, textColor=text_color))
    t = Table([[p]], colWidths=[170*mm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), bg_color),
        ('ROUNDEDCORNERS', [3,3,3,3]),
        ('TOPPADDING', (0,0), (-1,-1), 10),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ('LEFTPADDING', (0,0), (-1,-1), 12),
        ('RIGHTPADDING', (0,0), (-1,-1), 12),
    ]))
    return t

def section_divider():
    return HRFlowable(width="100%", thickness=1, color=HexColor('#E2E8F0'), spaceAfter=3*mm, spaceBefore=3*mm)


# ─── PDF 생성 ───
doc = SimpleDocTemplate(
    r"C:\Users\admin\Desktop\work\CNC_CODES\O0852_교육자료_v2.pdf",
    pagesize=A4,
    topMargin=20*mm, bottomMargin=20*mm,
    leftMargin=18*mm, rightMargin=18*mm,
    title="O0852 CNC 매크로 프로그램 교육자료",
    author="CNC Education"
)

story = []
pw = 174*mm  # page width minus margins

# ═══════════════════════════════════════════
# 표지
# ═══════════════════════════════════════════
story.append(Spacer(1, 40*mm))

cover_title = Paragraph("O0852 프로그램 완전 분석", ParagraphStyle('ct', fontName='MalgunBd', fontSize=38, leading=48, textColor=NAVY, alignment=TA_CENTER))
story.append(cover_title)
story.append(Spacer(1, 5*mm))

cover_sub = Paragraph("CNC 매크로 프로그램 교육자료", ParagraphStyle('cs', fontName='MalgunBd', fontSize=22, leading=30, textColor=ACCENT, alignment=TA_CENTER))
story.append(cover_sub)
story.append(Spacer(1, 3*mm))

story.append(HRFlowable(width="60%", thickness=2, color=ACCENT, spaceAfter=8*mm, spaceBefore=5*mm))

cover_desc = Paragraph("링형 부품 자동 연속 가공 시스템  |  FANUC 매크로", ParagraphStyle('cd', fontName='Malgun', fontSize=16, leading=22, textColor=GRAY, alignment=TA_CENTER))
story.append(cover_desc)
story.append(Spacer(1, 20*mm))

# Cover feature boxes
cover_items = [
    ["매크로 자동화", "변수 입력만으로 소재/치수/조건 설정"],
    ["연속 가공", "면삭→스텝→챔퍼→절단 사이클 자동 반복"],
    ["오토링크", "소재 부족 시 자동 인출 후 재가공"],
    ["안전 검증", "30개 이상 알람으로 입력 오류 사전 차단"],
]
for title, desc in cover_items:
    p = Paragraph(f'<b>{title}</b>  —  {desc}', ParagraphStyle('ci', fontName='Malgun', fontSize=15, leading=24, textColor=DARK_BLUE, alignment=TA_CENTER))
    story.append(p)

story.append(PageBreak())

# ═══════════════════════════════════════════
# 목차
# ═══════════════════════════════════════════
story.append(Paragraph("목차", styles['title']))
story.append(HRFlowable(width="100%", thickness=2, color=ACCENT, spaceAfter=8*mm))

toc_items = [
    ("01", "전체 요약", "프로그램이 하는 일 한눈에 보기"),
    ("02", "프로그램 구조", "4개 서브프로그램 관계도"),
    ("03", "입력 파라미터", "작업자가 수정하는 변수 (#101~#123)"),
    ("04", "소재별 이송속도", "CN / RS / CM 자동 설정"),
    ("05", "고급 설정 & 안전체크", "시스템 변수 및 검증 로직"),
    ("06", "가공 사이클 상세", "면삭 → 스텝 → 챔퍼 → 절단"),
    ("07", "오토링크 & 잔재 가공", "자동 소재 인출 시스템"),
    ("08", "알람 코드표", "에러 번호별 원인과 대처"),
    ("09", "기계별 M코드 매핑", "9대 기계 설정표"),
    ("10", "작업자 주의사항", "안전 수칙 및 체크리스트"),
]
for num, title, desc in toc_items:
    story.append(Paragraph(
        f'{C(num, ACCENT)}  <b>{title}</b>  <font color="{GRAY}">— {desc}</font>',
        styles['toc']
    ))

story.append(PageBreak())

# ═══════════════════════════════════════════
# 1. 전체 요약
# ═══════════════════════════════════════════
story.append(Paragraph("1. 전체 요약", styles['title']))
story.append(HRFlowable(width="100%", thickness=2, color=ACCENT, spaceAfter=5*mm))

story.append(info_box(
    "O0852는 CNC 선반에서 파이프 소재를 자동으로 연속 가공하여<br/>"
    "링형 부품(부싱/베어링류)을 대량 생산하는 FANUC 매크로 프로그램입니다."
))
story.append(Spacer(1, 5*mm))

story.append(Paragraph("현재 설정값 기준 가공 요약", styles['h3']))
summary_data = [
    ["항목", "소재", "완성품", "생산"],
    ["규격", "CN, OD70 × ID56\n길이 543mm", "OD64.9 × ID58.3\n길이 7.823mm", "1개당 9.843mm 소비\n1회 인출당 약 6개"],
]
story.append(make_table(summary_data, col_widths=[25*mm, 50*mm, 50*mm, 49*mm]))
story.append(Spacer(1, 5*mm))

story.append(Paragraph("핵심 특징", styles['h3']))
features = [
    [C("매크로 자동화", ACCENT), "변수 입력만으로 소재, 치수, 절삭 조건을 설정. 코드 수정 없이 다품종 대응"],
    [C("연속 가공", TEAL), "면삭 → 스텝절삭 → 챔퍼 → 절단 사이클을 자동으로 반복 실행"],
    [C("오토링크", ORANGE), "소재가 부족해지면 자동으로 소재를 인출(PULL)하여 가공을 이어감"],
    [C("안전 검증", RED), "30개 이상의 알람 체크로 입력 오류, 치수 모순, 안전 문제를 사전에 차단"],
]
for feat in features:
    story.append(Paragraph(f'<b>{feat[0]}</b>: {feat[1]}', styles['body']))

story.append(PageBreak())

# ═══════════════════════════════════════════
# 2. 프로그램 구조
# ═══════════════════════════════════════════
story.append(Paragraph("2. 프로그램 구조", styles['title']))
story.append(HRFlowable(width="100%", thickness=2, color=ACCENT, spaceAfter=5*mm))
story.append(Paragraph("4개의 서브프로그램이 순차적으로 호출됩니다.", styles['body']))
story.append(Spacer(1, 3*mm))

struct_data = [
    ["프로그램", "이름", "역할", "호출 방법"],
    ["O0852", "메인 셋업", "파라미터 입력, 입력값 검증, 기계 설정", "직접 실행"],
    ["O9001", "메인 로직", "길이/개수 계산, 알람 체크, 가공 시작", "M98 P9001"],
    ["O9002", "메인 가공", "면삭→스텝→챔퍼→절단 반복, 오토링크", "M98 P9002"],
    ["O9003", "잔재 가공", "남은 소재 인출, 척 재설정, 최종 가공", "M98 P9003"],
]
story.append(make_table(struct_data, col_widths=[25*mm, 25*mm, 75*mm, 25*mm]))
story.append(Spacer(1, 3*mm))

story.append(Paragraph("호출 흐름도", styles['h3']))
story.append(Paragraph(
    "O0852 (셋업)  →  O9001 (계산/검증)  →  O9002 (가공 사이클)  →  O9003 (잔재)",
    styles['center']
))
story.append(Spacer(1, 3*mm))

story.append(Paragraph("사용 공구", styles['h3']))
tool_data = [
    ["공구", "용도", "주요 공정", "비고"],
    ["T01", "보링바 (내경 가공)", "황삭, 정삭, 챔퍼", "메인 가공 공구"],
    ["T02", "절단 바이트", "면삭, 절단", "면삭 + 파팅"],
    ["T03", "오토링크용", "소재 인출", "클램프/언클램프"],
]
story.append(make_table(tool_data, col_widths=[25*mm, 45*mm, 45*mm, 45*mm]))

story.append(PageBreak())

# ═══════════════════════════════════════════
# 3. 입력 파라미터
# ═══════════════════════════════════════════
story.append(Paragraph("3. 입력 파라미터", styles['title']))
story.append(HRFlowable(width="100%", thickness=2, color=ACCENT, spaceAfter=5*mm))
story.append(info_box("작업자가 수정하는 영역입니다 (Line 4~29, #101~#123만 수정)", GREEN_BG, GREEN))
story.append(Spacer(1, 5*mm))

# 3-1 기본 설정
story.append(Paragraph("3-1. 기본 설정 (#101 ~ #108)", styles['h2']))
data = [
    ["변수", "설명", "현재값", "입력 규칙"],
    ["#101", "소재 길이 끝자리", "43", "0 ~ 99 범위"],
    ["#102", "소재 길이 백자리", "500", "0, 100, 200, 300, 400, 500만 가능"],
    ["#103", "척 길이 (물림 길이)", "75", "mm 단위"],
    ["#104", "초기 면삭량", "1", "최대 40까지"],
    ["#105", "소재 종류", "1", "1=CN, 2=RS, 3=CM"],
    ["#106", "가공 유형", "3", "3=챔퍼, 4=챔퍼(확장)"],
    ["#107", "황삭 ON/OFF", "0", "0=OFF, 1=ON"],
    ["#108", "단품 모드", "1", "1=한개씩 가공"],
]
story.append(make_table(data, col_widths=[22*mm, 42*mm, 22*mm, 70*mm]))
story.append(Spacer(1, 3*mm))
story.append(info_box(
    "소재 길이 계산: #101 + #102 = 전체 길이  →  예: 500 + 43 = 543mm",
    YELLOW_BG, black
))

story.append(Spacer(1, 5*mm))

# 3-2 치수 데이터
story.append(Paragraph("3-2. 치수 데이터 (#109 ~ #118)", styles['h2']))
data = [
    ["변수", "설명", "현재값", "단위"],
    ["#109", "소재 외경 (RAW OD)", "70", "mm"],
    ["#110", "소재 내경 (RAW ID)", "56", "mm"],
    ["#111", "완성 외경 (FIN OD)", "64.90", "mm"],
    ["#112", "완성 내경 (FIN ID)", "58.30", "mm"],
    ["#113", "완성 길이 (FIN LENGTH)", "7.823", "mm"],
    ["#114", "절단 바이트 폭 (TIP WIDTH)", "2.02", "mm"],
    ["#115", "내경 R (ID RADIUS)", "0.34", "mm"],
    ["#116", "외경 R (OD RADIUS)", "0.36", "mm"],
    ["#117", "내경 챔퍼 (ID CHAMFER)", "0.37", "mm"],
    ["#118", "외경 챔퍼 (OD CHAMFER)", "0.60", "mm"],
    ["#119", "길이>40mm 확인 플래그", "0", "40초과 시 1 필수"],
]
story.append(make_table(data, col_widths=[22*mm, 55*mm, 22*mm, 55*mm]))

# 3-3 절삭 조건
story.append(Paragraph("3-3. 절삭 조건 (#120 ~ #123)", styles['h2']))
data = [
    ["변수", "설명", "현재값", "비고"],
    ["#120", "T01 공구 RPM", "1700", "보링바 회전수"],
    ["#121", "T02 공구 RPM", "1700", "절단 바이트 회전수"],
    ["#122", "인출 거리 (PULL DIST)", "1", "오토링크 시 소재 인출 거리"],
    ["#123", "여유값 (MARGIN)", "0.2", "0~3 범위, 가공 마진"],
]
story.append(make_table(data, col_widths=[22*mm, 45*mm, 22*mm, 65*mm]))

# ═══════════════════════════════════════════
# 4. 소재별 이송속도
# ═══════════════════════════════════════════
story.append(Paragraph("4. 소재별 이송속도 자동 설정", styles['title']))
story.append(HRFlowable(width="100%", thickness=2, color=ACCENT, spaceAfter=5*mm))
story.append(Paragraph("#105 소재 종류에 따라 이송 계수가 자동으로 결정됩니다.", styles['body']))
story.append(Spacer(1, 3*mm))

data = [
    ["소재", "#124 (T01 황삭)", "#125 (T01 스텝)", "#126 (T01 챔퍼)", "#127 (T02 절단)"],
    ["CN (#105=1)", "0.15", "0.12", "0.15", "0.09"],
    ["RS (#105=2)", "0.13", "0.08", "0.05", "0.08"],
    ["CM (#105=3)", "0.13", "0.09", "0.09", "0.09"],
    ["기타 (DEFAULT)", "0.15", "0.10", "0.10", "0.10"],
]
story.append(make_table(data, col_widths=[34*mm, 35*mm, 35*mm, 35*mm, 35*mm]))
story.append(Spacer(1, 5*mm))

story.append(Paragraph("이송속도 계산 공식", styles['h3']))
story.append(info_box("F값 = RPM × Feed 계수", BLUE_BG, ACCENT))
story.append(Spacer(1, 3*mm))
story.append(Paragraph("계산 예시 (RPM = 1700 기준)", styles['h3']))

calc_examples = [
    ["공정", "소재", "계산식", "결과"],
    ["T01 황삭", "CN", "1700 × 0.15", "F255 mm/min"],
    ["T01 스텝", "CN", "1700 × 0.12", "F204 mm/min"],
    ["T01 챔퍼", "CN", "1700 × 0.15", "F255 mm/min"],
    ["T02 절단", "CN", "1700 × 0.09", "F153 mm/min"],
    ["T01 황삭", "RS", "1700 × 0.13", "F221 mm/min"],
    ["T01 챔퍼", "RS", "1700 × 0.05", "F85 mm/min"],
]
story.append(make_table(calc_examples, col_widths=[35*mm, 30*mm, 50*mm, 40*mm], header_color=TEAL))
story.append(Spacer(1, 3*mm))
story.append(Paragraph(
    f'{C("참고:", GRAY)} RS 소재는 경도가 높아 전체적으로 이송이 낮게 설정됩니다.',
    styles['small']
))

story.append(PageBreak())

# ═══════════════════════════════════════════
# 5. 고급 설정 & 안전체크
# ═══════════════════════════════════════════
story.append(Paragraph("5. 고급 설정 & 안전 체크", styles['title']))
story.append(HRFlowable(width="100%", thickness=2, color=ACCENT, spaceAfter=5*mm))
story.append(info_box(
    "DANGER: STOP EDITING — 이 구간은 시스템 자동 설정입니다. 작업자 수정 금지!",
    RED_BG, RED
))
story.append(Spacer(1, 5*mm))

story.append(Paragraph("5-1. 시스템 변수 (#500 시리즈)", styles['h2']))
data = [
    ["변수", "설명", "현재값", "역할"],
    ["#530", "안전 길이", "12", "척 끝에서 안전거리 확보"],
    ["#531", "척 잔여 최소 길이", "15", "가공 후 최소 15mm 조(jaw) 물림 유지\n→ 잔여 길이가 짧으면 소재 이탈 위험"],
    ["#532", "잔여-안전 보정거리", "1.03", "마지막 오토링크 풀링 시\n정확한 소재 위치 계산용 보정값"],
    ["#506", "Z축 황삭 최대 깊이", "55", "한 번에 가공할 최대 Z길이"],
    ["#507", "Z축 황삭 여유 깊이", "1", "황삭 시 Z방향 여유"],
    ["#508", "T01 노즈 R", "0.2", "공구 끝 반경"],
    ["#509", "황삭 간격", "0.4", "X방향 간격"],
    ["#510", "챔퍼 간격", "1", "챔퍼 절삭 시 간격"],
]
story.append(make_table(data, col_widths=[22*mm, 40*mm, 20*mm, 72*mm]))
story.append(Spacer(1, 5*mm))

story.append(Paragraph("5-2. RS40 / 소경 소재 안전 보정", styles['h2']))
story.append(info_box(
    "RS40 소재(#105=2) 또는 외경 30mm 미만(#109&lt;30) → #129=15 자동 설정<br/>"
    "잔재 보정값 15mm를 추가하여 강성 부족에 의한 소재 이탈을 방지합니다.",
    YELLOW_BG, black
))
story.append(Spacer(1, 5*mm))

story.append(Paragraph("5-3. 입력값 검증 흐름", styles['h2']))
story.append(Paragraph("프로그램은 가공 시작 전에 모든 변수를 자동으로 검증합니다:", styles['body']))
story.append(Paragraph("① 모든 변수(#101~#130, #506~#532)가 입력되었는지 확인 (-9999 체크)", styles['body']))
story.append(Paragraph("② 완성길이 > 40mm인 경우 #119=1 확인 플래그 체크", styles['body']))
story.append(Paragraph("③ 미입력 또는 조건 불만족 시 해당 알람 번호로 정지 (#3000 알람)", styles['body']))

story.append(PageBreak())

# ═══════════════════════════════════════════
# 6. 가공 사이클 상세
# ═══════════════════════════════════════════
story.append(Paragraph("6. 가공 사이클 상세", styles['title']))
story.append(HRFlowable(width="100%", thickness=2, color=ACCENT, spaceAfter=5*mm))

story.append(Paragraph("6-1. 전체 흐름", styles['h2']))

flow_data = [
    ["순서", "공정", "공구", "주요 동작"],
    ["①", "좌표계 설정", "-", "G10 L2 P0 Z[#103] — G54 워크좌표 Z원점 설정"],
    ["②", "면삭 (Face)", "T02", "소재 끝면 평삭 (외경→내경 방향)"],
    ["③", "스텝절삭", "T01", "황삭(옵션) + 정삭 — Z방향 절삭"],
    ["④", "챔퍼", "T01", "OD-R, ID-R 챔퍼 가공"],
    ["⑤", "절단", "T02", "부품 분리 + M12 카운트"],
    ["", "↻ 반복", "", "③~⑤를 개수(#517)만큼 반복"],
]
story.append(make_table(flow_data, col_widths=[15*mm, 30*mm, 17*mm, 92*mm], header_color=DARK_BLUE))
story.append(Spacer(1, 5*mm))

story.append(Paragraph("주요 계산 변수", styles['h3']))
calc_data = [
    ["계산식", "의미", "현재값"],
    ["#505 = #113 + #114", "1개당 가공 단위 (완성길이 + 절단폭)", "9.843mm"],
    ["#500 = #103 - #530", "유효 가공 길이 (척길이 - 안전길이)", "63mm"],
    ["#517 = FIX[#500 / #505]", "1회 인출당 가공 개수", "6개"],
    ["#518 = #517 × #505", "1회 총 가공 길이", "59.058mm"],
    ["#140 = #101 + #102 - #129", "전체 소재 길이 (안전 보정 적용)", "543mm"],
]
story.append(make_table(calc_data, col_widths=[50*mm, 70*mm, 34*mm], header_color=TEAL))

story.append(Spacer(1, 8*mm))

# 6-2 면삭
story.append(Paragraph("6-2. 면삭 (N100) — T02", styles['h2']))
story.append(Paragraph("소재 끝면을 평평하게 깎는 공정입니다. 외경에서 내경 방향으로 절삭합니다.", styles['body']))
code_lines = [
    ("G00 W3. T02", "T02 공구 선택, Z+3mm 접근"),
    ("X[#501]  (=74)", "소재외경+4mm 위치로 급속이동"),
    ("Z0.", "Z원점(워크 끝면)으로 이동"),
    ("G97 M03 S1700", "정속회전 1700rpm, 정회전 시작"),
    ("G98 G01 X[#502]  (=52.5)", "내경까지 면삭, F=1700×0.09=153mm/min"),
    ("G00 X[#111] T01", "T01 공구로 교환, 완성외경 위치로 이동"),
]
for code, desc in code_lines:
    story.append(Paragraph(f'<font face="Courier" size="12">{code}</font>  <font color="{GRAY}" size="11">← {desc}</font>', styles['body']))

story.append(Spacer(1, 5*mm))

# 6-3 스텝절삭
story.append(Paragraph("6-3. 스텝절삭 (N130) — T01", styles['h2']))

story.append(Paragraph(f'{C("[황삭]", RED)} #107=1일 때만 실행', styles['body_bold']))
story.append(Paragraph("X[#111+0.4] (완성외경+황삭간격) 위치에서 Z방향으로 직선절삭", styles['body']))
story.append(Paragraph(f'이송속도: F = 1700 × 0.15 = {C("255 mm/min", ACCENT)}', styles['body']))
story.append(Spacer(1, 2*mm))

story.append(Paragraph(f'{C("[정삭]", TEAL)} #106=3일 때 실행', styles['body_bold']))
story.append(Paragraph("X[#111] (완성외경) 위치에서 Z방향으로 정밀 절삭", styles['body']))
story.append(Paragraph(f'이송속도: F = 1700 × 0.12 = {C("204 mm/min", ACCENT)}', styles['body']))
story.append(Paragraph("M[#134] 보링바 DOWN → G4 X.1 (0.1초 대기) → U0.2 (빠짐, 스프링백 방지)", styles['body']))

story.append(PageBreak())

# 6-4 챔퍼
story.append(Paragraph("6-4. 챔퍼 가공 (N160) — T01", styles['h2']))

story.append(Paragraph(f'{C("TYPE 3", TEAL)} (#106=3) — 기본 챔퍼', styles['body_bold']))
story.append(Paragraph("완성외경 위치에서 챔퍼 시작점으로 접근 후:", styles['body']))
story.append(Paragraph("• X[#111 - (#510+#116)×2] → 외경 R(OD-R) 챔퍼 가공", styles['body']))
story.append(Paragraph("• X[#111 + (#510+#115)×2] → 내경 R(ID-R) 챔퍼 가공", styles['body']))
story.append(Paragraph(f'이송속도: F = 1700 × 0.15 = {C("255 mm/min", ACCENT)}', styles['body']))
story.append(Spacer(1, 3*mm))

story.append(Paragraph(f'{C("TYPE 4", ORANGE)} (#106=4) — 확장 챔퍼', styles['body_bold']))
story.append(Paragraph("TYPE 3과 유사하나 추가 동작이 있습니다:", styles['body']))
story.append(Paragraph("• 보링바 UP/DOWN 동작 추가", styles['body']))
story.append(Paragraph("• 완성길이 위치에서 추가 챔퍼: #118(외경챔퍼 0.60mm), #117(내경챔퍼 0.37mm)", styles['body']))
story.append(Paragraph("• RS 소재(#105=2)일 경우 G4 X0.5 (0.5초 대기) 추가 — 절삭 안정화", styles['body']))

story.append(Spacer(1, 5*mm))

# 6-5 절단
story.append(Paragraph("6-5. 절단 (N180) — T02", styles['h2']))

story.append(Paragraph(f'{C("TYPE 3", TEAL)} — 챔퍼 동시 절단', styles['body_bold']))
code_lines = [
    ("G00 Z[#118-#522] T02", "절단 위치 (챔퍼 고려)"),
    ("G97 G00 X[#111+1] S1700 M03", "외경+1mm에서 접근"),
    ("G98 G01 X[#111] F[...]", "완성외경까지 절삭"),
    ("U-[#118×2] W-[#118]", "외경 챔퍼 동시 가공 (경사절삭)"),
    ("X[#504]  (=53.3)", "내경까지 관통 절단"),
    ("M12", "완성 부품 수 카운트"),
]
for code, desc in code_lines:
    story.append(Paragraph(f'<font face="Courier" size="12">{code}</font>  <font color="{GRAY}" size="11">← {desc}</font>', styles['body']))

story.append(Spacer(1, 3*mm))
story.append(Paragraph(f'{C("TYPE 4", ORANGE)} — 직선 절단', styles['body_bold']))
story.append(Paragraph("Z-[#522]에서 수직 진입 → X[#503]→X[#111]→X[#504] 관통 → M12 카운트", styles['body']))

story.append(Spacer(1, 5*mm))

# 6-6 반복 판단
story.append(Paragraph("6-6. 반복 판단 (N190)", styles['h2']))
story.append(Paragraph("#520(현재 가공 수)과 #515(스텝 개수), #517(총 개수)을 비교합니다:", styles['body']))
repeat_data = [
    ["조건", "동작", "의미"],
    ["#520 < #515", "GOTO 160 (챔퍼)", "현재 스텝 내 아직 남음"],
    ["(#521+#515) > #517", "GOTO 200 (마무리)", "다음 스텝 실행하면 총 개수 초과"],
    ["그 외", "GOTO 130 (스텝절삭)", "다음 스텝 시작"],
]
story.append(make_table(repeat_data, col_widths=[45*mm, 50*mm, 59*mm]))

story.append(PageBreak())

# ═══════════════════════════════════════════
# 7. 오토링크 & 잔재 가공
# ═══════════════════════════════════════════
story.append(Paragraph("7. 오토링크 & 잔재 가공", styles['title']))
story.append(HRFlowable(width="100%", thickness=2, color=ACCENT, spaceAfter=5*mm))

story.append(Paragraph("7-1. 오토링크 동작 순서 (N210)", styles['h2']))
story.append(Paragraph("소재가 부족해지면 자동으로 소재를 당겨오는 핵심 기능입니다.", styles['body']))
story.append(Spacer(1, 3*mm))

al_data = [
    ["순서", "코드", "동작", "설명"],
    ["①", "M05", "스핀들 정지", "안전을 위해 회전 정지"],
    ["②", "G00 T03, X0", "공구 선택", "오토링크 공구, 중심 이동"],
    ["③", "Z[-#522+#122]", "인출 위치 접근", "F200으로 위치 이동"],
    ["④", "M[#132]", "오토로더 CLOSE", "오토로더가 소재를 잡음"],
    ["⑤", "M69 + G04 P2500", "척 UNCLAMP", "척이 소재를 놓음 (2.5초 대기)"],
    ["⑥", "W[가공길이+여유]", "소재 당김(PULL)", "소재를 앞으로 인출"],
    ["⑦", "M68 + G04 P2500", "척 CLAMP", "척이 소재를 다시 잡음 (2.5초 대기)"],
    ["⑧", "M[#131]", "오토로더 OPEN", "오토로더 해제"],
    ["⑨", "G00 Z100 → GOTO 100", "복귀", "안전위치 → 면삭부터 재시작"],
]
story.append(make_table(al_data, col_widths=[15*mm, 35*mm, 33*mm, 71*mm], header_color=ORANGE))
story.append(Spacer(1, 5*mm))

story.append(info_box(
    "안전 경고: 오토링크 동작 중에는 척이 열리고 닫히는 구간이 있습니다.<br/>"
    "이 구간에서 절대로 소재나 공구에 손대지 마세요!<br/>"
    "대기 시간(G04)이 충분히 설정되어 있으므로 임의로 줄이지 마세요.",
    RED_BG, RED
))
story.append(Spacer(1, 5*mm))

story.append(Paragraph("7-2. 잔재 가공 (O9003)", styles['h2']))
story.append(Paragraph("오토링크 횟수를 모두 사용한 뒤, 남은 소재를 처리하는 서브프로그램입니다.", styles['body']))
story.append(Paragraph("① 남은 가공 가능 길이 계산 (#540 = #539 × #505)", styles['body']))
story.append(Paragraph("② 척 길이 재설정 (#103 = #530 + #540)", styles['body']))
story.append(Paragraph("③ 좌표계 재설정 (G10 L2 P0 Z[#103])", styles['body']))
story.append(Paragraph("④ 면삭 → 스텝절삭 → 챔퍼 → 절단 사이클 반복", styles['body']))
story.append(Paragraph("⑤ 완료 후 M99 리턴", styles['body']))

story.append(PageBreak())

# ═══════════════════════════════════════════
# 8. 알람 코드표
# ═══════════════════════════════════════════
story.append(Paragraph("8. 알람 코드표", styles['title']))
story.append(HRFlowable(width="100%", thickness=2, color=ACCENT, spaceAfter=5*mm))
story.append(Paragraph("에러 발생 시 알람 번호(#3000)를 확인하고 아래 표를 참조하세요.", styles['body']))
story.append(Spacer(1, 3*mm))

story.append(Paragraph("8-1. 가공 로직 알람", styles['h2']))
alarm_data1 = [
    ["알람", "조건", "의미", "조치"],
    ["1", "#140 > 580", "소재 길이 580mm 초과", "#101, #102 확인"],
    ["2", "#109 < #111", "소재 외경 < 완성 외경", "#109, #111 확인"],
    ["3", "#110 > #112", "소재 내경 > 완성 내경", "#110, #112 확인"],
    ["4", "#125 ≥ 0.14", "T01 스텝 이송 과대", "소재종류(#105) 확인"],
    ["5", "#102 단위 오류", "100단위 아님", "0/100/200/300/400/500"],
    ["6", "#101 > 99", "끝자리 범위 초과", "0~99로 수정"],
    ["7", "#140 ≤ #103", "소재 길이 부족", "소재 길이 확인"],
    ["10", "#104 > 40", "초기 면삭량 과대", "#104를 40 이하로"],
    ["11", "#129 > 30", "잔재 보정 과대", "#129 확인"],
    ["12", "#123 범위 초과", "여유값 0~3 범위 밖", "#123 수정"],
    ["13", "#106 ≠ 3,4", "가공유형 오류", "#106을 3 또는 4로"],
]
story.append(make_table(alarm_data1, col_widths=[17*mm, 35*mm, 48*mm, 54*mm], header_color=RED))
story.append(Spacer(1, 5*mm))

story.append(Paragraph("8-2. 안전 검증 알람", styles['h2']))
alarm_data2 = [
    ["알람", "조건", "의미", "조치"],
    ["203", "OD ≤ ID", "소재 외경 ≤ 소재 내경", "#109, #110 확인"],
    ["204", "FIN OD ≤ FIN ID", "완성 외경 ≤ 완성 내경", "#111, #112 확인"],
    ["205", "#109 ≤ 0", "외경이 0 이하", "#109 확인"],
    ["206", "#120 ≤ 0", "T01 RPM이 0 이하", "#120 확인"],
    ["207", "#121 ≤ 0", "T02 RPM이 0 이하", "#121 확인"],
    ["208", "#114 ≥ #113", "절단폭 ≥ 완성길이", "#114, #113 확인"],
]
story.append(make_table(alarm_data2, col_widths=[17*mm, 35*mm, 48*mm, 54*mm], header_color=ORANGE))
story.append(Spacer(1, 5*mm))

story.append(Paragraph("8-3. 변수 미입력 알람 (101~130, 506~532)", styles['h2']))
story.append(info_box(
    "알람 번호가 101~130 또는 506~532인 경우: 해당 번호의 변수(#번호)가 미입력 상태입니다.<br/>"
    "예) 알람 109 → #109(소재 외경)가 입력되지 않았음 → 해당 변수에 값을 입력하세요.",
    YELLOW_BG, black
))

story.append(PageBreak())

# ═══════════════════════════════════════════
# 9. 기계별 M코드 매핑
# ═══════════════════════════════════════════
story.append(Paragraph("9. 기계별 M코드 매핑", styles['title']))
story.append(HRFlowable(width="100%", thickness=2, color=ACCENT, spaceAfter=5*mm))
story.append(Paragraph("#130 기계번호에 따라 오토로더/보링바 M코드가 자동 매핑됩니다.", styles['body']))
story.append(Spacer(1, 3*mm))

machine_data = [
    ["기계\n번호", "이름", "AL OPEN\n(#131)", "AL CLOSE\n(#132)", "BR UP\n(#133)", "BR DOWN\n(#134)"],
    ["1", "AL-1", "M56", "M55", "M54", "M53"],
    ["4", "HA-4", "M52", "M51", "M54", "M53"],
    ["5", "AL-5", "M171", "M170", "M53", "M54"],
    ["7", "AL-7", "M64", "M63", "M53", "M54"],
    ["8", "AL-8", "M64", "M63", "M54", "M53"],
    ["9", "AL-9", "M63", "M64", "M53", "M54"],
    ["10", "AL-10", "M64", "M63", "M54", "M53"],
    ["13", "HA-S3", "M56", "M55", "M54", "M53"],
    ["14", "HA-S4", "M56", "M55", "M54", "M53"],
]
story.append(make_table(machine_data, col_widths=[18*mm, 22*mm, 28*mm, 28*mm, 28*mm, 28*mm]))
story.append(Spacer(1, 5*mm))

story.append(info_box(
    "AL = Auto Loader (자동 소재 공급장치)  |  BR = Boring bar (보링바)<br/>"
    "기계 변경 시 #130 값만 해당 번호로 수정하면 M코드가 자동 매핑됩니다.",
    BLUE_BG, DARK_BLUE
))
story.append(Spacer(1, 5*mm))

story.append(Paragraph("주의: 기계별 BR UP/DOWN이 다릅니다", styles['h3']))
story.append(Paragraph("• AL-5, AL-7, AL-9: BR UP=M53, BR DOWN=M54 (반대 배치)", styles['body']))
story.append(Paragraph("• 나머지 기계: BR UP=M54, BR DOWN=M53 (표준 배치)", styles['body']))
story.append(Paragraph("기계를 변경할 때 반드시 #130만 수정하고, M코드를 직접 수정하지 마세요.", styles['warn']))

# ═══════════════════════════════════════════
# 10. 작업자 주의사항
# ═══════════════════════════════════════════
story.append(Paragraph("10. 작업자 주의사항", styles['title']))
story.append(HRFlowable(width="100%", thickness=2, color=ACCENT, spaceAfter=5*mm))

story.append(Paragraph("10-1. 필수 확인 사항", styles['h2']))
must_checks = [
    ["번호", "확인 사항", "상세"],
    ["1", "수정 가능 영역 확인", "#101~#123만 수정 가능.\nDANGER 이후 구간은 절대 수정 금지"],
    ["2", "#119 확인 (긴 부품)", "완성길이(#113) > 40mm일 경우\n반드시 #119=1로 설정"],
    ["3", "#102 입력 형식", "0, 100, 200, 300, 400, 500만 허용\n그 외 값 입력 시 알람 5번"],
    ["4", "#101 입력 범위", "0~99 범위만 허용\n100 이상 입력 시 알람 6번"],
    ["5", "기계 변경", "#130만 수정 (1,4,5,7,8,9,10,13,14)\nM코드 직접 수정 금지"],
]
story.append(make_table(must_checks, col_widths=[15*mm, 40*mm, 99*mm], header_color=NAVY))
story.append(Spacer(1, 5*mm))

story.append(Paragraph("10-2. 안전 수칙", styles['h2']))
story.append(info_box(
    "① 오토링크 중 개입 금지 — 척 개폐 동작 중 소재/공구 접촉 절대 금지!<br/>"
    "② RS40 소재 — 안전보정 15mm 자동 적용, 임의 해제 금지<br/>"
    "③ 소경 소재 (OD&lt;30mm) — 동일 안전보정 자동 적용<br/>"
    "④ 알람 발생 시 — 알람 번호 확인 → 코드표 참조 → 해당 변수 수정 후 재시작<br/>"
    "⑤ G04 대기 시간 — 오토링크 내 대기 시간을 임의로 줄이지 마세요",
    RED_BG, RED
))
story.append(Spacer(1, 5*mm))

story.append(Paragraph("10-3. 가공 전 체크리스트", styles['h2']))
checklist = [
    ["", "체크 항목"],
    ["□", "소재 종류(#105)와 실제 소재가 일치하는가?"],
    ["□", "소재 외경/내경(#109, #110)과 실측값이 일치하는가?"],
    ["□", "완성 치수(#111~#118)와 도면이 일치하는가?"],
    ["□", "기계번호(#130)와 실제 작업 기계가 일치하는가?"],
    ["□", "RPM(#120, #121)이 소재/공구에 적합한가?"],
    ["□", "첫 가공 시 단품모드(#108=1)로 시운전을 실시했는가?"],
    ["□", "오토로더 동작 확인이 완료되었는가?"],
    ["□", "절단 바이트(T02) 마모 상태를 확인했는가?"],
]
t = Table(checklist, colWidths=[14*mm, 140*mm])
t.setStyle(TableStyle([
    ('FONTNAME', (0,0), (-1,0), 'MalgunBd'),
    ('FONTNAME', (0,1), (-1,-1), 'Malgun'),
    ('FONTSIZE', (0,0), (-1,-1), 13),
    ('LEADING', (0,0), (-1,-1), 20),
    ('BACKGROUND', (0,0), (-1,0), GREEN),
    ('TEXTCOLOR', (0,0), (-1,0), white),
    ('TEXTCOLOR', (0,1), (-1,-1), black),
    ('ALIGN', (0,0), (0,-1), 'CENTER'),
    ('ALIGN', (1,0), (1,-1), 'LEFT'),
    ('GRID', (0,0), (-1,-1), 0.5, HexColor('#CBD5E1')),
    ('TOPPADDING', (0,0), (-1,-1), 4),
    ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ('ROWBACKGROUNDS', (0,1), (-1,-1), [white, GREEN_BG]),
]))
story.append(t)

# ═══════════════════════════════════════════
# 마지막: 핵심 정리
# ═══════════════════════════════════════════
story.append(Spacer(1, 10*mm))
story.append(Paragraph("O0852 프로그램 핵심 정리", styles['title']))
story.append(HRFlowable(width="100%", thickness=2, color=ACCENT, spaceAfter=8*mm))

summary_items = [
    "작업자는 #101~#123 변수만 수정합니다 (소재, 치수, 조건)",
    "소재 종류(CN/RS/CM)에 따라 이송속도가 자동 설정됩니다",
    "면삭 → 스텝절삭 → 챔퍼 → 절단 사이클이 자동으로 반복됩니다",
    "소재 부족 시 오토링크로 자동 인출 후 재가공합니다",
    "30개 이상의 알람으로 입력 오류/안전 문제를 사전에 차단합니다",
    "9대 기계의 M코드가 자동 매핑됩니다 (#130 변경만으로 전환)",
]

for i, item in enumerate(summary_items):
    story.append(Paragraph(
        f'<font color="{ACCENT}"><b>{i+1}.</b></font>  {item}',
        ParagraphStyle('sum', fontName='Malgun', fontSize=16, leading=28, textColor=NAVY, leftIndent=5*mm, spaceAfter=4*mm)
    ))

story.append(Spacer(1, 15*mm))
story.append(HRFlowable(width="40%", thickness=1, color=GRAY, spaceAfter=5*mm, spaceBefore=5*mm))
story.append(Paragraph("질문이 있으시면 언제든 문의하세요", ParagraphStyle('footer', fontName='Malgun', fontSize=15, textColor=GRAY, alignment=TA_CENTER)))

# ─── 빌드 ───
doc.build(story)
print("PDF saved successfully!")
