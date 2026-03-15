"""
폐수배출시설 운영일지 자동 생성기
- 월별 docx 하나에 일별 데이터 기록
- 1호기·2호기 계량기 지침 입력 → 용수사용량 자동 계산
- 주말/쉬는날 자동 처리
- 날짜 열 클릭 → 쉬는날 토글
- 템플릿(폐수배출시설_운영일지_template.docx) 기반 XML 치환 (순수 문자열 조작)
- Python 기본 라이브러리만 사용 (ElementTree 미사용 → XML 손상 없음)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import zipfile, os, re, io, json, random, calendar
from datetime import date, timedelta

# ── 경로 ──
_HERE         = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH   = os.path.join(_HERE, 'water_config.json')
TEMPLATE_PATH = os.path.join(_HERE, '폐수배출시설_운영일지_template.docx')

# ── 설정 저장/로드 ──
def load_config():
    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}

def save_config(data: dict):
    try:
        cfg = load_config()
        cfg.update(data)
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _rows_key(year, month):
    return f'rows_{year}_{month:02d}'


def save_rows(year, month, rows):
    """연/월 행 데이터를 config에 저장"""
    data_to_save = []
    for r in rows:
        data_to_save.append({
            'day':        r['day'],
            'is_holiday': r.get('is_holiday', False),
            'w1':         r.get('w1', ''),
            'w2':         r.get('w2', ''),
            'm1':         r.get('m1', ''),
            'm2':         r.get('m2', ''),
            'note':       r.get('note', ''),
        })
    save_config({_rows_key(year, month): data_to_save})


def load_rows(year, month):
    """config에서 연/월 행 데이터 불러오기. 없으면 None"""
    cfg = load_config()
    return cfg.get(_rows_key(year, month))

# ── 상수 ──
DOW_KR = ['월', '화', '수', '목', '금', '토', '일']
REPRESENTATIVE = '김예배'
ENV_TECH       = '신종한'

COLS = ('날짜', '요일', '가동시간', '1호기용수(L)', '2호기용수(L)',
        '1호기계량기', '2호기계량기', '비고')
COL_WIDTHS = [78, 38, 90, 82, 82, 90, 90, 100]

# ── 유틸 ──
def is_wknd(d):          return d.weekday() >= 5
def get_dow(d):          return DOW_KR[d.weekday()]
def days_in_month(y, m): return calendar.monthrange(y, m)[1]


# ══════════════════════════════════════════════════════════════════
# 순수 문자열 XML 조작 (ElementTree 미사용 → 네임스페이스 보존)
# ══════════════════════════════════════════════════════════════════

def _xml_escape(text):
    """XML 특수문자 이스케이프"""
    return (str(text)
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;'))


def _set_cell_text(cell_xml, new_text):
    """
    셀 XML 내 <w:t> 태그 교체.
    첫 번째 <w:t>에 new_text 를 넣고 나머지는 비움.
    <w:t> 가 하나도 없는 빈 셀이면 </w:p> 직전에 <w:r><w:t>TEXT</w:t></w:r> 삽입.
    NOTE: 이 템플릿은 셀 안에 중첩 테이블이 없다고 가정합니다.
    """
    escaped = _xml_escape(new_text)

    # <w:t ...>...</w:t> 만 정확히 매칭 (접두사가 'w:t' + 공백 또는 바로 '>')
    pat = re.compile(r'<w:t(\s[^>]*)?>.*?</w:t>', re.DOTALL)
    matches = list(pat.finditer(cell_xml))

    if not matches:
        # 기존 <w:t> 가 없는 빈 셀 → </w:p> 직전에 텍스트 런 삽입
        if not new_text:
            return cell_xml  # 비어 있어도 괜찮으면 그대로 반환
        wp_end = cell_xml.rfind('</w:p>')
        if wp_end == -1:
            return cell_xml  # <w:p> 를 찾지 못하면 포기
        # 간단한 텍스트 런 삽입 (서식 없음)
        insert = f'<w:r><w:t>{escaped}</w:t></w:r>'
        return cell_xml[:wp_end] + insert + cell_xml[wp_end:]

    result = cell_xml
    # 뒤에서부터 치환 → 앞쪽 오프셋 유지
    for i in range(len(matches) - 1, -1, -1):
        m = matches[i]
        attrs = m.group(1) or ''
        if i == 0:
            # 공백 보존이 필요하면 xml:space="preserve" 추가
            if new_text and (new_text != new_text.strip() or '  ' in new_text):
                if 'xml:space' not in attrs:
                    attrs = ' xml:space="preserve"' + attrs
            new_tag = f'<w:t{attrs}>{escaped}</w:t>'
        else:
            new_tag = f'<w:t{attrs}></w:t>'
        result = result[:m.start()] + new_tag + result[m.end():]

    return result


def _find_row_spans(tbl_xml):
    """tbl_xml 내 모든 <w:tr>...</w:tr> 위치 반환 (중첩 없다고 가정)"""
    spans = []
    pos = 0
    while True:
        rs = tbl_xml.find('<w:tr', pos)
        if rs == -1:
            break
        re_end = tbl_xml.find('</w:tr>', rs)
        if re_end == -1:
            break
        re_end += 7  # len('</w:tr>')
        spans.append((rs, re_end))
        pos = re_end
    return spans


def _find_cell_spans(row_xml):
    """row_xml 내 모든 <w:tc>...</w:tc> 위치 반환 (중첩 없다고 가정)"""
    spans = []
    pos = 0
    while True:
        cs = row_xml.find('<w:tc', pos)
        if cs == -1:
            break
        ce = row_xml.find('</w:tc>', cs)
        if ce == -1:
            break
        ce += 7  # len('</w:tc>')
        spans.append((cs, ce))
        pos = ce
    return spans


def _process_table(tbl_xml, tidx, year, month, dim, daily_data, holidays, yy):
    """
    테이블 XML을 받아 데이터 행을 채워서 반환.
    tidx=0 → 1~15일, tidx=1 → 16~말일
    """
    day_start    = 1 if tidx == 0 else 16
    header_count = 5
    footer_count = 1

    row_spans = _find_row_spans(tbl_xml)
    data_start = header_count
    data_end   = len(row_spans) - footer_count

    new_tbl = tbl_xml

    # 뒤에서부터 처리 → 앞쪽 오프셋 불변
    for row_i in range(data_end - 1, data_start - 1, -1):
        day_offset = row_i - data_start
        day_num    = day_start + day_offset

        rs, re_end  = row_spans[row_i]
        row_xml     = new_tbl[rs:re_end]
        cell_spans  = _find_cell_spans(row_xml)

        if len(cell_spans) < 10:
            continue

        edits = {}  # {cell_idx: new_text}

        if day_num > dim:
            # 해당 월에 존재하지 않는 날 → 전부 비움 (예: 2월 30·31일)
            for ci in range(len(cell_spans)):
                edits[ci] = ''
        else:
            d       = date(year, month, day_num)
            weekend = is_wknd(d) or (day_num in holidays)
            dd      = daily_data.get(day_num, {})

            # 셀 0: 날짜
            if row_i == data_start:
                edits[0] = f'{yy}/{month:02d}/{day_num:02d}'
            else:
                edits[0] = str(day_num)

            # 셀 1-2: 결재 (대표자 / 환경기술인)
            if row_i == data_start:
                edits[1] = REPRESENTATIVE
                edits[2] = ENV_TECH
            else:
                edits[1] = '\u201c'   # "
                edits[2] = '\u201c'

            # 셀 3-5: 가동시간 (3개 셀로 분할된 템플릿)
            if weekend:
                if day_num in holidays and d.weekday() < 5:
                    edits[3] = '쉬는날'
                elif d.weekday() == 5:
                    edits[3] = '토요일'
                else:
                    edits[3] = '일요일'
                edits[4] = ''
                edits[5] = ''
            else:
                edits[3] = '08:00~1'
                edits[4] = '7'
                edits[5] = ':00'

            # 셀 6-7: 용수사용량
            edits[6] = str(dd.get('w1', '0' if weekend else ''))
            edits[7] = str(dd.get('w2', '0' if weekend else ''))

            # 셀 8-9: 계량기지침
            edits[8] = str(dd.get('m1', ''))
            edits[9] = str(dd.get('m2', ''))

            # 셀 10 이상: 비움
            for ci in range(10, len(cell_spans)):
                edits[ci] = ''

        # 셀 뒤에서부터 치환 → 앞쪽 오프셋 불변
        modified_row = row_xml
        for ci in sorted(edits.keys(), reverse=True):
            if ci >= len(cell_spans):
                continue
            cs, ce    = cell_spans[ci]
            new_cell  = _set_cell_text(modified_row[cs:ce], edits[ci])
            modified_row = modified_row[:cs] + new_cell + modified_row[ce:]

        new_tbl = new_tbl[:rs] + modified_row + new_tbl[re_end:]

    return new_tbl


def _fill_document_xml(xml_str, year, month, dim, daily_data, holidays):
    """
    document.xml 문자열을 받아 2개 테이블의 데이터 행을 채운 뒤 반환.
    ElementTree 를 전혀 사용하지 않아 네임스페이스 선언이 보존됩니다.
    """
    yy = year % 100

    # ── 테이블 2개 위치 찾기 ──
    # 이 템플릿에는 중첩 테이블이 없으므로 단순 find 로 충분합니다.
    table_spans = []
    pos = 0
    while len(table_spans) < 2:
        s = xml_str.find('<w:tbl', pos)
        if s == -1:
            break
        e = xml_str.find('</w:tbl>', s)
        if e == -1:
            break
        e += 8  # len('</w:tbl>')
        table_spans.append((s, e))
        pos = e

    if len(table_spans) < 2:
        return xml_str  # 테이블을 못 찾으면 원본 반환

    # ── 테이블별 처리 ──
    new_tables = []
    for tidx in range(2):
        tbl_s, tbl_e = table_spans[tidx]
        tbl_xml = xml_str[tbl_s:tbl_e]
        new_tbl = _process_table(tbl_xml, tidx, year, month, dim,
                                 daily_data, holidays, yy)
        new_tables.append(new_tbl)

    # ── 재조립 (원본 위치를 사용하므로 오프셋 문제 없음) ──
    t0_s, t0_e = table_spans[0]
    t1_s, t1_e = table_spans[1]
    result = (xml_str[:t0_s]
              + new_tables[0]
              + xml_str[t0_e:t1_s]
              + new_tables[1]
              + xml_str[t1_e:])
    return result


def fill_water_docx(tpl_bytes, year, month, daily_data, holidays=set()):
    """템플릿 docx 를 기반으로 월별 운영일지를 생성하여 bytes 로 반환."""
    dim    = days_in_month(year, month)
    buf_in = io.BytesIO(tpl_bytes)
    buf_out= io.BytesIO()

    with zipfile.ZipFile(buf_in, 'r') as src, \
         zipfile.ZipFile(buf_out, 'w', zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            raw = src.read(item.filename)
            if item.filename == 'word/document.xml':
                xml_str = raw.decode('utf-8')
                xml_str = _fill_document_xml(
                    xml_str, year, month, dim, daily_data, holidays)
                raw = xml_str.encode('utf-8')
            dst.writestr(item, raw)

    return buf_out.getvalue()


# ══════════════════════════════════════════════════════════════════
# GUI
# ══════════════════════════════════════════════════════════════════
class WaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title('폐수배출시설 운영일지 생성기')
        self.root.geometry('980x720')
        self.root.configure(bg='#0f1923')
        self.root.resizable(True, True)

        self.tpl_bytes = None
        self.out_dir   = tk.StringVar(value='선택 안 됨')
        self.rows: list = []
        self._current_year  = None
        self._current_month = None

        now = date.today()
        self.year_var  = tk.IntVar(value=now.year)
        self.month_var = tk.IntVar(value=now.month)

        self._load_template()

        cfg = load_config()
        if cfg.get('out_dir'):
            self.out_dir.set(cfg['out_dir'])

        self._setup_style()
        self._build_ui()

        if cfg.get('out_dir'):
            self.out_lbl.config(fg='#34d399')

        self.refresh_table()
        self.update_status()

    # ── 템플릿 로드 ──
    def _load_template(self):
        if os.path.exists(TEMPLATE_PATH):
            try:
                with open(TEMPLATE_PATH, 'rb') as f:
                    self.tpl_bytes = f.read()
            except Exception:
                self.tpl_bytes = None

    # ── 스타일 ──
    def _setup_style(self):
        s = ttk.Style()
        s.theme_use('default')
        s.configure('Treeview',
            background='#131f2e', foreground='#e0e8f0',
            fieldbackground='#131f2e', rowheight=23,
            font=('맑은 고딕', 9))
        s.configure('Treeview.Heading',
            background='#0d1e33', foreground='#7dd3fc',
            font=('맑은 고딕', 9, 'bold'))
        s.map('Treeview', background=[('selected', '#1e3a5f')])
        s.configure('Vertical.TScrollbar',   background='#1e3a5f', troughcolor='#0d1e33')
        s.configure('Horizontal.TScrollbar', background='#1e3a5f', troughcolor='#0d1e33')

    def _btn(self, parent, text, cmd, color='#2e75b6', fg='white', **kw):
        b = tk.Button(parent, text=text, command=cmd,
                      bg=color, fg=fg, font=('맑은 고딕', 9, 'bold'),
                      relief='flat', cursor='hand2', **kw)
        def lighten(c):
            r, g, bl = int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)
            return '#{:02x}{:02x}{:02x}'.format(
                min(r+25, 255), min(g+25, 255), min(bl+25, 255))
        b.bind('<Enter>', lambda e: b.config(bg=lighten(color)))
        b.bind('<Leave>', lambda e: b.config(bg=color))
        return b

    def _section(self, parent, title):
        lbl = tk.Label(parent, text=title, fg='#7dd3fc', bg='#0f1923',
                       font=('맑은 고딕', 9, 'bold'), anchor='w')
        lbl.pack(fill='x', pady=(10, 4), padx=6)
        sep = tk.Frame(parent, bg='#1e3a5f', height=1)
        sep.pack(fill='x', padx=6, pady=(0, 6))

    # ── UI 구성 ──
    def _build_ui(self):
        # 상태바는 반드시 먼저 pack (bottom) → expand=True 위젯이 나머지를 채워야 함
        self.status = tk.Label(self.root, text='', bg='#0a1018', fg='#64748b',
                               font=('맑은 고딕', 8), anchor='w', padx=8)
        self.status.pack(fill='x', side='bottom')

        # 메인 컨테이너
        main_frame = tk.Frame(self.root, bg='#0f1923')
        main_frame.pack(fill='both', expand=True)

        # ── 오른쪽 사이드바 (고정 너비 300px) ──
        # side='right' 를 먼저 pack 해야 left 가 남은 공간을 올바로 채움
        right_outer = tk.Frame(main_frame, bg='#0f1923', width=300)
        right_outer.pack(side='right', fill='y')
        right_outer.pack_propagate(False)          # 프레임 크기 고정

        right_canvas = tk.Canvas(right_outer, bg='#0f1923', highlightthickness=0)
        sb_scroll = ttk.Scrollbar(right_outer, orient='vertical',
                                  command=right_canvas.yview)
        sb_frame = tk.Frame(right_canvas, bg='#0f1923')
        sb_frame.bind('<Configure>',
            lambda e: right_canvas.configure(
                scrollregion=right_canvas.bbox('all')))
        right_canvas.create_window((0, 0), window=sb_frame, anchor='nw')
        right_canvas.configure(yscrollcommand=sb_scroll.set)
        sb_scroll.pack(side='right', fill='y')
        right_canvas.pack(side='left', fill='both', expand=True)
        right_canvas.bind_all('<MouseWheel>',
            lambda e: right_canvas.yview_scroll(
                int(-1 * (e.delta / 120)), 'units'))

        self._build_sidebar(sb_frame)

        # ── 왼쪽 테이블 (나머지 공간 전부) ──
        left = tk.Frame(main_frame, bg='#0f1923')
        left.pack(side='left', fill='both', expand=True)
        self._build_table(left)

    def _build_table(self, parent):
        self.tree = ttk.Treeview(parent, columns=COLS, show='headings',
                                  selectmode='extended')
        for i, c in enumerate(COLS):
            w = COL_WIDTHS[i] if i < len(COL_WIDTHS) else 80
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, minwidth=40, anchor='center')

        vsb = ttk.Scrollbar(parent, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        self.tree.pack(fill='both', expand=True)

        self.tree.bind('<ButtonRelease-1>', self.on_click)
        self.tree.bind('<Double-1>',        self.on_double_click)
        self.tree.bind('<Control-c>',       self.copy_selection)
        self.tree.bind('<Control-v>',       self.paste_to_table)

        self.tree.tag_configure('weekend', foreground='#f87171')
        self.tree.tag_configure('holiday', foreground='#fb923c')

    def _build_sidebar(self, parent):
        # ① 년월 설정
        self._section(parent, '① 년월 설정')

        row_f = tk.Frame(parent, bg='#0f1923')
        row_f.pack(fill='x', padx=6, pady=2)
        tk.Label(row_f, text='년:', fg='#94a3b8', bg='#0f1923',
                 font=('맑은 고딕', 9)).pack(side='left')
        tk.Spinbox(row_f, from_=2020, to=2040, width=6,
                   textvariable=self.year_var,
                   bg='#1e3a5f', fg='#e0e8f0', font=('맑은 고딕', 9),
                   buttonbackground='#2e75b6').pack(side='left', padx=(2, 10))
        tk.Label(row_f, text='월:', fg='#94a3b8', bg='#0f1923',
                 font=('맑은 고딕', 9)).pack(side='left')
        tk.Spinbox(row_f, from_=1, to=12, width=4,
                   textvariable=self.month_var,
                   bg='#1e3a5f', fg='#e0e8f0', font=('맑은 고딕', 9),
                   buttonbackground='#2e75b6').pack(side='left', padx=2)

        self._btn(parent, '📅 테이블 새로고침', self.refresh_table
                  ).pack(fill='x', padx=6, pady=4)

        # ② 쉬는날 설정
        self._section(parent, '② 쉬는날 설정')
        tk.Label(parent,
                 text='  [방법] 왼쪽 날짜 열 클릭 → 쉬는날 토글\n'
                      '  [휴] 표시 = 쉬는날  /  빨간색 = 주말\n'
                      '  주말은 자동 처리 (토글 불가)',
                 fg='#fb923c', bg='#0f1923',
                 font=('맑은 고딕', 8), justify='left'
                 ).pack(padx=6, pady=(0, 4), anchor='w')

        # ③ 계량기 추정
        self._section(parent, '③ 계량기 추정값 설정')
        self._btn(parent, '⚡ 추정값 입력', self.open_estimate_dialog,
                  color='#7c3aed').pack(fill='x', padx=6, pady=4)
        self._btn(parent, '🗑 이번 달 데이터 초기화', self.clear_month_data,
                  color='#7f1d1d').pack(fill='x', padx=6, pady=(0, 4))

        # ④ 출력 폴더
        self._section(parent, '④ 출력 폴더')
        self.out_lbl = tk.Label(parent, textvariable=self.out_dir,
                                fg='#f87171', bg='#0f1923',
                                font=('맑은 고딕', 8), anchor='w',
                                wraplength=280)
        self.out_lbl.pack(fill='x', padx=10)
        self._btn(parent, '📁 폴더 선택', self.choose_dir
                  ).pack(fill='x', padx=6, pady=4)

        # ⑤ 파일 생성
        self._section(parent, '⑤ 파일 생성')
        self._btn(parent, '📄 운영일지 생성', self.generate,
                  color='#16a34a').pack(fill='x', padx=6, pady=4)

    # ── 테이블 갱신 ──
    def refresh_table(self, save_current=True):
        """year_var / month_var 기준으로 테이블 재구성.
        save_current=True 이면 현재 rows 를 config 에 저장 후 새 월 데이터 로드.
        """
        # 현재 편집 중인 데이터 먼저 저장
        if save_current and self.rows:
            old_y = getattr(self, '_current_year', None)
            old_m = getattr(self, '_current_month', None)
            if old_y and old_m:
                save_rows(old_y, old_m, self.rows)

        y   = self.year_var.get()
        m   = self.month_var.get()
        dim = days_in_month(y, m)

        # config 에 저장된 데이터 먼저 시도
        saved = load_rows(y, m)
        saved_map = {r['day']: r for r in saved} if saved else {}

        # 현재 메모리 데이터 (같은 월을 재로드할 때 유지)
        old_map = {r['day']: r for r in self.rows
                   if getattr(self, '_current_year', None) == y
                   and getattr(self, '_current_month', None) == m}

        self.rows = []
        for day in range(1, dim + 1):
            d = date(y, m, day)
            # 우선순위: 메모리 > config 저장 > 빈값
            src = old_map.get(day) or saved_map.get(day) or {}
            self.rows.append({
                'day':        day,
                'date':       d,
                'dow':        get_dow(d),
                'is_holiday': src.get('is_holiday', False),
                'w1':         src.get('w1', ''),
                'w2':         src.get('w2', ''),
                'm1':         src.get('m1', ''),
                'm2':         src.get('m2', ''),
                'note':       src.get('note', ''),
            })

        self._current_year  = y
        self._current_month = m

        self.render_table()
        self.update_status()

    def render_table(self):
        self.tree.delete(*self.tree.get_children())
        for r in self.rows:
            d       = r['date']
            wknd    = is_wknd(d)
            holiday = r.get('is_holiday', False)

            if holiday and not wknd:
                time_str = '쉬는날'
            elif wknd:
                time_str = '토요일' if d.weekday() == 5 else '일요일'
            else:
                time_str = '08:00~17:00'

            # 쉬는날 표시: [휴] 접두사
            prefix = '[휴] ' if (holiday and not wknd) else ''
            vals = (
                prefix + f"{d.month}/{d.day}",
                r['dow'],
                time_str,
                r['w1'],
                r['w2'],
                r['m1'],
                r['m2'],
                r['note'],
            )
            tag = 'holiday' if holiday else ('weekend' if wknd else '')
            self.tree.insert('', 'end', values=vals,
                             tags=(tag,) if tag else ())

    def update_status(self):
        total    = len(self.rows)
        holidays = sum(1 for r in self.rows if r.get('is_holiday'))
        weekends = sum(1 for r in self.rows if is_wknd(r['date']))
        filled   = sum(1 for r in self.rows if r.get('m1') or r.get('m2'))
        work_days= sum(1 for r in self.rows
                       if not is_wknd(r['date']) and not r.get('is_holiday'))

        tpl_ok = '✅' if self.tpl_bytes else '❌'
        self.status.config(
            text=(f'  총 {total}일 | 근무일 {work_days} | 주말 {weekends} '
                  f'| 쉬는날 {holidays} | 계량기 입력 {filled}일 | 템플릿 {tpl_ok}'))

    def _autosave(self):
        """현재 rows 를 config 에 자동 저장"""
        y = getattr(self, '_current_year', None)
        m = getattr(self, '_current_month', None)
        if y and m:
            save_rows(y, m, self.rows)

    # ── 클릭 이벤트 ──
    def on_click(self, event):
        """날짜 열(#1) 클릭 → 쉬는날 토글"""
        col = self.tree.identify_column(event.x)
        if col != '#1':
            return
        item = self.tree.identify_row(event.y)
        if not item:
            return
        idx = self.tree.index(item)
        if idx < len(self.rows):
            r = self.rows[idx]
            # 주말은 토글 불가 (자동 처리)
            if is_wknd(r['date']):
                return
            r['is_holiday'] = not r.get('is_holiday', False)
            self.render_table()
            self.update_status()
            self._autosave()

    def on_double_click(self, event):
        """셀 더블클릭 → 인라인 편집"""
        col  = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        if not item:
            return
        idx     = self.tree.index(item)
        col_idx = int(col.replace('#', '')) - 1

        col_map = {3: 'w1', 4: 'w2', 5: 'm1', 6: 'm2', 7: 'note'}
        if col_idx not in col_map:
            return

        key  = col_map[col_idx]
        r    = self.rows[idx]
        bbox = self.tree.bbox(item, col)
        if not bbox:
            return

        entry = tk.Entry(self.tree, bg='#1e3a5f', fg='#e0e8f0',
                         font=('맑은 고딕', 9), justify='center')
        entry.insert(0, str(r.get(key, '')))
        entry.select_range(0, 'end')
        entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        entry.focus_set()

        def save_edit(e=None):
            val    = entry.get().strip()
            r[key] = val
            if key in ('m1', 'm2'):
                self._auto_calc_usage(idx)
            entry.destroy()
            self.render_table()
            self.update_status()
            self._autosave()

        entry.bind('<Return>',   save_edit)
        entry.bind('<FocusOut>', save_edit)
        entry.bind('<Escape>',   lambda e: entry.destroy())

    def _auto_calc_usage(self, idx):
        """계량기 지침 입력 시 용수사용량 자동 계산"""
        r = self.rows[idx]
        prev_m1 = prev_m2 = ''
        for i in range(idx - 1, -1, -1):
            if self.rows[i].get('m1'):
                prev_m1 = self.rows[i]['m1']
                break
        for i in range(idx - 1, -1, -1):
            if self.rows[i].get('m2'):
                prev_m2 = self.rows[i]['m2']
                break
        try:
            if r.get('m1') and prev_m1:
                r['w1'] = str(int(r['m1']) - int(prev_m1))
        except (ValueError, TypeError):
            pass
        try:
            if r.get('m2') and prev_m2:
                r['w2'] = str(int(r['m2']) - int(prev_m2))
        except (ValueError, TypeError):
            pass

    # ── 복사 / 붙여넣기 ──
    def copy_selection(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        lines = []
        for item in sel:
            vals  = self.tree.item(item, 'values')
            clean = [str(v).replace('[휴] ', '') for v in vals]
            lines.append('\t'.join(clean))
        self.root.clipboard_clear()
        self.root.clipboard_append('\n'.join(lines))

    def paste_to_table(self, event=None):
        try:
            clip = self.root.clipboard_get()
        except Exception:
            return
        lines     = [l for l in clip.strip().split('\n') if l.strip()]
        sel       = self.tree.selection()
        start_idx = self.tree.index(sel[0]) if sel else 0

        for li, line in enumerate(lines):
            idx = start_idx + li
            if idx >= len(self.rows):
                break
            parts = line.split('\t')
            r     = self.rows[idx]

            if len(parts) == 2:
                r['m1'] = parts[0].strip()
                r['m2'] = parts[1].strip()
                self._auto_calc_usage(idx)
            elif len(parts) >= 6:
                r['w1'] = parts[3].strip() if len(parts) > 3 else ''
                r['w2'] = parts[4].strip() if len(parts) > 4 else ''
                r['m1'] = parts[5].strip() if len(parts) > 5 else ''
                r['m2'] = parts[6].strip() if len(parts) > 6 else ''

        self.render_table()
        self.update_status()
        self._autosave()

    # ── 추정값 대화상자 ──
    def open_estimate_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title('계량기 추정값 설정')
        dlg.geometry('520x620')
        dlg.configure(bg='#0f1923')
        dlg.transient(self.root)
        dlg.grab_set()

        tk.Label(dlg,
                 text='시작값(이전달 마지막 계량기) / 끝값(이번달 마지막 계량기)을 입력하면\n'
                      '근무일 기준으로 하루 증가량을 자동 계산합니다.',
                 fg='#94a3b8', bg='#0f1923', font=('맑은 고딕', 9),
                 justify='left').pack(padx=10, pady=(10, 6), anchor='w')

        grid_f = tk.Frame(dlg, bg='#0f1923')
        grid_f.pack(fill='x', padx=10, pady=6)

        headers = ['설비', '시작값(이전실측)', '끝값(이번달말)', '하루증가량']
        for ci, h in enumerate(headers):
            tk.Label(grid_f, text=h, fg='#7dd3fc', bg='#0d1e33',
                     font=('맑은 고딕', 8, 'bold'), width=14, relief='ridge'
                     ).grid(row=0, column=ci, sticky='nsew', padx=1, pady=1)

        labels     = ['1호기', '2호기']
        entries    = {}
        inc_labels = {}

        cfg = load_config()
        est = cfg.get('water_estimate', {})

        for ri, lbl in enumerate(labels):
            tk.Label(grid_f, text=lbl, fg='#e0e8f0', bg='#1e3a5f',
                     font=('맑은 고딕', 9), width=14, relief='ridge'
                     ).grid(row=ri+1, column=0, sticky='nsew', padx=1, pady=1)

            for ci in range(1, 3):
                key = f'm{ri+1}_{"start" if ci==1 else "end"}'
                e = tk.Entry(grid_f, bg='#1e3a5f', fg='#e0e8f0',
                             font=('맑은 고딕', 9), width=14, justify='center')
                e.grid(row=ri+1, column=ci, sticky='nsew', padx=1, pady=1)
                e.insert(0, est.get(key, ''))
                entries[(ri, ci)] = e

            inc_lbl = tk.Label(grid_f, text='-', fg='#34d399',
                               bg='#131f2e', font=('맑은 고딕', 9),
                               width=14, relief='ridge')
            inc_lbl.grid(row=ri+1, column=3, sticky='nsew', padx=1, pady=1)
            inc_labels[ri] = inc_lbl

        # 랜덤 변동폭
        var_f = tk.Frame(dlg, bg='#0f1923')
        var_f.pack(fill='x', padx=10, pady=6)
        tk.Label(var_f, text='랜덤 변동폭 (±):', fg='#94a3b8', bg='#0f1923',
                 font=('맑은 고딕', 9)).pack(side='left')
        var_entry = tk.Entry(var_f, bg='#1e3a5f', fg='#e0e8f0',
                             font=('맑은 고딕', 9), width=8, justify='center')
        var_entry.pack(side='left', padx=4)
        var_entry.insert(0, est.get('variation', '5'))

        # 주말 증가량
        wknd_f = tk.Frame(dlg, bg='#0f1923')
        wknd_f.pack(fill='x', padx=10, pady=4)
        tk.Label(wknd_f, text='주말/쉬는날 계량기 증가:', fg='#94a3b8',
                 bg='#0f1923', font=('맑은 고딕', 9)).pack(side='left')
        wknd_entry = tk.Entry(wknd_f, bg='#1e3a5f', fg='#e0e8f0',
                              font=('맑은 고딕', 9), width=8, justify='center')
        wknd_entry.pack(side='left', padx=4)
        wknd_entry.insert(0, est.get('wknd_inc', '0'))

        # ── 추정 시작일 ──────────────────────────────────────────
        sday_f = tk.Frame(dlg, bg='#0f1923')
        sday_f.pack(fill='x', padx=10, pady=(6, 2))
        tk.Label(sday_f, text='추정 시작일:', fg='#94a3b8', bg='#0f1923',
                 font=('맑은 고딕', 9)).pack(side='left')
        start_day_var = tk.IntVar(value=int(est.get('start_day', '1')))
        tk.Spinbox(sday_f, from_=1, to=31, width=4,
                   textvariable=start_day_var,
                   bg='#1e3a5f', fg='#e0e8f0', font=('맑은 고딕', 9),
                   buttonbackground='#2e75b6').pack(side='left', padx=4)
        tk.Label(sday_f, text='일  (이날부터 추정값 적용, 이전 날은 변경 안 함)',
                 fg='#64748b', bg='#0f1923',
                 font=('맑은 고딕', 8)).pack(side='left')

        work_days_lbl = tk.Label(dlg, text='', fg='#94a3b8', bg='#0f1923',
                                  font=('맑은 고딕', 9))
        work_days_lbl.pack(padx=10, pady=2, anchor='w')
        # ─────────────────────────────────────────────────────────

        def get_wd():
            sd = start_day_var.get()
            return sum(1 for r in self.rows
                       if r['day'] >= sd
                       and not is_wknd(r['date'])
                       and not r.get('is_holiday'))

        def refresh_inc(*args):
            wd = get_wd()
            sd = start_day_var.get()
            work_days_lbl.config(text=f'{sd}일 ~ 월말  근무일: {wd}일')
            for ri in range(2):
                try:
                    s_val = float(entries[(ri, 1)].get())
                    e_val = float(entries[(ri, 2)].get())
                    diff  = e_val - s_val
                    inc   = diff / wd if wd > 0 else 0
                    inc_labels[ri].config(text=f'{inc:.1f}')
                except (ValueError, ZeroDivisionError):
                    inc_labels[ri].config(text='-')

        start_day_var.trace('w', refresh_inc)
        for e in entries.values():
            e.bind('<KeyRelease>', refresh_inc)
        refresh_inc()

        # 엑셀 붙여넣기 지원
        def paste_handler(event):
            try:
                clip = dlg.clipboard_get()
            except Exception:
                return
            lines   = [l for l in clip.strip().split('\n') if l.strip()]
            focused = dlg.focus_get()
            for (ri, ci), e in entries.items():
                if e == focused:
                    for li, line in enumerate(lines):
                        r_idx = ri + li
                        if r_idx >= 2:
                            break
                        parts = line.split('\t')
                        for pi, part in enumerate(parts):
                            c_idx = ci + pi
                            if (r_idx, c_idx) in entries:
                                entries[(r_idx, c_idx)].delete(0, 'end')
                                entries[(r_idx, c_idx)].insert(0, part.strip())
                    refresh_inc()
                    return 'break'

        for e in entries.values():
            e.bind('<Control-v>', paste_handler)

        # ── 월말 계측을 못했을 때 끝값 역산 ──────────────────────────
        tk.Frame(dlg, bg='#1e3a5f', height=1).pack(fill='x', padx=10, pady=(10, 0))

        tk.Label(dlg,
                 text='월말 계측을 못했다면? → 끝값 자동 추정',
                 fg='#fb923c', bg='#0f1923',
                 font=('맑은 고딕', 9, 'bold')
                 ).pack(anchor='w', padx=10, pady=(6, 0))
        tk.Label(dlg,
                 text='다음번 실측일·실측값을 입력하면, 이번달이 차지하는 근무일 비율로\n'
                      '끝값(이번달말 추정치)을 자동 계산해 위 끝값 칸에 채워줍니다.',
                 fg='#64748b', bg='#0f1923',
                 font=('맑은 고딕', 8), justify='left'
                 ).pack(anchor='w', padx=10, pady=(2, 4))

        cross_grid = tk.Frame(dlg, bg='#0f1923')
        cross_grid.pack(fill='x', padx=10, pady=2)

        # 다음 실측일
        tk.Label(cross_grid, text='다음 실측일 (YYYY/MM/DD):',
                 fg='#94a3b8', bg='#0f1923', font=('맑은 고딕', 9)
                 ).grid(row=0, column=0, sticky='w', pady=3)
        _ny = self.year_var.get() + (1 if self.month_var.get() == 12 else 0)
        _nm = (self.month_var.get() % 12) + 1
        cross_date_e = tk.Entry(cross_grid, bg='#1e3a5f', fg='#e0e8f0',
                                font=('맑은 고딕', 9), width=14, justify='center')
        cross_date_e.grid(row=0, column=1, sticky='w', padx=6, pady=3)
        cross_date_e.insert(0, f'{_ny}/{_nm:02d}/10')

        # 1호기·2호기 실측값
        cross_entries = {}
        for ri, label in enumerate(['1호기 실측값:', '2호기 실측값:']):
            tk.Label(cross_grid, text=label, fg='#94a3b8', bg='#0f1923',
                     font=('맑은 고딕', 9)
                     ).grid(row=ri+1, column=0, sticky='w', pady=3)
            ce = tk.Entry(cross_grid, bg='#1e3a5f', fg='#e0e8f0',
                          font=('맑은 고딕', 9), width=14, justify='center')
            ce.grid(row=ri+1, column=1, sticky='w', padx=6, pady=3)
            cross_entries[ri] = ce

        def calc_cross_end():
            """다음 실측일·값으로 이번달 끝값 역산"""
            try:
                raw  = cross_date_e.get().strip().replace('-', '/')
                pts  = raw.split('/')
                next_dt = date(int(pts[0]), int(pts[1]), int(pts[2]))
            except Exception:
                messagebox.showerror('오류', '날짜를 YYYY/MM/DD 형식으로 입력하세요.', parent=dlg)
                return

            y = self.year_var.get()
            m = self.month_var.get()
            last_day = date(y, m, days_in_month(y, m))

            if next_dt <= last_day:
                messagebox.showerror('오류',
                    f'다음 실측일은 {last_day.strftime("%Y/%m/%d")} 이후여야 합니다.',
                    parent=dlg)
                return

            # 이번달 근무일 (사용자가 설정한 쉬는날 반영)
            cur_holidays = {r['day'] for r in self.rows if r.get('is_holiday')}
            wd_this = sum(
                1 for r in self.rows
                if not is_wknd(r['date']) and r['day'] not in cur_holidays
            )

            # 다음달 근무일: 이번달 말일+1 ~ 다음 실측일 (공휴일 모르므로 주말만 제외)
            wd_next = 0
            chk = last_day + timedelta(days=1)
            while chk <= next_dt:
                if chk.weekday() < 5:
                    wd_next += 1
                chk += timedelta(days=1)

            total_wd = wd_this + wd_next
            if total_wd == 0:
                messagebox.showerror('오류', '근무일 계산 오류.', parent=dlg)
                return

            ratio = wd_this / total_wd
            results = []
            for ri in range(2):
                try:
                    sv = float(entries[(ri, 1)].get())
                    nv = float(cross_entries[ri].get())
                except ValueError:
                    messagebox.showerror('오류',
                        f'm{ri+1} 시작값 또는 실측값을 확인하세요.', parent=dlg)
                    return
                end_v = sv + ratio * (nv - sv)
                entries[(ri, 2)].delete(0, 'end')
                entries[(ri, 2)].insert(0, f'{end_v:.0f}')
                results.append(end_v)

            refresh_inc()
            messagebox.showinfo(
                '끝값 추정 완료',
                f'이번달 근무일     : {wd_this}일\n'
                f'{next_dt.strftime("%m/%d")}까지 근무일 : {wd_next}일\n'
                f'이번달 비율       : {wd_this}/{total_wd} = {ratio*100:.1f}%\n'
                f'─────────────────────\n'
                f'1호기 끝값 추정: {results[0]:,.0f}\n'
                f'2호기 끝값 추정: {results[1]:,.0f}\n\n'
                f'위 끝값 칸에 자동 입력됐습니다.',
                parent=dlg)

        self._btn(dlg, '끝값으로 추정 →', calc_cross_end,
                  color='#b45309').pack(fill='x', padx=10, pady=(4, 8))
        # ────────────────────────────────────────────────────────────

        def apply():
            sd = start_day_var.get()
            est_data = {
                'variation': var_entry.get().strip(),
                'wknd_inc':  wknd_entry.get().strip(),
                'start_day': str(sd),
            }
            for ri in range(2):
                est_data[f'm{ri+1}_start'] = entries[(ri, 1)].get().strip()
                est_data[f'm{ri+1}_end']   = entries[(ri, 2)].get().strip()
            save_config({'water_estimate': est_data})
            self._apply_estimate(est_data, start_day=sd)
            dlg.destroy()

        btn_f = tk.Frame(dlg, bg='#0f1923')
        btn_f.pack(fill='x', padx=10, pady=10)
        self._btn(btn_f, '적용',  apply,       color='#16a34a').pack(side='right', padx=4)
        self._btn(btn_f, '취소',  dlg.destroy, color='#64748b').pack(side='right', padx=4)

    def _apply_estimate(self, est, start_day=1):
        try:
            variation = int(est.get('variation', '5'))
        except ValueError:
            variation = 5
        try:
            wknd_inc = int(est.get('wknd_inc', '0'))
        except ValueError:
            wknd_inc = 0

        # start_day 이후 행만 추정 적용 (이전 행은 기존 값 유지)
        target_rows = [r for r in self.rows if r['day'] >= start_day]

        work_days = sum(1 for r in target_rows
                        if not is_wknd(r['date']) and not r.get('is_holiday'))

        for mi in range(1, 3):
            try:
                start_val = float(est.get(f'm{mi}_start', ''))
                end_val   = float(est.get(f'm{mi}_end',   ''))
            except (ValueError, TypeError):
                continue

            total_diff         = end_val - start_val
            total_wknd_days    = sum(1 for r in target_rows
                                     if is_wknd(r['date']) or r.get('is_holiday'))
            total_wknd_contrib = wknd_inc * total_wknd_days
            workday_total      = total_diff - total_wknd_contrib
            daily_inc          = workday_total / work_days if work_days > 0 else 0

            current = start_val
            for r in target_rows:          # start_day 이후만 순회
                d       = r['date']
                weekend = is_wknd(d) or r.get('is_holiday', False)
                if weekend:
                    inc = wknd_inc
                else:
                    inc = daily_inc + random.randint(-variation, variation)
                    inc = max(0, inc)
                current     += inc
                r[f'm{mi}']  = str(int(round(current)))
                r[f'w{mi}']  = str(int(round(inc)))

        self.render_table()
        self.update_status()
        self._autosave()

    def clear_month_data(self):
        """이번 달 계량기·용수 데이터를 모두 지움 (쉬는날 설정은 유지)"""
        y = self.year_var.get()
        m = self.month_var.get()
        if not messagebox.askyesno('데이터 초기화',
                f'{y}년 {m}월 계량기/용수 데이터를 모두 지우겠습니까?\n'
                '(쉬는날 설정은 유지됩니다)'):
            return
        for r in self.rows:
            r['w1'] = ''; r['w2'] = ''
            r['m1'] = ''; r['m2'] = ''
            r['note'] = ''
        self.render_table()
        self.update_status()
        self._autosave()

    # ── 출력 폴더 선택 ──
    def choose_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.out_dir.set(d)
            self.out_lbl.config(fg='#34d399')
            save_config({'out_dir': d})

    # ── 파일 생성 ──
    def generate(self):
        if not self.tpl_bytes:
            messagebox.showerror('오류',
                f'템플릿 파일을 찾을 수 없습니다.\n{TEMPLATE_PATH}')
            return

        out = self.out_dir.get()
        if not out or out == '선택 안 됨' or not os.path.isdir(out):
            messagebox.showerror('오류', '출력 폴더를 먼저 선택하세요.')
            return

        y = self.year_var.get()
        m = self.month_var.get()
        dim = days_in_month(y, m)

        # ── 생성 전 데이터 현황 확인 ────────────────────────────────
        total_days   = len(self.rows)
        work_days    = sum(1 for r in self.rows
                          if not is_wknd(r['date']) and not r.get('is_holiday'))
        filled_days  = sum(1 for r in self.rows if r.get('m1') or r.get('m2'))
        empty_work   = sum(1 for r in self.rows
                          if not is_wknd(r['date']) and not r.get('is_holiday')
                          and not r.get('m1') and not r.get('m2'))

        # 빈 근무일이 있으면 경고
        if empty_work > 0:
            empty_days_list = [str(r['day']) for r in self.rows
                               if not is_wknd(r['date']) and not r.get('is_holiday')
                               and not r.get('m1') and not r.get('m2')]
            days_str = ', '.join(empty_days_list[:10])
            if len(empty_days_list) > 10:
                days_str += f' 외 {len(empty_days_list)-10}일'

            msg = (f'⚠️  계량기 값이 비어있는 근무일이 {empty_work}일 있습니다.\n\n'
                   f'빈 날: {days_str}\n\n'
                   f'─────────────────────────────\n'
                   f'총 {total_days}일 | 근무일 {work_days}일\n'
                   f'계량기 입력된 날: {filled_days}일\n'
                   f'계량기 비어있는 근무일: {empty_work}일\n\n'
                   f'이대로 생성하시겠습니까?\n'
                   f'(빈 근무일은 Word 파일에 공백으로 출력됩니다)')
            if not messagebox.askyesno('데이터 확인', msg):
                return
        # ────────────────────────────────────────────────────────────

        daily_data = {}
        holidays   = set()
        for r in self.rows:
            day = r['day']
            if r.get('is_holiday'):
                holidays.add(day)
            daily_data[day] = {
                'w1': r.get('w1', ''),
                'w2': r.get('w2', ''),
                'm1': r.get('m1', ''),
                'm2': r.get('m2', ''),
            }

        try:
            result = fill_water_docx(self.tpl_bytes, y, m, daily_data, holidays)
            yy     = y % 100
            fname  = f'폐수배출시설_운영일지({yy}년{m}월).docx'
            fpath  = os.path.join(out, fname)
            with open(fpath, 'wb') as f:
                f.write(result)
            messagebox.showinfo('완료',
                f'생성 완료!\n{fpath}\n\n'
                f'계량기 입력: {filled_days}/{total_days}일')
        except Exception as e:
            messagebox.showerror('오류', f'생성 중 오류:\n{e}')


if __name__ == '__main__':
    root = tk.Tk()
    app  = WaterApp(root)
    root.mainloop()
