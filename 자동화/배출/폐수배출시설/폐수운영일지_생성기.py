"""
폐수배출시설 운영일지 자동 생성기
- 월별 docx 하나에 일별 데이터 기록
- 1호기·2호기 계량기 지침 입력 → 용수사용량 자동 계산
- 주말/쉬는날 자동 처리
- 템플릿(폐수배출시설_운영일지_template.docx) 기반 XML 치환
- Python 기본 라이브러리만 사용
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import zipfile, os, re, io, json, random, calendar
import xml.etree.ElementTree as ET
from datetime import date, timedelta

# ── 경로 ──
_HERE        = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH  = os.path.join(_HERE, 'water_config.json')
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

# ── 상수 ──
DOW_KR = ['월','화','수','목','금','토','일']
REPRESENTATIVE = '김예배'
ENV_TECH       = '신종한'

COLS = ('날짜','요일','가동시간','1호기용수(L)','2호기용수(L)',
        '1호기계량기','2호기계량기','비고')
COL_WIDTHS = [78, 38, 90, 82, 82, 90, 90, 100]

# ── 유틸 ──
def is_wknd(d):  return d.weekday() >= 5
def get_dow(d):  return DOW_KR[d.weekday()]
def days_in_month(y, m): return calendar.monthrange(y, m)[1]

# ── XML 네임스페이스 ──
NS = {
    'w':    'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14':  'http://schemas.microsoft.com/office/word/2010/wordml',
    'r':    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc':   'http://schemas.openxmlformats.org/markup-compatibility/2006',
}

def _register_ns():
    """ElementTree 직렬화 시 네임스페이스 접두사 유지"""
    # 전체 네임스페이스 목록 (원본 docx에 있는 모든 것)
    all_ns = {
        'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
        'cx':  'http://schemas.microsoft.com/office/drawing/2014/chartex',
        'cx1': 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex',
        'cx2': 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex',
        'cx3': 'http://schemas.microsoft.com/office/drawing/2016/5/9/chartex',
        'cx4': 'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex',
        'cx5': 'http://schemas.microsoft.com/office/drawing/2016/5/11/chartex',
        'cx6': 'http://schemas.microsoft.com/office/drawing/2016/5/12/chartex',
        'cx7': 'http://schemas.microsoft.com/office/drawing/2016/5/13/chartex',
        'cx8': 'http://schemas.microsoft.com/office/drawing/2016/5/14/chartex',
        'mc':  'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'aink':'http://schemas.microsoft.com/office/drawing/2016/ink',
        'am3d':'http://schemas.microsoft.com/office/drawing/2017/model3d',
        'o':   'urn:schemas-microsoft-com:office:office',
        'oel': 'http://schemas.microsoft.com/office/2019/extlst',
        'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'm':   'http://schemas.openxmlformats.org/officeDocument/2006/math',
        'v':   'urn:schemas-microsoft-com:vml',
        'wp14':'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
        'wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'w10': 'urn:schemas-microsoft-com:office:word',
        'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
        'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
        'w16cex':'http://schemas.microsoft.com/office/word/2018/wordml/cex',
        'w16cid':'http://schemas.microsoft.com/office/word/2016/wordml/cid',
        'w16': 'http://schemas.microsoft.com/office/word/2018/wordml',
        'w16du':'http://schemas.microsoft.com/office/word/2023/wordml/word16du',
        'w16sdtdh':'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash',
        'w16sdtfl':'http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock',
        'w16se':'http://schemas.microsoft.com/office/word/2015/wordml/symex',
        'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
        'wpi': 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
        'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
        'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    }
    for prefix, uri in all_ns.items():
        ET.register_namespace(prefix, uri)

_register_ns()

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

def set_cell_text_et(tc_elem, new_text):
    """ElementTree 셀 요소 내 모든 w:t 텍스트를 new_text로 교체"""
    wt_elems = tc_elem.findall(f'.//{W}t')
    if wt_elems:
        wt_elems[0].text = new_text
        for wt in wt_elems[1:]:
            wt.text = ''
    else:
        # w:t가 없으면 첫 번째 w:p에 w:r > w:t 추가
        wp = tc_elem.find(f'.//{W}p')
        if wp is not None:
            wr = ET.SubElement(wp, f'{W}r')
            wt = ET.SubElement(wr, f'{W}t')
            wt.text = new_text

def get_cell_text_et(tc_elem):
    """셀의 모든 w:t 텍스트를 합쳐서 반환"""
    return ''.join(t.text or '' for t in tc_elem.findall(f'.//{W}t'))

def fill_water_docx(tpl_bytes, year, month, daily_data, holidays=set()):
    """템플릿 docx를 기반으로 월별 운영일지를 생성."""
    dim = days_in_month(year, month)
    buf_in = io.BytesIO(tpl_bytes)
    buf_out = io.BytesIO()

    with zipfile.ZipFile(buf_in, 'r') as src, \
         zipfile.ZipFile(buf_out, 'w', zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            raw = src.read(item.filename)
            if item.filename == 'word/document.xml':
                xml_str = raw.decode('utf-8')
                xml_str = _fill_document_xml(xml_str, year, month, dim, daily_data, holidays)
                raw = xml_str.encode('utf-8')
            dst.writestr(item, raw)

    return buf_out.getvalue()

def _fill_document_xml(xml_str, year, month, dim, daily_data, holidays):
    """document.xml 내 2개 테이블의 데이터 행을 채움 (ElementTree 사용)"""
    root = ET.fromstring(xml_str)
    body = root.find(f'{W}body')
    if body is None:
        return xml_str

    tables = body.findall(f'{W}tbl')
    if len(tables) < 2:
        return xml_str

    yy = year % 100

    for tidx, tbl in enumerate(tables):
        rows = tbl.findall(f'{W}tr')
        header_count = 5
        footer_count = 1
        data_start = header_count
        data_end = len(rows) - footer_count

        if tidx == 0:
            day_start = 1
        else:
            day_start = 16

        for row_i in range(data_start, data_end):
            day_offset = row_i - data_start
            day_num = day_start + day_offset

            tr = rows[row_i]
            cells = tr.findall(f'{W}tc')
            if len(cells) < 10:
                continue

            if day_num > dim:
                # 해당 월에 없는 날 → 전부 비움
                for ci in range(len(cells)):
                    set_cell_text_et(cells[ci], '')
            else:
                d = date(year, month, day_num)
                weekend = is_wknd(d) or (day_num in holidays)
                dd = daily_data.get(day_num, {})

                # 셀 0: 연월일
                if row_i == data_start:
                    date_str = f'{yy}/{month:02d}/{day_num:02d}'
                else:
                    date_str = str(day_num)
                set_cell_text_et(cells[0], date_str)

                # 셀 1-2: 결재
                if row_i == data_start:
                    set_cell_text_et(cells[1], REPRESENTATIVE)
                    set_cell_text_et(cells[2], ENV_TECH)
                else:
                    set_cell_text_et(cells[1], '\u201c')
                    set_cell_text_et(cells[2], '\u201c')

                # 셀 3-5: 가동시간
                if weekend:
                    if day_num in holidays and d.weekday() < 5:
                        dow_name = '쉬는날'
                    elif d.weekday() == 5:
                        dow_name = '토요일'
                    else:
                        dow_name = '일요일'
                    set_cell_text_et(cells[3], dow_name)
                    if len(cells) > 4: set_cell_text_et(cells[4], '')
                    if len(cells) > 5: set_cell_text_et(cells[5], '')
                else:
                    set_cell_text_et(cells[3], '08:00~1')
                    if len(cells) > 4: set_cell_text_et(cells[4], '7')
                    if len(cells) > 5: set_cell_text_et(cells[5], ':00')

                # 셀 6-7: 용수사용량
                w1 = dd.get('w1', '0' if weekend else '')
                w2 = dd.get('w2', '0' if weekend else '')
                set_cell_text_et(cells[6], str(w1))
                set_cell_text_et(cells[7], str(w2))

                # 셀 8-9: 계량기지침
                m1 = dd.get('m1', '')
                m2 = dd.get('m2', '')
                set_cell_text_et(cells[8], str(m1))
                set_cell_text_et(cells[9], str(m2))

                # 셀 10+: 비움
                for ci in range(10, len(cells)):
                    set_cell_text_et(cells[ci], '')

    # XML 직렬화
    out = io.StringIO()
    tree = ET.ElementTree(root)
    tree.write(out, encoding='unicode', xml_declaration=True)
    result = out.getvalue()

    # mc:Ignorable 속성 복원 (ElementTree가 날릴 수 있음)
    if 'mc:Ignorable' not in result:
        result = result.replace(
            f'<w:document ',
            '<w:document mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14" ',
            1
        )

    return result


# ══════════════════════════════════════════════
# GUI
# ══════════════════════════════════════════════
class WaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title('폐수배출시설 운영일지 생성기')
        self.root.geometry('980x700')
        self.root.configure(bg='#0f1923')
        self.root.resizable(True, True)

        self.tpl_bytes = None
        self.out_dir   = tk.StringVar(value='선택 안 됨')
        self.rows: list[dict] = []

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

    def _load_template(self):
        if os.path.exists(TEMPLATE_PATH):
            try:
                with open(TEMPLATE_PATH, 'rb') as f:
                    self.tpl_bytes = f.read()
            except Exception:
                self.tpl_bytes = None

    def _setup_style(self):
        s = ttk.Style()
        s.theme_use('default')
        s.configure('Treeview',
            background='#131f2e', foreground='#e0e8f0',
            fieldbackground='#131f2e', rowheight=23, font=('맑은 고딕',9))
        s.configure('Treeview.Heading',
            background='#0d1e33', foreground='#7dd3fc', font=('맑은 고딕',9,'bold'))
        s.map('Treeview', background=[('selected','#1e3a5f')])
        s.configure('Vertical.TScrollbar',  background='#1e3a5f', troughcolor='#0d1e33')
        s.configure('Horizontal.TScrollbar', background='#1e3a5f', troughcolor='#0d1e33')

    def _btn(self, parent, text, cmd, color='#2e75b6', fg='white', **kw):
        b = tk.Button(parent, text=text, command=cmd,
                      bg=color, fg=fg, font=('맑은 고딕',9,'bold'),
                      relief='flat', cursor='hand2', **kw)
        def lighten(c):
            r,g,b2 = int(c[1:3],16), int(c[3:5],16), int(c[5:7],16)
            return '#{:02x}{:02x}{:02x}'.format(min(r+25,255), min(g+25,255), min(b2+25,255))
        b.bind('<Enter>', lambda e: b.config(bg=lighten(color)))
        b.bind('<Leave>', lambda e: b.config(bg=color))
        return b

    def _section(self, parent, title):
        lbl = tk.Label(parent, text=title, fg='#7dd3fc', bg='#0f1923',
                       font=('맑은 고딕',9,'bold'), anchor='w')
        lbl.pack(fill='x', pady=(10,4), padx=6)
        sep = tk.Frame(parent, bg='#1e3a5f', height=1)
        sep.pack(fill='x', padx=6, pady=(0,6))

    def _build_ui(self):
        main = tk.PanedWindow(self.root, orient='horizontal', bg='#0f1923',
                              sashwidth=4, sashrelief='flat')
        main.pack(fill='both', expand=True)

        # 왼쪽: 테이블
        left = tk.Frame(main, bg='#0f1923')
        main.add(left, width=660)
        self._build_table(left)

        # 오른쪽: 사이드바
        right_canvas = tk.Canvas(main, bg='#0f1923', highlightthickness=0)
        main.add(right_canvas, width=310)
        sb_scroll = ttk.Scrollbar(right_canvas, orient='vertical', command=right_canvas.yview)
        sb_frame = tk.Frame(right_canvas, bg='#0f1923')
        sb_frame.bind('<Configure>', lambda e: right_canvas.configure(scrollregion=right_canvas.bbox('all')))
        right_canvas.create_window((0,0), window=sb_frame, anchor='nw')
        right_canvas.configure(yscrollcommand=sb_scroll.set)
        sb_scroll.pack(side='right', fill='y')
        right_canvas.pack(side='left', fill='both', expand=True)
        right_canvas.bind_all('<MouseWheel>',
            lambda e: right_canvas.yview_scroll(int(-1*(e.delta/120)), 'units'))

        self._build_sidebar(sb_frame)

        # 상태바
        self.status = tk.Label(self.root, text='', bg='#0a1018', fg='#64748b',
                               font=('맑은 고딕',8), anchor='w', padx=8)
        self.status.pack(fill='x', side='bottom')

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
        self.tree.bind('<Double-1>', self.on_double_click)
        self.tree.bind('<Control-c>', self.copy_selection)
        self.tree.bind('<Control-v>', self.paste_to_table)

        self.tree.tag_configure('weekend', foreground='#f87171')
        self.tree.tag_configure('holiday', foreground='#fb923c')

    def _build_sidebar(self, parent):
        # ① 년월 설정
        self._section(parent, '① 년월 설정')

        row_f = tk.Frame(parent, bg='#0f1923')
        row_f.pack(fill='x', padx=6, pady=2)
        tk.Label(row_f, text='년:', fg='#94a3b8', bg='#0f1923',
                 font=('맑은 고딕',9)).pack(side='left')
        tk.Spinbox(row_f, from_=2020, to=2040, width=6,
                   textvariable=self.year_var,
                   bg='#1e3a5f', fg='#e0e8f0', font=('맑은 고딕',9),
                   buttonbackground='#2e75b6').pack(side='left', padx=(2,10))

        tk.Label(row_f, text='월:', fg='#94a3b8', bg='#0f1923',
                 font=('맑은 고딕',9)).pack(side='left')
        tk.Spinbox(row_f, from_=1, to=12, width=4,
                   textvariable=self.month_var,
                   bg='#1e3a5f', fg='#e0e8f0', font=('맑은 고딕',9),
                   buttonbackground='#2e75b6').pack(side='left', padx=2)

        self._btn(parent, '📅 테이블 새로고침', self.refresh_table
                  ).pack(fill='x', padx=6, pady=4)

        # ② 계량기 추정
        self._section(parent, '② 계량기 추정값 설정')

        self._btn(parent, '⚡ 추정값 입력', self.open_estimate_dialog,
                  color='#7c3aed').pack(fill='x', padx=6, pady=4)

        # ③ 출력 폴더
        self._section(parent, '③ 출력 폴더')

        self.out_lbl = tk.Label(parent, textvariable=self.out_dir,
                                fg='#f87171', bg='#0f1923',
                                font=('맑은 고딕',8), anchor='w', wraplength=280)
        self.out_lbl.pack(fill='x', padx=10)

        self._btn(parent, '📁 폴더 선택', self.choose_dir
                  ).pack(fill='x', padx=6, pady=4)

        # ④ 파일 생성
        self._section(parent, '④ 파일 생성')

        self._btn(parent, '📄 운영일지 생성', self.generate,
                  color='#16a34a').pack(fill='x', padx=6, pady=4)

    # ── 테이블 ──
    def refresh_table(self):
        y = self.year_var.get()
        m = self.month_var.get()
        dim = days_in_month(y, m)

        old_map = {r['day']: r for r in self.rows}

        self.rows = []
        for day in range(1, dim + 1):
            d = date(y, m, day)
            old = old_map.get(day, {})
            self.rows.append({
                'day': day,
                'date': d,
                'dow': get_dow(d),
                'is_holiday': old.get('is_holiday', False),
                'w1': old.get('w1', ''),
                'w2': old.get('w2', ''),
                'm1': old.get('m1', ''),
                'm2': old.get('m2', ''),
                'note': old.get('note', ''),
            })

        self.render_table()
        self.update_status()

    def render_table(self):
        self.tree.delete(*self.tree.get_children())
        for r in self.rows:
            d = r['date']
            wknd = is_wknd(d)
            holiday = r.get('is_holiday', False)

            if holiday and not wknd:
                time_str = '쉬는날'
            elif wknd:
                time_str = '토요일' if d.weekday() == 5 else '일요일'
            else:
                time_str = '08:00~17:00'

            prefix = '🚫 ' if holiday else ''
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
            self.tree.insert('', 'end', values=vals, tags=(tag,) if tag else ())

    def update_status(self):
        total = len(self.rows)
        holidays = sum(1 for r in self.rows if r.get('is_holiday'))
        weekends = sum(1 for r in self.rows if is_wknd(r['date']))
        filled = sum(1 for r in self.rows if r.get('m1') or r.get('m2'))
        work_days = sum(1 for r in self.rows
                        if not is_wknd(r['date']) and not r.get('is_holiday'))

        tpl_ok = '✅' if self.tpl_bytes else '❌'
        self.status.config(
            text=f'  총 {total}일 | 근무일 {work_days} | 주말 {weekends} | 쉬는날 {holidays} '
                 f'| 입력 {filled}일 | 템플릿 {tpl_ok}')

    # ── 클릭 이벤트 ──
    def on_click(self, event):
        col = self.tree.identify_column(event.x)
        if col != '#1':
            return
        item = self.tree.identify_row(event.y)
        if not item:
            return
        idx = self.tree.index(item)
        if idx < len(self.rows):
            r = self.rows[idx]
            r['is_holiday'] = not r.get('is_holiday', False)
            self.render_table()
            self.update_status()

    def on_double_click(self, event):
        col = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        if not item:
            return
        idx = self.tree.index(item)
        col_idx = int(col.replace('#','')) - 1

        col_map = {3: 'w1', 4: 'w2', 5: 'm1', 6: 'm2', 7: 'note'}
        if col_idx not in col_map:
            return

        key = col_map[col_idx]
        r = self.rows[idx]

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
            val = entry.get().strip()
            r[key] = val
            if key in ('m1', 'm2'):
                self._auto_calc_usage(idx)
            entry.destroy()
            self.render_table()
            self.update_status()

        entry.bind('<Return>', save_edit)
        entry.bind('<FocusOut>', save_edit)
        entry.bind('<Escape>', lambda e: entry.destroy())

    def _auto_calc_usage(self, idx):
        """계량기 지침 변경 시 용수사용량 자동 계산"""
        r = self.rows[idx]
        prev_m1, prev_m2 = '', ''
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

    # ── 복사/붙여넣기 ──
    def copy_selection(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        lines = []
        for item in sel:
            vals = self.tree.item(item, 'values')
            clean = [str(v).replace('🚫 ', '') for v in vals]
            lines.append('\t'.join(clean))
        self.root.clipboard_clear()
        self.root.clipboard_append('\n'.join(lines))

    def paste_to_table(self, event=None):
        try:
            clip = self.root.clipboard_get()
        except Exception:
            return
        lines = [l for l in clip.strip().split('\n') if l.strip()]
        sel = self.tree.selection()
        start_idx = self.tree.index(sel[0]) if sel else 0

        for li, line in enumerate(lines):
            idx = start_idx + li
            if idx >= len(self.rows):
                break
            parts = line.split('\t')
            r = self.rows[idx]

            if len(parts) == 2:
                # 계량기지침 2열만
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

    # ── 추정값 대화상자 ──
    def open_estimate_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title('계량기 추정값 설정')
        dlg.geometry('520x400')
        dlg.configure(bg='#0f1923')
        dlg.transient(self.root)
        dlg.grab_set()

        tk.Label(dlg, text='시작값(이전달 마지막 계량기) / 끝값(이번달 마지막 계량기)을 입력하면\n'
                           '근무일 기준으로 하루 증가량을 자동 계산합니다.',
                 fg='#94a3b8', bg='#0f1923', font=('맑은 고딕',9),
                 justify='left').pack(padx=10, pady=(10,6), anchor='w')

        grid_f = tk.Frame(dlg, bg='#0f1923')
        grid_f.pack(fill='x', padx=10, pady=6)

        headers = ['설비', '시작값(이전실측)', '끝값(이번달말)', '하루증가량']
        for ci, h in enumerate(headers):
            tk.Label(grid_f, text=h, fg='#7dd3fc', bg='#0d1e33',
                     font=('맑은 고딕',8,'bold'), width=14, relief='ridge'
                     ).grid(row=0, column=ci, sticky='nsew', padx=1, pady=1)

        labels = ['1호기', '2호기']
        entries = {}
        inc_labels = {}

        cfg = load_config()
        est = cfg.get('water_estimate', {})

        for ri, lbl in enumerate(labels):
            tk.Label(grid_f, text=lbl, fg='#e0e8f0', bg='#1e3a5f',
                     font=('맑은 고딕',9), width=14, relief='ridge'
                     ).grid(row=ri+1, column=0, sticky='nsew', padx=1, pady=1)

            for ci in range(1, 3):
                key = f'm{ri+1}_{"start" if ci==1 else "end"}'
                e = tk.Entry(grid_f, bg='#1e3a5f', fg='#e0e8f0',
                             font=('맑은 고딕',9), width=14, justify='center')
                e.grid(row=ri+1, column=ci, sticky='nsew', padx=1, pady=1)
                e.insert(0, est.get(key, ''))
                entries[(ri, ci)] = e

            inc_lbl = tk.Label(grid_f, text='-', fg='#34d399', bg='#131f2e',
                               font=('맑은 고딕',9), width=14, relief='ridge')
            inc_lbl.grid(row=ri+1, column=3, sticky='nsew', padx=1, pady=1)
            inc_labels[ri] = inc_lbl

        # 랜덤 변동폭
        var_f = tk.Frame(dlg, bg='#0f1923')
        var_f.pack(fill='x', padx=10, pady=6)
        tk.Label(var_f, text='랜덤 변동폭 (±):', fg='#94a3b8', bg='#0f1923',
                 font=('맑은 고딕',9)).pack(side='left')
        var_entry = tk.Entry(var_f, bg='#1e3a5f', fg='#e0e8f0',
                             font=('맑은 고딕',9), width=8, justify='center')
        var_entry.pack(side='left', padx=4)
        var_entry.insert(0, est.get('variation', '5'))

        # 주말 증가량
        wknd_f = tk.Frame(dlg, bg='#0f1923')
        wknd_f.pack(fill='x', padx=10, pady=4)
        tk.Label(wknd_f, text='주말/쉬는날 계량기 증가:', fg='#94a3b8', bg='#0f1923',
                 font=('맑은 고딕',9)).pack(side='left')
        wknd_entry = tk.Entry(wknd_f, bg='#1e3a5f', fg='#e0e8f0',
                              font=('맑은 고딕',9), width=8, justify='center')
        wknd_entry.pack(side='left', padx=4)
        wknd_entry.insert(0, est.get('wknd_inc', '0'))

        work_days = sum(1 for r in self.rows
                        if not is_wknd(r['date']) and not r.get('is_holiday'))
        tk.Label(dlg, text=f'이번달 근무일: {work_days}일',
                 fg='#94a3b8', bg='#0f1923', font=('맑은 고딕',9)
                 ).pack(padx=10, pady=2, anchor='w')

        def refresh_inc(*args):
            for ri in range(2):
                try:
                    s = float(entries[(ri, 1)].get())
                    e_val = float(entries[(ri, 2)].get())
                    diff = e_val - s
                    inc = diff / work_days if work_days > 0 else 0
                    inc_labels[ri].config(text=f'{inc:.1f}')
                except (ValueError, ZeroDivisionError):
                    inc_labels[ri].config(text='-')

        for e in entries.values():
            e.bind('<KeyRelease>', refresh_inc)
        refresh_inc()

        # 엑셀 붙여넣기
        def paste_handler(event):
            try:
                clip = dlg.clipboard_get()
            except Exception:
                return
            lines = [l for l in clip.strip().split('\n') if l.strip()]
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

        def apply():
            est_data = {
                'variation': var_entry.get().strip(),
                'wknd_inc': wknd_entry.get().strip(),
            }
            for ri in range(2):
                est_data[f'm{ri+1}_start'] = entries[(ri, 1)].get().strip()
                est_data[f'm{ri+1}_end'] = entries[(ri, 2)].get().strip()

            save_config({'water_estimate': est_data})
            self._apply_estimate(est_data)
            dlg.destroy()

        btn_f = tk.Frame(dlg, bg='#0f1923')
        btn_f.pack(fill='x', padx=10, pady=10)
        self._btn(btn_f, '적용', apply, color='#16a34a').pack(side='right', padx=4)
        self._btn(btn_f, '취소', dlg.destroy, color='#64748b').pack(side='right', padx=4)

    def _apply_estimate(self, est):
        try:
            variation = int(est.get('variation', '5'))
        except ValueError:
            variation = 5
        try:
            wknd_inc = int(est.get('wknd_inc', '0'))
        except ValueError:
            wknd_inc = 0

        work_days = sum(1 for r in self.rows
                        if not is_wknd(r['date']) and not r.get('is_holiday'))

        for mi in range(1, 3):
            try:
                start_val = float(est.get(f'm{mi}_start', ''))
                end_val   = float(est.get(f'm{mi}_end', ''))
            except (ValueError, TypeError):
                continue

            total_diff = end_val - start_val
            total_wknd_days = sum(1 for r in self.rows
                                  if is_wknd(r['date']) or r.get('is_holiday'))
            total_wknd_contribution = wknd_inc * total_wknd_days
            workday_total = total_diff - total_wknd_contribution
            daily_inc = workday_total / work_days if work_days > 0 else 0

            current = start_val
            for r in self.rows:
                d = r['date']
                weekend = is_wknd(d) or r.get('is_holiday', False)

                if weekend:
                    inc = wknd_inc
                else:
                    inc = daily_inc + random.randint(-variation, variation)
                    inc = max(0, inc)

                current += inc
                r[f'm{mi}'] = str(int(round(current)))
                r[f'w{mi}'] = str(int(round(inc)))

        self.render_table()
        self.update_status()

    # ── 출력 폴더 ──
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

        daily_data = {}
        holidays = set()
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
            yy = y % 100
            fname = f'폐수배출시설_운영일지({yy}년{m}월).docx'
            fpath = os.path.join(out, fname)
            with open(fpath, 'wb') as f:
                f.write(result)
            messagebox.showinfo('완료', f'생성 완료!\n{fpath}')
        except Exception as e:
            messagebox.showerror('오류', f'생성 중 오류:\n{e}')


if __name__ == '__main__':
    root = tk.Tk()
    app = WaterApp(root)
    root.mainloop()
