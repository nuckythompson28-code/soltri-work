"""
배출방지시설 운영기록부 자동 생성기 (Python 버전)
- 브라우저 다운로드 없이 직접 파일 생성 → Windows 보안 경고 없음
- Python 기본 라이브러리만 사용 (별도 설치 불필요)
- 기상청 공공데이터 API 연동으로 날씨·기온 자동 입력
- template.docx 를 이 스크립트와 같은 폴더에 두면 자동 인식
- 실행: python 운영기록부_생성기.py  또는 더블클릭
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import zipfile, os, re, io, random, json, threading
from datetime import date, timedelta
from urllib.request import urlopen
from urllib.parse import urlencode, quote_plus
from urllib.error import URLError

# ── 경로 (스크립트와 같은 폴더 기준) ──
_HERE        = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH  = os.path.join(_HERE, 'config.json')
TEMPLATE_PATH = os.path.join(_HERE, 'template.docx')

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
DOW_KR      = ['월', '화', '수', '목', '금', '토', '일']
WEATHER_OPT = ['맑음', '흐림', '구름조금', '비', '눈']
TEMP_RANGE  = {
    1:(-5,2),2:(-3,5),3:(2,11),4:(8,18),5:(14,24),6:(19,27),
    7:(23,30),8:(24,31),9:(18,26),10:(11,20),11:(4,12),12:(-2,4)
}
COLS = ('날짜','요일','전력1(VAL1)','전력2(VAL2)','전력3(VAL3)','전력4(VAL4)',
        '전력5(VAL5)','전력6(VAL6)','전력7(VAL7)','전력8(VAL8)','날씨','기온(℃)')
COL_WIDTHS = [88,36,82,82,82,82,82,82,82,82,68,58]

# 기상청 관측소 목록 (주요 지점)
STATIONS = {
    '서울 (108)':      '108',
    '부산 (159)':      '159',
    '대구 (143)':      '143',
    '인천 (112)':      '112',
    '광주 (156)':      '156',
    '대전 (133)':      '133',
    '울산 (152)':      '152',
    '수원 (119)':      '119',
    '춘천 (101)':      '101',
    '강릉 (105)':      '105',
    '청주 (131)':      '131',
    '전주 (146)':      '146',
    '포항 (138)':      '138',
    '창원 (155)':      '155',
    '목포 (165)':      '165',
    '여수 (168)':      '168',
    '제주 (184)':      '184',
    '고창 (172)':      '172',
    '군산 (140)':      '140',
    '안동 (136)':      '136',
}

# ── 유틸 ──
def rand_temp(m):   mn,mx=TEMP_RANGE[m]; return random.randint(mn,mx)
def rand_weather(): return random.choice(['맑음','맑음','맑음','흐림','구름조금'])
def is_wknd(d):     return d.weekday()>=5
def get_dow(d):     return DOW_KR[d.weekday()]

def replace_in_xml(xml, ph, val):
    return re.sub(
        r'(<w:t(?:[^>]*)>)([\s\S]*?)(</w:t>)',
        lambda m: m.group(1)+m.group(2).replace(ph,val)+m.group(3),
        xml
    )

def fill_docx(tpl_bytes, row):
    d    = row['date']
    wknd = is_wknd(d) or row.get('is_holiday', False)
    time17 = '' if wknd else '07:00 ~ 12:00, 13:00~17:00'
    time18 = '' if wknd else '07:00 ~ 12:00, 13:00~18:00'
    rep = [
        ('YYYY',str(d.year)),('MM',str(d.month)),('DD',str(d.day)),
        ('WD',get_dow(d)),('WTH',row['weather'] or '맑음'),('TMP',row['temp'] or ''),
        ('TIME17',time17),('TIME18',time18),
        *[(f'VAL{i+1}', row['power'][i] or '') for i in range(8)],
    ]
    xml_targets = ['word/document.xml','word/header1.xml','word/footer1.xml',
                   'word/header2.xml','word/footer2.xml','word/header3.xml','word/footer3.xml']
    buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(tpl_bytes),'r') as src, \
         zipfile.ZipFile(buf,'w',zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename in xml_targets:
                xml = data.decode('utf-8')
                for ph,val in rep:
                    xml = replace_in_xml(xml, ph, val)
                data = xml.encode('utf-8')
            dst.writestr(item, data)
    return buf.getvalue()

# ── 기상청 API ──
def weather_from_asos(avg_ta, sum_rn, avg_tca):
    try: rn  = float(sum_rn)  if sum_rn  else 0.0
    except: rn  = 0.0
    try: tca = float(avg_tca) if avg_tca else -1
    except: tca = -1
    try: ta  = float(avg_ta)  if avg_ta  else 999
    except: ta  = 999

    if rn > 0:
        return '눈' if ta <= 0 else '비'
    if tca < 0:
        return '맑음'
    if tca <= 2:  return '맑음'
    if tca <= 5:  return '구름조금'
    return '흐림'

def fetch_asos(api_key, stn_id, start_dt, end_dt):
    base = 'http://apis.data.go.kr/1360000/AsosDalyInfoService/getWthrDataList'
    params = {
        'serviceKey': api_key,
        'pageNo':     '1',
        'numOfRows':  '999',
        'dataType':   'JSON',
        'dataCd':     'ASOS',
        'dateCd':     'DAY',
        'startDt':    start_dt.replace('-',''),
        'endDt':      end_dt.replace('-',''),
        'stnIds':     stn_id,
    }
    url = base + '?' + urlencode(params, quote_via=quote_plus)
    with urlopen(url, timeout=15) as resp:
        raw  = resp.read().decode('utf-8')
    data = json.loads(raw)

    items = (data.get('response',{})
                 .get('body',{})
                 .get('items',{})
                 .get('item',[]))

    result = {}
    for it in items:
        tm  = it.get('tm','')
        ta  = it.get('avgTa','')
        rn  = it.get('sumRn','')
        tca = it.get('avgTca','')
        if not tm: continue
        result[tm] = {
            'temp':    str(round(float(ta))) if ta else '',
            'weather': weather_from_asos(ta, rn, tca),
        }
    return result


# ══════════════════════════════════════════════
# GUI
# ══════════════════════════════════════════════
class App:
    def __init__(self, root):
        self.root = root
        self.root.title('배출방지시설 운영기록부 자동 생성기 (Python)')
        self.root.geometry('1160x740')
        self.root.configure(bg='#0f1923')
        self.root.resizable(True, True)

        self.tpl_bytes  = None
        self.out_dir    = tk.StringVar(value='선택 안 됨')
        self.start_var  = tk.StringVar(value='2026-01-01')
        self.end_var    = tk.StringVar(value='2026-01-31')
        self.skip_empty = tk.BooleanVar(value=True)
        self.skip_wknd  = tk.BooleanVar(value=False)
        self.temp_var   = tk.StringVar()
        self.api_key    = tk.StringVar()
        self.station    = tk.StringVar(value='서울 (108)')
        self.rows: list[dict] = []

        # template.docx 자동 로드
        self._load_template()

        # 저장된 설정 불러오기
        cfg = load_config()
        if cfg.get('api_key'):
            self.api_key.set(cfg['api_key'])
        if cfg.get('station'):
            self.station.set(cfg['station'])
        if cfg.get('out_dir'):
            self.out_dir.set(cfg['out_dir'])

        self._setup_style()
        self._build_ui()

        # 저장된 출력 폴더 라벨 색 적용
        if cfg.get('out_dir'):
            self.out_lbl.config(fg='#34d399')

        # API 키·관측소 변경 시 자동 저장
        self.api_key.trace_add('write', lambda *_: save_config({'api_key': self.api_key.get()}))
        self.station.trace_add('write', lambda *_: save_config({'station': self.station.get()}))

        # 초기 상태 업데이트
        self.update_status()

    def _load_template(self):
        """스크립트 폴더의 template.docx 자동 로드"""
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
        s.configure('Horizontal.TScrollbar',background='#1e3a5f', troughcolor='#0d1e33')
        s.configure('TCombobox', fieldbackground='#1e3a5f', background='#1e3a5f',
                    foreground='#e0e8f0', selectbackground='#2e75b6')

    def _btn(self, parent, text, cmd, color='#2e75b6', fg='white', **kw):
        b = tk.Button(parent, text=text, command=cmd,
                      bg=color, fg=fg, font=('맑은 고딕',9,'bold'),
                      relief='flat', cursor='hand2', **kw)
        def lighten(c):
            r,g,b2=int(c[1:3],16),int(c[3:5],16),int(c[5:7],16)
            return '#{:02x}{:02x}{:02x}'.format(min(r+25,255),min(g+25,255),min(b2+25,255))
        b.bind('<Enter>', lambda e: b.config(bg=lighten(color)))
        b.bind('<Leave>', lambda e: b.config(bg=color))
        return b

    def _section(self, parent, title):
        outer = tk.Frame(parent, bg='#131f2e', bd=1, relief='solid')
        outer.pack(fill='x', pady=(0,8), padx=2)
        hdr = tk.Frame(outer, bg='#0d1e33')
        hdr.pack(fill='x')
        tk.Label(hdr, text=title, bg='#0d1e33', fg='#7dd3fc',
                 font=('맑은 고딕',9,'bold'), pady=5, padx=10).pack(anchor='w')
        body = tk.Frame(outer, bg='#131f2e')
        body.pack(fill='x', padx=10, pady=7)
        return body

    def _lbl(self, parent, text, **kw):
        return tk.Label(parent, text=text, bg='#131f2e', fg='#7dd3fc',
                        font=('맑은 고딕',8,'bold'), **kw)

    def _entry(self, parent, var, **kw):
        return tk.Entry(parent, textvariable=var, bg='#1e3a5f', fg='#e0e8f0',
                        insertbackground='white', font=('맑은 고딕',10),
                        relief='flat', **kw)

    # ── UI 빌드 ──
    def _build_ui(self):
        hdr = tk.Frame(self.root, bg='#0d1e33', pady=10)
        hdr.pack(fill='x')
        tk.Label(hdr, text='🏭  배출방지시설 운영기록부 자동 생성기',
                 bg='#0d1e33', fg='#7dd3fc', font=('맑은 고딕',13,'bold'), padx=18).pack(side='left')
        # template.docx 상태 표시
        tpl_ok = self.tpl_bytes is not None
        tpl_txt = '✅ template.docx 로드됨' if tpl_ok else '⚠️ template.docx 없음'
        tpl_fg  = '#34d399' if tpl_ok else '#f87171'
        tk.Label(hdr, text=tpl_txt, bg='#0d1e33', fg=tpl_fg,
                 font=('맑은 고딕',8,'bold'), padx=10).pack(side='right', padx=8)

        main = tk.Frame(self.root, bg='#0f1923')
        main.pack(fill='both', expand=True, padx=10, pady=8)

        sidebar = tk.Frame(main, bg='#0f1923', width=268)
        sidebar.pack(side='left', fill='y')
        sidebar.pack_propagate(False)

        right = tk.Frame(main, bg='#0f1923')
        right.pack(side='left', fill='both', expand=True, padx=(10,0))

        self._build_sidebar(sidebar)
        self._build_table(right)

    def _build_sidebar(self, parent):
        # ── 스크롤 가능한 사이드바 ──
        canvas = tk.Canvas(parent, bg='#0f1923', highlightthickness=0)
        vsb    = ttk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        inner = tk.Frame(canvas, bg='#0f1923')
        win   = canvas.create_window((0,0), window=inner, anchor='nw')

        def on_resize(e):
            canvas.itemconfig(win, width=e.width)
        def on_frame(e):
            canvas.configure(scrollregion=canvas.bbox('all'))
        canvas.bind('<Configure>', on_resize)
        inner.bind('<Configure>', on_frame)
        canvas.bind_all('<MouseWheel>',
            lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), 'units'))

        p = inner  # 이하 parent 역할

        # ① 기간
        s = self._section(p, '① 기간 설정')
        self._lbl(s,'시작일').pack(anchor='w')
        self._entry(s, self.start_var).pack(fill='x', pady=(2,6))
        self._lbl(s,'종료일').pack(anchor='w')
        self._entry(s, self.end_var).pack(fill='x', pady=(2,6))
        self._btn(s,'📊  테이블 생성', self.build_table,
                  color='#1e3a5f', fg='#94a3b8', pady=6).pack(fill='x')

        # ─── 기상청 API ───
        s = self._section(p, '🌡️  기상청 API 날씨 자동 입력')

        info = tk.Frame(s, bg='#0d1e33', padx=8, pady=7)
        info.pack(fill='x', pady=(0,8))
        guide = ("① data.go.kr 회원가입\n"
                 "② '기상청_지상(ASOS) 일별자료' 검색\n"
                 "③ 활용신청 → API 키 복사\n"
                 "④ 아래에 입력 후 자동 입력 클릭")
        tk.Label(info, text=guide, bg='#0d1e33', fg='#94a3b8',
                 font=('맑은 고딕',8), justify='left').pack(anchor='w')
        link_btn = tk.Label(info, text='🔗 data.go.kr 바로가기',
                            bg='#0d1e33', fg='#60a5fa',
                            font=('맑은 고딕',8,'underline'), cursor='hand2')
        link_btn.pack(anchor='w', pady=(4,0))
        link_btn.bind('<Button-1>', lambda e: __import__('webbrowser').open(
            'https://www.data.go.kr/data/15012690/openapi.do'))

        self._lbl(s,'API 키').pack(anchor='w')
        self._entry(s, self.api_key, show='').pack(fill='x', pady=(2,6))

        self._lbl(s,'관측소').pack(anchor='w')
        cb = ttk.Combobox(s, textvariable=self.station,
                          values=list(STATIONS.keys()), state='readonly',
                          font=('맑은 고딕',9))
        cb.pack(fill='x', pady=(2,8))

        self.api_btn = self._btn(s, '🌤️  날씨·기온 자동 입력',
                                  self.fetch_weather_api,
                                  color='#1e40af', pady=8)
        self.api_btn.pack(fill='x')
        self.api_lbl = tk.Label(s, text='API 키를 입력하고 테이블을 먼저 생성하세요',
                                bg='#131f2e', fg='#64748b',
                                font=('맑은 고딕',8), wraplength=230, justify='left')
        self.api_lbl.pack(anchor='w', pady=(5,0))

        # ⚡ 전력 추정값 자동 생성
        s = self._section(p, '⚡  전력계 추정값 자동 생성')
        info2 = tk.Frame(s, bg='#0d1e33', padx=8, pady=6)
        info2.pack(fill='x', pady=(0,8))
        tk.Label(info2,
                 text='VAL별 시작값·하루증가량을 설정하면\n매일 비슷하게 올라가는 값을 자동 생성합니다.',
                 bg='#0d1e33', fg='#94a3b8', font=('맑은 고딕',8), justify='left').pack(anchor='w')
        self._btn(s, '⚙️  추정값 설정 및 생성', self.open_estimate_dialog,
                  color='#7c3aed', pady=8).pack(fill='x')
        self.est_lbl = tk.Label(s, text='설정창에서 VAL별 수치를 입력하세요',
                                bg='#131f2e', fg='#64748b',
                                font=('맑은 고딕',8), wraplength=230, justify='left')
        self.est_lbl.pack(anchor='w', pady=(5,0))

        # ② CSV
        s = self._section(p, '② data.csv 불러오기 (선택)')
        self._btn(s,'📂  CSV 파일 선택', self.load_csv,
                  color='#1e3a5f', fg='#94a3b8', pady=6).pack(fill='x')
        self.csv_lbl = tk.Label(s, text='CSV 불러오기 전', bg='#131f2e', fg='#64748b',
                                font=('맑은 고딕',8))
        self.csv_lbl.pack(anchor='w', pady=(4,0))

        # 날씨·기온 수동 일괄
        s = self._section(p, '🌤️  날씨 · 기온 수동 일괄 설정')
        wf = tk.Frame(s, bg='#131f2e')
        wf.pack(fill='x')
        for lbl,val in [('🎲 랜덤','random'),('☀️ 맑음','맑음'),('🌥️ 흐림','흐림'),
                         ('⛅ 구름조금','구름조금'),('🌧️ 비','비'),('❄️ 눈','눈')]:
            tk.Button(wf, text=lbl, command=lambda v=val: self.fill_weather(v),
                      bg='#1e3a5f', fg='#94a3b8', font=('맑은 고딕',8),
                      relief='flat', cursor='hand2', pady=3, padx=3).pack(
                      side='left', padx=1, pady=1)
        tf = tk.Frame(s, bg='#131f2e')
        tf.pack(fill='x', pady=(6,0))
        tk.Label(tf, text='기온(℃):', bg='#131f2e', fg='#7dd3fc',
                 font=('맑은 고딕',8)).pack(side='left')
        self._entry(tf, self.temp_var, width=7).pack(side='left', padx=4)
        self._btn(tf,'적용', self.fill_temp,
                  color='#1e3a5f', fg='#94a3b8', padx=8).pack(side='left')

        # 옵션
        s = self._section(p, '⚙️  생성 옵션')
        for text,var in [('전력 데이터 없는 날짜 건너뜀', self.skip_empty),
                          ('주말 건너뜀', self.skip_wknd)]:
            tk.Checkbutton(s, text=text, variable=var, command=self.update_status,
                           bg='#131f2e', fg='#94a3b8', selectcolor='#1e3a5f',
                           activebackground='#131f2e', font=('맑은 고딕',9),
                           relief='flat').pack(anchor='w')

        # ③ 출력 폴더
        s = self._section(p, '③ 출력 폴더 선택')
        self._btn(s,'📁  폴더 선택', self.select_output,
                  color='#1e3a5f', fg='#94a3b8', pady=6).pack(fill='x')
        self.out_lbl = tk.Label(s, textvariable=self.out_dir, bg='#131f2e', fg='#64748b',
                                font=('맑은 고딕',8), wraplength=220, justify='left')
        self.out_lbl.pack(anchor='w', pady=(4,0))

        # ④ 생성
        s = self._section(p, '④ 파일 생성')
        self.status_lbl = tk.Label(s, text='⚠️ 설정을 완료하세요', bg='#131f2e',
                                    fg='#fbbf24', font=('맑은 고딕',8),
                                    wraplength=220, justify='left')
        self.status_lbl.pack(anchor='w', pady=(0,6))
        self.gen_btn = self._btn(s,'🚀  워드 파일 생성', self.generate_all,
                                  color='#059669', pady=11)
        self.gen_btn.pack(fill='x')
        self.gen_btn.config(state='disabled', bg='#064e3b')
        self.prog_lbl = tk.Label(s, text='', bg='#131f2e', fg='#64748b',
                                  font=('맑은 고딕',8), wraplength=220)
        self.prog_lbl.pack(anchor='w', pady=(5,0))

    def _build_table(self, parent):
        hdr = tk.Frame(parent, bg='#0d1e33')
        hdr.pack(fill='x')
        tk.Label(hdr, text='📋  전력계 데이터 입력표',
                 bg='#0d1e33', fg='#7dd3fc', font=('맑은 고딕',10,'bold'),
                 pady=7, padx=10).pack(side='left')
        tk.Label(hdr,
                 text='💡 날짜클릭:쉬는날  드래그:다중선택  Ctrl+C:복사  Ctrl+V:붙여넣기  더블클릭:수정',
                 bg='#0d1e33', fg='#64748b', font=('맑은 고딕',8), padx=10).pack(side='right')

        frame = tk.Frame(parent, bg='#131f2e', bd=1, relief='solid')
        frame.pack(fill='both', expand=True, pady=(4,0))

        vsb = ttk.Scrollbar(frame, orient='vertical')
        hsb = ttk.Scrollbar(frame, orient='horizontal')
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')

        self.tree = ttk.Treeview(frame, columns=COLS, show='headings',
                                  yscrollcommand=vsb.set, xscrollcommand=hsb.set,
                                  selectmode='extended')
        for col,w in zip(COLS, COL_WIDTHS):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor='center', minwidth=40)
        self.tree.tag_configure('wknd',    foreground='#f87171')
        self.tree.tag_configure('normal',  foreground='#e0e8f0')
        self.tree.tag_configure('api_ok',  foreground='#34d399')
        self.tree.tag_configure('holiday', foreground='#fb923c', background='#1c1006')
        self.tree.pack(fill='both', expand=True)
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        self.tree.bind('<Double-1>',   self.on_dbl_click)
        self.tree.bind('<Button-1>',   self.on_click)
        self.tree.bind('<Control-c>',  self.copy_selection)
        self.tree.bind('<Control-C>',  self.copy_selection)
        self.tree.bind('<Control-v>',  self.paste_to_table)
        self.tree.bind('<Control-V>',  self.paste_to_table)

        self.cnt_lbl = tk.Label(parent,
            text='← 기간을 설정하고 테이블 생성을 누르세요',
            bg='#0a1520', fg='#64748b', font=('맑은 고딕',9), pady=5, anchor='w')
        self.cnt_lbl.pack(fill='x', side='bottom')

    # ── 동작 ──
    def select_output(self):
        p = filedialog.askdirectory(title='출력 폴더 선택')
        if p:
            self.out_dir.set(p)
            self.out_lbl.config(fg='#34d399')
            save_config({'out_dir': p})
            self.update_status()

    def build_table(self):
        try:
            sd = date.fromisoformat(self.start_var.get())
            ed = date.fromisoformat(self.end_var.get())
        except ValueError:
            messagebox.showerror('오류','날짜 형식: YYYY-MM-DD'); return
        if sd > ed:
            messagebox.showerror('오류','시작일이 종료일보다 늦습니다'); return
        # 기존 쉬는날 보존
        prev_holidays = {r['date_str'] for r in self.rows if r.get('is_holiday')}
        self.rows = []
        cur = sd
        while cur <= ed:
            self.rows.append({
                'date':       cur,
                'date_str':   cur.isoformat(),
                'dow':        get_dow(cur),
                'is_wknd':    is_wknd(cur),
                'is_holiday': cur.isoformat() in prev_holidays,
                'power':      ['']*8,
                'weather':    rand_weather(),
                'temp':       str(rand_temp(cur.month)),
            })
            cur += timedelta(days=1)
        self.render_table()
        self.update_status()
        self.update_count()

    # ── 전력 추정값 다이얼로그 ──
    def open_estimate_dialog(self):
        if not self.rows:
            messagebox.showwarning('알림', '먼저 테이블을 생성하세요'); return

        dlg = tk.Toplevel(self.root)
        dlg.title('⚡ 전력계 추정값 설정')
        dlg.configure(bg='#0f1923')
        dlg.resizable(False, False)
        dlg.grab_set()

        cfg     = load_config()
        est_cfg = cfg.get('estimate', {})

        VAL_NAMES = [
            ('VAL1', '흡수시설 (3층옥상)'),
            ('VAL2', '흡착시설 (3층옥상)'),
            ('VAL3', '흡착시설 (3층옥상)'),
            ('VAL4', '여과시설 (3층옥상)'),
            ('VAL5', '여과시설 (3층옥상)'),
            ('VAL6', '여과시설 (2층외부)'),
            ('VAL7', '여과시설 (3층옥상)'),
            ('VAL8', '흡수시설 (별관3층)'),
        ]

        # ── 헤더 ──
        hdr = tk.Frame(dlg, bg='#0d1e33', pady=10)
        hdr.pack(fill='x')
        tk.Label(hdr, text='⚡  전력계 추정값 설정 (실측값 기반)',
                 bg='#0d1e33', fg='#7dd3fc', font=('맑은 고딕',11,'bold'), padx=16).pack(side='left')

        # ── 설명 ──
        desc = tk.Frame(dlg, bg='#131f2e', padx=14, pady=9)
        desc.pack(fill='x', padx=12, pady=(10,0))
        tk.Label(desc,
                 text=('시작값  :  오늘 이전 마지막 실측값\n'
                       '끝  값  :  오늘 실측값\n'
                       '하루증가량 = (끝값 − 시작값) ÷ 기간 내 평일 수  (자동 계산)'),
                 bg='#131f2e', fg='#94a3b8', font=('맑은 고딕',8), justify='left').pack(anchor='w')
        tk.Label(desc,
                 text='💡 Excel 열(시작값 8행 또는 끝값 8행) 복사 후 첫 번째 셀에 붙여넣기(Ctrl+V)',
                 bg='#131f2e', fg='#60a5fa', font=('맑은 고딕',8,'bold')).pack(anchor='w', pady=(5,0))

        # ── 주말 옵션 (평일 수 계산에 영향) ──
        opt_frame = tk.Frame(dlg, bg='#0d1e33', padx=14, pady=6)
        opt_frame.pack(fill='x', padx=12, pady=(8,0))
        wknd_inc  = tk.BooleanVar(value=est_cfg.get('wknd_inc', False))
        wd_lbl    = tk.Label(opt_frame, bg='#0d1e33', fg='#fbbf24',
                             font=('맑은 고딕',8,'bold'))
        wd_lbl.pack(side='right', padx=6)

        def calc_wd():
            return sum(1 for r in self.rows
                       if not r.get('is_holiday')
                       and not (r['is_wknd'] and not wknd_inc.get()))

        def refresh_wd(*_):
            wd_lbl.config(text=f'기간 내 평일 수: {calc_wd()}일')
            refresh_inc()

        tk.Checkbutton(opt_frame,
                       text='주말에도 증가량 적용',
                       variable=wknd_inc, command=refresh_wd,
                       bg='#0d1e33', fg='#94a3b8', selectcolor='#1e3a5f',
                       activebackground='#0d1e33', font=('맑은 고딕',9), relief='flat').pack(side='left')

        # ── 컬럼 헤더 ──
        COL_DEFS = [
            ('',              4,  '#0d1e33'),   # 번호
            ('설비',          10, '#0d1e33'),
            ('시작값\n(이전실측)', 12, '#0d1e33'),
            ('끝값\n(오늘실측)',   12, '#0d1e33'),
            ('변동폭(%)',      8,  '#0d1e33'),
            ('하루증가량\n(자동)',  12, '#0d1e33'),
        ]
        hdr_f = tk.Frame(dlg, bg='#0d1e33')
        hdr_f.pack(fill='x', padx=12, pady=(8,0))
        for txt, w, bg in COL_DEFS:
            tk.Label(hdr_f, text=txt, bg=bg, fg='#7dd3fc',
                     font=('맑은 고딕',8,'bold'), width=w, anchor='center',
                     pady=4).pack(side='left', padx=1)

        # ── 입력 그리드 ──
        grid_f = tk.Frame(dlg, bg='#0a1520')
        grid_f.pack(fill='x', padx=12, pady=(1,0))

        sv_starts = []   # StringVar 목록 (시작값)
        sv_ends   = []   # StringVar 목록 (끝값)
        sv_vars   = []   # StringVar 목록 (변동폭)
        inc_vars  = []   # StringVar 목록 (하루증가량 미리보기, read-only)
        ent_starts = []  # Entry 위젯 목록
        ent_ends   = []  # Entry 위젯 목록

        def refresh_inc(*_):
            """하루증가량 자동 계산 및 미리보기 갱신"""
            wd = max(calc_wd(), 1)
            for i in range(8):
                try:
                    s = float(sv_starts[i].get().replace(',', ''))
                    e = float(sv_ends[i].get().replace(',', ''))
                    inc = (e - s) / wd
                    # 소수점 처리
                    inc_vars[i].set(f'{inc:.2f}' if inc != int(inc) else str(int(inc)))
                except (ValueError, IndexError):
                    if i < len(inc_vars):
                        inc_vars[i].set('—')

        def make_paste_handler(col_sv_list, row_i, other_col_sv=None):
            """Excel 여러 행 붙여넣기: 탭·줄바꿈으로 구분된 값을 아래 행에 채움"""
            def handler(event):
                try:
                    cb = event.widget.clipboard_get()
                    # 탭 구분(행 내 셀) 또는 줄바꿈 구분(행 간 셀) 모두 지원
                    # Excel에서 단일 열 복사 → 줄바꿈 / 여러 열 복사 → 탭+줄바꿈
                    lines = [ln for ln in cb.splitlines() if ln.strip()]
                    if len(lines) <= 1:
                        return   # 단일값은 기본 붙여넣기 사용
                    for j, line in enumerate(lines):
                        # 탭이 있으면 첫 번째 열만 사용
                        val = line.split('\t')[0].strip()
                        idx = row_i + j
                        if idx < len(col_sv_list):
                            col_sv_list[idx].set(val)
                    refresh_inc()
                    return 'break'
                except Exception:
                    pass
            return handler

        for i, (vcode, vname) in enumerate(VAL_NAMES):
            key   = f'val{i}'
            saved = est_cfg.get(key, {})
            sv_s  = tk.StringVar(value=str(saved.get('start', '')))
            sv_e  = tk.StringVar(value=str(saved.get('end',   '')))
            sv_v  = tk.StringVar(value=str(saved.get('var',   '10')))
            sv_i  = tk.StringVar(value='—')
            sv_starts.append(sv_s); sv_ends.append(sv_e)
            sv_vars.append(sv_v);   inc_vars.append(sv_i)

            bg = '#131f2e' if i % 2 == 0 else '#0d1e33'
            row_f = tk.Frame(grid_f, bg=bg)
            row_f.pack(fill='x', pady=0)

            # 번호
            tk.Label(row_f, text=str(i+1), bg=bg, fg='#64748b',
                     font=('맑은 고딕',8), width=4, anchor='center').pack(side='left', padx=1)
            # 설비명
            tk.Label(row_f, text=f'{vcode}\n{vname}', bg=bg, fg='#e0e8f0',
                     font=('맑은 고딕',7), width=10, anchor='w', justify='left').pack(side='left', padx=1)

            CELL_STYLE = dict(bg='#1e3a5f', fg='#e0e8f0', insertbackground='white',
                              font=('맑은 고딕',9), relief='flat', justify='center',
                              highlightthickness=1, highlightbackground='#2e4a6f',
                              highlightcolor='#60a5fa')

            # 시작값
            es = tk.Entry(row_f, textvariable=sv_s, width=12, **CELL_STYLE)
            es.pack(side='left', padx=1, pady=2)
            sv_s.trace_add('write', refresh_inc)
            ent_starts.append(es)

            # 끝값
            ee = tk.Entry(row_f, textvariable=sv_e, width=12, **CELL_STYLE)
            ee.pack(side='left', padx=1, pady=2)
            sv_e.trace_add('write', refresh_inc)
            ent_ends.append(ee)

            # 변동폭
            ev = tk.Entry(row_f, textvariable=sv_v, width=8, **CELL_STYLE)
            ev.pack(side='left', padx=1, pady=2)

            # 하루증가량 (read-only 미리보기)
            tk.Label(row_f, textvariable=sv_i, bg=bg, fg='#34d399',
                     font=('맑은 고딕',9,'bold'), width=12, anchor='center').pack(side='left', padx=4)

        # 붙여넣기 핸들러 등록 (열별로 독립 적용)
        for i in range(8):
            ent_starts[i].bind('<Control-v>', make_paste_handler(sv_starts, i))
            ent_ends[i].bind('<Control-v>',   make_paste_handler(sv_ends,   i))

        # 초기 미리보기 및 평일 수 표시
        refresh_wd()

        # ── 버튼 ──
        btn_frame = tk.Frame(dlg, bg='#0f1923', pady=12)
        btn_frame.pack(fill='x', padx=12)

        def on_apply():
            wd = calc_wd()
            if wd == 0:
                messagebox.showerror('오류', '기간 내 유효한 평일이 없습니다', parent=dlg)
                return

            settings = []
            for i in range(8):
                key = f'val{i}'
                try:
                    s_val = float(sv_starts[i].get().replace(',','')) if sv_starts[i].get().strip() else None
                    e_val = float(sv_ends[i].get().replace(',',''))   if sv_ends[i].get().strip()   else None
                    v_val = float(sv_vars[i].get())   if sv_vars[i].get().strip()   else 10.0
                except ValueError:
                    messagebox.showerror('오류', f'VAL{i+1}: 숫자만 입력하세요', parent=dlg)
                    return

                if s_val is None and e_val is None:
                    settings.append({'start': None, 'inc': None, 'var': v_val})
                elif s_val is not None and e_val is not None:
                    inc = (e_val - s_val) / wd
                    settings.append({'start': s_val, 'inc': inc, 'var': v_val})
                else:
                    messagebox.showerror('오류',
                        f'VAL{i+1}: 시작값·끝값을 둘 다 입력하거나 둘 다 비워두세요', parent=dlg)
                    return

            # config 저장
            est_save = {f'val{i}': {'start': float(sv_starts[i].get().replace(',','') or 0)
                                             if sv_starts[i].get().strip() else None,
                                    'end':   float(sv_ends[i].get().replace(',','') or 0)
                                             if sv_ends[i].get().strip() else None,
                                    'var':   float(sv_vars[i].get() or 10)}
                        for i in range(8)}
            est_save['wknd_inc'] = wknd_inc.get()
            save_config({'estimate': est_save})

            self._apply_estimate(settings, wknd_inc.get())
            dlg.destroy()

        self._btn(btn_frame, '✅  적용 (테이블에 자동 입력)', on_apply,
                  color='#7c3aed', pady=9).pack(side='left', fill='x', expand=True, padx=(0,6))
        self._btn(btn_frame, '취소', dlg.destroy,
                  color='#1e3a5f', fg='#94a3b8', pady=9).pack(side='left', ipadx=10)

        dlg.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width()  - dlg.winfo_width())  // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dlg.winfo_height()) // 2
        dlg.geometry(f'+{x}+{y}')

    def _apply_estimate(self, settings, wknd_inc: bool):
        cur_vals          = [s['start'] for s in settings]
        prev_weekday_vals = [s['start'] for s in settings]

        filled = 0
        for row in self.rows:
            new_power = []
            for i, s in enumerate(settings):
                start = s['start']
                inc   = s['inc']
                var   = s['var'] / 100.0

                if start is None and inc is None:
                    new_power.append(row['power'][i])
                    continue

                if cur_vals[i] is None:
                    cur_vals[i] = start or 0.0

                is_rest = row['is_wknd'] or row.get('is_holiday', False)
                if is_rest and not wknd_inc:
                    val = prev_weekday_vals[i] if prev_weekday_vals[i] is not None else cur_vals[i]
                else:
                    if inc is not None:
                        delta = inc * (1 + random.uniform(-var, var))
                        cur_vals[i] = (cur_vals[i] or 0) + delta
                    val = cur_vals[i]
                    if not is_rest:
                        prev_weekday_vals[i] = val

                new_power.append(f'{val:.2f}')

            row['power'] = new_power
            if any(p for p in new_power): filled += 1

        self.render_table()
        self.update_count()
        self.est_lbl.config(text=f'✅ {filled}일 추정값 생성 완료', fg='#c4b5fd')

    # ── 복사 / 붙여넣기 ──
    def copy_selection(self, event=None):
        """선택된 행을 TSV로 클립보드에 복사 (Excel 붙여넣기 호환)"""
        sel = self.tree.selection()
        if not sel: return
        lines = []
        for item in sel:
            vals = list(self.tree.item(item, 'values'))
            vals[0] = vals[0].replace('🚫 ', '').strip()   # 쉬는날 프리픽스 제거
            lines.append('\t'.join(str(v) for v in vals))
        self.root.clipboard_clear()
        self.root.clipboard_append('\n'.join(lines))
        self._flash_status(f'✅ {len(sel)}행 복사됨 — Excel에 Ctrl+V로 붙여넣기 가능')
        return 'break'

    def paste_to_table(self, event=None):
        """클립보드 TSV를 테이블에 붙여넣기
        - 우리 앱 형식(12열): 날짜·요일 제외하고 VAL1-8 자동 인식
        - Excel 원시 형식(N열): 순서대로 VAL1-8에 채움
        """
        try:
            cb = self.root.clipboard_get()
        except Exception:
            return
        lines = [l for l in cb.splitlines() if l.strip()]
        if not lines: return

        sel = self.tree.selection()
        start_idx = self.tree.index(sel[0]) if sel else 0

        changed = 0
        for j, line in enumerate(lines):
            row_idx = start_idx + j
            if row_idx >= len(self.rows): break
            cols = line.split('\t')
            row  = self.rows[row_idx]

            # 형식 자동 감지: 첫 열이 날짜(YYYY-MM-DD)이면 앱 복사본
            first = cols[0].strip().replace('🚫 ', '')
            if re.match(r'^\d{4}-\d{2}-\d{2}$', first):
                # 앱 형식: 날짜(0) 요일(1) VAL1-8(2~9) 날씨(10) 기온(11)
                power_vals = [c.strip() for c in cols[2:10]]
            else:
                # Excel 원시 형식: 전체 열을 VAL1-8로 사용
                power_vals = [c.strip().replace(',', '') for c in cols[:8]]

            for k, v in enumerate(power_vals):
                if k < 8 and v:
                    row['power'][k] = v
            if any(power_vals):
                changed += 1

        if changed:
            self.render_table()
            self.update_count()
            self._flash_status(f'✅ {changed}행 붙여넣기 완료')
        return 'break'

    def _flash_status(self, msg, ms=2500):
        """상태바에 잠깐 메시지 표시 후 원복"""
        self.cnt_lbl.config(text=msg, fg='#34d399')
        self.root.after(ms, self.update_count)

    def on_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != 'cell': return
        col    = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if not row_id or col != '#1': return
        row_idx = self.tree.index(row_id)
        row = self.rows[row_idx]
        row['is_holiday'] = not row['is_holiday']
        self._refresh_row(row_id, row)
        self.update_count()

    def _refresh_row(self, row_id, row):
        date_disp = ('🚫 ' + row['date_str']) if row['is_holiday'] else row['date_str']
        vals = (date_disp, row['dow'], *row['power'], row['weather'], row['temp'])
        self.tree.item(row_id, values=vals)
        if row['is_holiday']:
            tag = 'holiday'
        elif row['is_wknd']:
            tag = 'wknd'
        else:
            tag = 'normal'
        self.tree.item(row_id, tags=(tag,))

    def render_table(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        for row in self.rows:
            date_disp = ('🚫 ' + row['date_str']) if row['is_holiday'] else row['date_str']
            vals = (date_disp, row['dow'], *row['power'],
                    row['weather'], row['temp'])
            if row['is_holiday']:
                tag = 'holiday'
            elif row['is_wknd']:
                tag = 'wknd'
            else:
                tag = 'normal'
            self.tree.insert('', 'end', values=vals, tags=(tag,))
        self.update_count()

    # ── 기상청 API 호출 ──
    def fetch_weather_api(self):
        key = self.api_key.get().strip()
        if not key:
            messagebox.showwarning('알림','API 키를 입력하세요'); return
        if not self.rows:
            messagebox.showwarning('알림','먼저 테이블을 생성하세요'); return

        self.api_btn.config(state='disabled', text='⏳ 가져오는 중...')
        self.api_lbl.config(text='기상청 서버에 요청 중...', fg='#fbbf24')
        self.root.update()

        stn_id    = STATIONS.get(self.station.get(), '108')
        start_str = self.rows[0]['date_str']
        end_str   = self.rows[-1]['date_str']

        def run():
            try:
                result = fetch_asos(key, stn_id, start_str, end_str)
                self.root.after(0, lambda: self._apply_api_result(result))
            except URLError as e:
                self.root.after(0, lambda: self._api_error(f'네트워크 오류: {e.reason}'))
            except Exception as e:
                self.root.after(0, lambda: self._api_error(str(e)))

        threading.Thread(target=run, daemon=True).start()

    def _apply_api_result(self, result):
        filled = 0
        for row in self.rows:
            d = result.get(row['date_str'])
            if d:
                row['weather'] = d['weather']
                row['temp']    = d['temp']
                filled += 1
        self.render_table()
        for i, item in enumerate(self.tree.get_children()):
            if result.get(self.rows[i]['date_str']):
                r = self.rows[i]
                if r.get('is_holiday'):
                    cur_tag = 'holiday'
                elif r['is_wknd']:
                    cur_tag = 'wknd'
                else:
                    cur_tag = 'normal'
                self.tree.item(item, tags=(cur_tag, 'api_ok'))
        self.api_lbl.config(
            text=f'✅ {filled}일 날씨·기온 자동 입력 완료 (전체 {len(self.rows)}일 중)',
            fg='#34d399')
        self.api_btn.config(state='normal', text='🌤️  날씨·기온 자동 입력')

    def _api_error(self, msg):
        self.api_lbl.config(text=f'❌ {msg}', fg='#f87171')
        self.api_btn.config(state='normal', text='🌤️  날씨·기온 자동 입력')
        messagebox.showerror('API 오류', msg)

    def on_dbl_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != 'cell': return
        col    = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if not row_id: return
        col_idx = int(col[1:]) - 1
        row_idx = self.tree.index(row_id)
        if col_idx < 1: return
        try:
            x, y, w, h = self.tree.bbox(row_id, COLS[col_idx])
        except Exception:
            return
        cur_val = self.tree.set(row_id, COLS[col_idx])

        if col_idx == 10:
            cb = ttk.Combobox(self.tree, values=WEATHER_OPT, state='readonly',
                              font=('맑은 고딕',9))
            cb.set(cur_val)
            cb.place(x=x, y=y, width=w, height=h)
            def save_cb(e=None):
                v = cb.get()
                self.rows[row_idx]['weather'] = v
                self.tree.set(row_id, COLS[col_idx], v)
                cb.destroy()
            cb.bind('<<ComboboxSelected>>', save_cb)
            cb.bind('<FocusOut>', save_cb)
            cb.focus_set()
        else:
            ent = tk.Entry(self.tree, bg='#1e3a5f', fg='#e0e8f0',
                           insertbackground='white', font=('맑은 고딕',9),
                           justify='center', relief='flat')
            ent.insert(0, cur_val)
            ent.select_range(0, 'end')
            ent.place(x=x, y=y, width=w, height=h)
            def save_ent(e=None):
                v = ent.get()
                if 2 <= col_idx <= 9:
                    self.rows[row_idx]['power'][col_idx-2] = v
                elif col_idx == 11:
                    self.rows[row_idx]['temp'] = v
                self.tree.set(row_id, COLS[col_idx], v)
                ent.destroy()
                self.update_count()
            ent.bind('<Return>', save_ent)
            ent.bind('<Tab>',    save_ent)
            ent.bind('<FocusOut>', save_ent)
            ent.focus_set()

    def load_csv(self):
        if not self.rows:
            messagebox.showwarning('경고','먼저 테이블을 생성하세요'); return
        p = filedialog.askopenfilename(title='data.csv 선택',
            filetypes=[('CSV 파일','*.csv'),('모든 파일','*.*')])
        if not p: return
        try:
            content = None
            for enc in ['euc-kr','utf-8-sig','utf-8','cp949']:
                try:
                    with open(p,'r',encoding=enc) as f: content=f.read(); break
                except (UnicodeDecodeError, LookupError): pass
            if content is None: raise ValueError('파일 인코딩을 읽을 수 없습니다')
            lines = [l.strip() for l in content.splitlines() if l.strip()]
            if len(lines)<2: raise ValueError('데이터가 없습니다')
            header = [h.strip() for h in lines[0].split(',')]
            di = next((i for i,h in enumerate(header) if '날짜' in h), 0)
            pi = [next((i for i,h in enumerate(header) if f'전력{n}' in h), -1)
                  for n in range(1,9)]
            dm = {}
            for line in lines[1:]:
                c  = line.split(',')
                ds = (c[di] if di<len(c) else '').strip()[:10]
                if not re.match(r'^\d{4}-\d{2}-\d{2}$', ds): continue
                dm[ds] = [c[i].strip() if 0<=i<len(c) else '' for i in pi]
            filled = 0
            for row in self.rows:
                if row['date_str'] in dm:
                    row['power'] = dm[row['date_str']]
                    filled += 1
            self.render_table()
            self.csv_lbl.config(text=f'✅ {filled}일 데이터 입력됨', fg='#34d399')
            self.update_status()
        except Exception as e:
            messagebox.showerror('CSV 오류', str(e))
            self.csv_lbl.config(text=f'❌ 오류: {e}', fg='#f87171')

    def fill_weather(self, mode):
        if not self.rows: return
        for row in self.rows:
            row['weather'] = rand_weather() if mode=='random' else mode
        self.render_table()

    def fill_temp(self):
        v = self.temp_var.get().strip()
        if not v or not self.rows: return
        for row in self.rows: row['temp'] = v
        self.render_table()

    def update_count(self):
        if not self.rows:
            self.cnt_lbl.config(text='← 기간을 설정하고 테이블 생성을 누르세요'); return
        se = self.skip_empty.get(); sw = self.skip_wknd.get()
        total   = len(self.rows)
        wknd_n  = sum(1 for r in self.rows if r['is_wknd'])
        hday_n  = sum(1 for r in self.rows if r.get('is_holiday'))
        has_d   = sum(1 for r in self.rows if any(p.strip() for p in r['power']))
        gen_n   = sum(1 for r in self.rows
                      if not (sw and r['is_wknd'])
                      and not r.get('is_holiday')
                      and not (se and not any(p.strip() for p in r['power'])))
        hday_txt = f'  |  🚫 쉬는날 {hday_n}일' if hday_n else ''
        self.cnt_lbl.config(
            text=f'전체 {total}일  |  데이터 입력 {has_d}일  |  주말 {wknd_n}일{hday_txt}  |  생성 예정 {gen_n}일')
        self.update_status()

    def update_status(self):
        def disable(msg):
            self.status_lbl.config(text=msg, fg='#fbbf24')
            self.gen_btn.config(state='disabled', bg='#064e3b')
        if not self.tpl_bytes:
            disable('⚠️ template.docx 를 이 스크립트와 같은 폴더에 넣으세요'); return
        if not self.rows:
            disable('⚠️ 기간을 설정하고 테이블을 생성하세요'); return
        if self.out_dir.get() == '선택 안 됨':
            disable('⚠️ 출력 폴더를 선택하세요'); return
        se = self.skip_empty.get(); sw = self.skip_wknd.get()
        gen = [r for r in self.rows
               if not r.get('is_holiday')
               and not (sw and r['is_wknd'])
               and not (se and not any(p.strip() for p in r['power']))]
        if not gen:
            disable('⚠️ 생성할 파일이 없습니다 (옵션 확인)'); return
        self.status_lbl.config(text=f'✅ 준비 완료 — {len(gen)}개 파일 생성 예정', fg='#34d399')
        self.gen_btn.config(state='normal', bg='#059669')

    def generate_all(self):
        se = self.skip_empty.get(); sw = self.skip_wknd.get()
        gen = [r for r in self.rows
               if not r.get('is_holiday')
               and not (sw and r['is_wknd'])
               and not (se and not any(p.strip() for p in r['power']))]
        if not gen: messagebox.showwarning('경고','생성할 파일이 없습니다'); return
        out = self.out_dir.get()
        self.gen_btn.config(state='disabled', text='⏳ 생성 중...')
        self.root.update()
        errors = []
        try:
            for i, row in enumerate(gen):
                month_dir = os.path.join(out, f"{row['date'].month:02d}월")
                os.makedirs(month_dir, exist_ok=True)
                tag   = row['date_str'][2:].replace('-','')
                fname = f"배출방지시설운영기록부({tag}).docx"
                try:
                    data = fill_docx(self.tpl_bytes, row)
                    with open(os.path.join(month_dir, fname),'wb') as f:
                        f.write(data)
                except Exception as e:
                    errors.append(f'{fname}: {e}')
                pct = round((i+1)/len(gen)*100)
                self.prog_lbl.config(text=f'({i+1}/{len(gen)})  {fname}  [{pct}%]')
                self.root.update()
            if errors:
                messagebox.showwarning('완료 (일부 오류)',
                    f'{len(gen)-len(errors)}개 성공\n오류:\n'+'\n'.join(errors[:5]))
            else:
                self.status_lbl.config(text=f'✅ {len(gen)}개 파일 생성 완료!', fg='#34d399')
                self.prog_lbl.config(text=f'완료!  {len(gen)}개 파일 → {out}')
                if messagebox.askyesno('완료',
                        f'✅ {len(gen)}개 파일이 생성되었습니다.\n\n출력 폴더를 열겠습니까?'):
                    os.startfile(out)
        except Exception as e:
            messagebox.showerror('오류', str(e))
        finally:
            self.gen_btn.config(state='normal', text='🚀  워드 파일 생성')


if __name__ == '__main__':
    root = tk.Tk()
    App(root)
    root.mainloop()
