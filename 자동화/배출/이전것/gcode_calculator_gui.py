#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
G-code 파라미터 자동 계산 프로그램 (GUI 버전)
O0852/O9001/O9002 로직 분석 및 검증
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import math


class GCodeCalculator:
    def __init__(self):
        self.params = {}
        
    def set_parameters(self, **kwargs):
        """파라미터 설정"""
        self.params.update(kwargs)
        
    def calculate_safety_override(self):
        """안전 잔량 계산 (#129)"""
        mat_type = self.params.get('mat_type', 1)
        raw_od = self.params.get('raw_od', 70)
        
        if mat_type == 2 or raw_od < 30:
            return 15
        return 0
    
    def calculate_total_length(self):
        """총 소재 길이 계산 (#140)"""
        mat_oo = self.params.get('mat_oo', 0)
        mat_ooo = self.params.get('mat_ooo', 0)
        safety_res = self.calculate_safety_override()
        
        return mat_oo + mat_ooo - safety_res
    
    def calculate_unit_length(self):
        """제품 1개 길이 (#505)"""
        fin_length = self.params.get('fin_length', 0)
        tip_width = self.params.get('tip_width', 0)
        
        return fin_length + tip_width
    
    def calculate_effective_chuck_length(self):
        """척 유효 길이 (#500)"""
        chuck_length = self.params.get('chuck_length', 0)
        safety_length = self.params.get('safety_length', 0)
        
        return chuck_length - safety_length
    
    def calculate_pieces_per_cycle(self):
        """한 사이클당 제품 개수 (#517)"""
        eff_length = self.calculate_effective_chuck_length()
        unit_length = self.calculate_unit_length()
        
        if unit_length == 0:
            return 0
        
        return int(eff_length / unit_length)
    
    def calculate_cycle_length(self):
        """한 사이클 가공 길이 (#518)"""
        pieces = self.calculate_pieces_per_cycle()
        unit_length = self.calculate_unit_length()
        
        return pieces * unit_length
    
    def calculate_autolink_count(self):
        """필요한 자동링크 횟수 (#519)"""
        total_length = self.calculate_total_length()
        chuck_left = self.params.get('chuck_left_limit', 0)
        chuck2jaw = self.params.get('chuck2jaw', 0)
        safety_length = self.params.get('safety_length', 0)
        cycle_length = self.calculate_cycle_length()
        tip_width = self.params.get('tip_width', 0)
        init_face = self.params.get('init_face_cut', 0)
        margin = self.params.get('margin', 0)
        
        numerator = (total_length - chuck_left - chuck2jaw - safety_length - 
                    cycle_length - tip_width - init_face)
        
        denominator = cycle_length + tip_width + margin
        
        if denominator == 0:
            return 0
        
        return int(numerator / denominator)
    
    def calculate_used_length_after_autolink(self, autolink_count):
        """자동링크 N회 후 사용된 총 길이"""
        init_face = self.params.get('init_face_cut', 0)
        cycle_length = self.calculate_cycle_length()
        tip_width = self.params.get('tip_width', 0)
        margin = self.params.get('margin', 0)
        
        first_cycle = cycle_length + tip_width
        autolink_cycles = autolink_count * (cycle_length + tip_width + margin)
        
        return init_face + first_cycle + autolink_cycles
    
    def calculate_remaining_material(self, autolink_count):
        """자동링크 후 남은 유효 소재 (#538, #539)"""
        total_length = self.calculate_total_length()
        used_length = self.calculate_used_length_after_autolink(autolink_count)
        chuck_left = self.params.get('chuck_left_limit', 0)
        chuck2jaw = self.params.get('chuck2jaw', 0)
        safety_length = self.params.get('safety_length', 0)
        tip_width = self.params.get('tip_width', 0)
        margin = self.params.get('margin', 0)
        
        remaining = (total_length - used_length) - (chuck_left + chuck2jaw + safety_length) - (tip_width + margin)
        
        unit_length = self.calculate_unit_length()
        if unit_length == 0:
            pieces = 0
        else:
            pieces = int(remaining / unit_length)
        
        return remaining, pieces


class GCodeGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("G-Code 파라미터 계산기 v1.0")
        self.root.geometry("1200x800")
        
        self.calculator = GCodeCalculator()
        self.entries = {}
        
        self.create_widgets()
        self.load_default_values()
        
    def create_widgets(self):
        """GUI 위젯 생성"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 좌측: 입력 영역 (스크롤 가능)
        input_outer_frame = ttk.LabelFrame(main_frame, text="입력 파라미터", padding="5")
        input_outer_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # 캔버스와 스크롤바 생성
        canvas = tk.Canvas(input_outer_frame, width=400, height=600)
        scrollbar = ttk.Scrollbar(input_outer_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 마우스 휠 스크롤 바인딩
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # 입력 필드 생성 (스크롤 가능한 프레임 안에)
        self.create_input_fields(scrollable_frame)
        
        # 버튼 프레임 (하단 고정)
        button_frame = ttk.Frame(input_outer_frame)
        button_frame.pack(side="bottom", fill="x", pady=10)
        
        # 계산 버튼
        calc_btn = ttk.Button(
            button_frame, 
            text="🔍 계산 실행", 
            command=self.calculate
        )
        calc_btn.pack(fill="x", pady=5)
        
        # 리셋 버튼
        reset_btn = ttk.Button(
            button_frame,
            text="🔄 기본값 복원",
            command=self.load_default_values
        )
        reset_btn.pack(fill="x")
        
        # 우측: 결과 영역
        output_frame = ttk.LabelFrame(main_frame, text="계산 결과", padding="10")
        output_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # 결과 텍스트 영역
        self.result_text = scrolledtext.ScrolledText(
            output_frame, 
            width=70, 
            height=40,
            font=('Consolas', 9),
            wrap=tk.WORD
        )
        self.result_text.pack(fill="both", expand=True)
        
        # 그리드 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=2)
        main_frame.rowconfigure(0, weight=1)
        
    def create_input_fields(self, parent):
        """입력 필드 생성"""
        row = 0
        
        # A. SETTINGS
        ttk.Label(parent, text="A. 설정 (SETTINGS)", font=('Arial', 10, 'bold')).grid(
            row=row, column=0, columnspan=2, sticky=tk.W, pady=(0, 5)
        )
        row += 1
        
        fields_a = [
            ("mat_oo", "#101 소재 길이 OO (0~99)", 60),
            ("mat_ooo", "#102 소재 길이 OOO (100~500)", 500),
            ("chuck_length", "#103 척 길이", 75),
            ("init_face_cut", "#104 초기 면 절삭", 1),
            ("mat_type", "#105 재질 타입 (1:CN 2:RS 3:CM)", 1),
            ("proc_type", "#106 프로세스 (3 or 4)", 3),
            ("rough", "#107 러프 (0:OFF 1:ON)", 0),
            ("single", "#108 싱글 (0:OFF 1:ON)", 1),
        ]
        
        for key, label, default in fields_a:
            self.add_input_row(parent, row, key, label, default)
            row += 1
        
        row += 1
        
        # B. DIMENSIONS
        ttk.Label(parent, text="B. 치수 (DIMENSIONS)", font=('Arial', 10, 'bold')).grid(
            row=row, column=0, columnspan=2, sticky=tk.W, pady=(10, 5)
        )
        row += 1
        
        fields_b = [
            ("raw_od", "#109 원소재 외경 (RAW OD)", 70),
            ("raw_id", "#110 원소재 내경 (RAW ID)", 56),
            ("fin_od", "#111 완제품 외경 (FIN OD)", 64.90),
            ("fin_id", "#112 완제품 내경 (FIN ID)", 58.30),
            ("fin_length", "#113 완제품 길이 (FIN LENGTH)", 7.823),
            ("tip_width", "#114 팁 폭 (TIP WIDTH)", 2.02),
            ("id_r", "#115 내경 R (ID-R)", 0.34),
            ("od_r", "#116 외경 R (OD-R)", 0.36),
            ("id_l", "#117 내경 모따기 (ID-L)", 0.37),
            ("od_l", "#118 외경 모따기 (OD-L)", 0.60),
        ]
        
        for key, label, default in fields_b:
            self.add_input_row(parent, row, key, label, default)
            row += 1
        
        row += 1
        
        # C. CUTTING CONDITIONS
        ttk.Label(parent, text="C. 절삭 조건 (CONDITIONS)", font=('Arial', 10, 'bold')).grid(
            row=row, column=0, columnspan=2, sticky=tk.W, pady=(10, 5)
        )
        row += 1
        
        fields_c = [
            ("t01_rpm", "#120 T01 RPM", 1700),
            ("t02_rpm", "#121 T02 RPM", 1700),
            ("pull_dist", "#122 풀 거리", 1),
            ("margin", "#123 마진", 0.2),
        ]
        
        for key, label, default in fields_c:
            self.add_input_row(parent, row, key, label, default)
            row += 1
        
        row += 1
        
        # D. SYSTEM PARAMETERS
        ttk.Label(parent, text="D. 시스템 변수 (SYSTEM)", font=('Arial', 10, 'bold')).grid(
            row=row, column=0, columnspan=2, sticky=tk.W, pady=(10, 5)
        )
        row += 1
        
        fields_d = [
            ("safety_length", "#530 안전 길이", 12),
            ("chuck_left_limit", "#531 척 왼쪽 한계", 15),
            ("chuck2jaw", "#532 척투조", 1.03),
            ("z_rough_max", "#506 Z축 러프 최대 깊이", 55),
            ("z_rough_free", "#507 Z축 러프 여유", 1),
            ("t01_nose_r", "#508 T01 팁 노즈 R", 0.2),
            ("rough_int", "#509 러프 간격", 0.4),
            ("chamfer_int", "#510 모따기 간격", 1),
        ]
        
        for key, label, default in fields_d:
            self.add_input_row(parent, row, key, label, default)
            row += 1
    
    def add_input_row(self, parent, row, key, label, default):
        """입력 행 추가"""
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, pady=2)
        
        entry = ttk.Entry(parent, width=15)
        entry.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=2, padx=(10, 0))
        entry.insert(0, str(default))
        
        self.entries[key] = entry
    
    def load_default_values(self):
        """기본값 로드"""
        defaults = {
            "mat_oo": 60,
            "mat_ooo": 500,
            "chuck_length": 75,
            "init_face_cut": 1,
            "mat_type": 1,
            "proc_type": 3,
            "rough": 0,
            "single": 1,
            "raw_od": 70,
            "raw_id": 56,
            "fin_od": 64.90,
            "fin_id": 58.30,
            "fin_length": 7.823,
            "tip_width": 2.02,
            "id_r": 0.34,
            "od_r": 0.36,
            "id_l": 0.37,
            "od_l": 0.60,
            "t01_rpm": 1700,
            "t02_rpm": 1700,
            "pull_dist": 1,
            "margin": 0.2,
            "safety_length": 12,
            "chuck_left_limit": 15,
            "chuck2jaw": 1.03,
            "z_rough_max": 55,
            "z_rough_free": 1,
            "t01_nose_r": 0.2,
            "rough_int": 0.4,
            "chamfer_int": 1,
        }
        
        for key, value in defaults.items():
            if key in self.entries:
                self.entries[key].delete(0, tk.END)
                self.entries[key].insert(0, str(value))
    
    def get_input_values(self):
        """입력값 가져오기"""
        try:
            values = {}
            for key, entry in self.entries.items():
                value_str = entry.get().strip()
                values[key] = float(value_str) if value_str else 0
            return values
        except ValueError as e:
            messagebox.showerror("입력 오류", f"숫자만 입력 가능합니다.\n{str(e)}")
            return None
    
    def calculate(self):
        """계산 실행"""
        values = self.get_input_values()
        if values is None:
            return
        
        # 계산기에 파라미터 설정
        self.calculator.set_parameters(**values)
        
        # 결과 생성
        result = self.generate_report()
        
        # 결과 표시
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(1.0, result)
        
        # 문제 발견 시 하이라이트
        if "⚠️  문제 발견!" in result:
            self.highlight_warning()
    
    def highlight_warning(self):
        """경고 부분 하이라이트"""
        self.result_text.tag_config("warning", background="yellow", foreground="red")
        
        start_idx = self.result_text.search("⚠️  문제 발견!", 1.0, tk.END)
        if start_idx:
            end_idx = self.result_text.search("✅ 권장 수정:", start_idx, tk.END)
            if end_idx:
                # 다음 줄까지 포함
                end_idx = f"{end_idx}+3l"
                self.result_text.tag_add("warning", start_idx, end_idx)
    
    def generate_report(self):
        """리포트 생성"""
        calc = self.calculator
        
        report = []
        report.append("=" * 70)
        report.append("G-CODE 파라미터 계산 결과")
        report.append("=" * 70)
        
        # 계산된 주요 값
        report.append("\n[계산된 주요 값]")
        report.append("=" * 70)
        
        safety_res = calc.calculate_safety_override()
        report.append(f"안전 잔량:             #129 = {safety_res}")
        
        total_length = calc.calculate_total_length()
        report.append(f"총 소재 길이:          #140 = {total_length:.3f} mm")
        
        unit_length = calc.calculate_unit_length()
        report.append(f"제품 1개 길이:         #505 = {unit_length:.3f} mm")
        
        eff_chuck = calc.calculate_effective_chuck_length()
        report.append(f"척 유효 길이:          #500 = {eff_chuck:.3f} mm")
        
        pieces_per_cycle = calc.calculate_pieces_per_cycle()
        report.append(f"사이클당 제품 개수:    #517 = {pieces_per_cycle} 개")
        
        cycle_length = calc.calculate_cycle_length()
        report.append(f"사이클 가공 길이:      #518 = {cycle_length:.3f} mm")
        
        autolink_count = calc.calculate_autolink_count()
        report.append(f"\n★ 자동링크 횟수:      #519 = {autolink_count} 회")
        
        # 가공 프로세스
        report.append("\n" + "=" * 70)
        report.append("[가공 프로세스 시뮬레이션]")
        report.append("=" * 70)
        
        init_face = calc.params.get('init_face_cut', 0)
        tip_width = calc.params.get('tip_width', 0)
        
        report.append(f"\n1. 초기 면 절삭:        {init_face:.3f} mm")
        report.append(f"2. 첫 사이클 가공:      {cycle_length:.3f} mm (제품 {pieces_per_cycle}개)")
        report.append(f"   - 제품 길이:         {cycle_length:.3f} mm")
        report.append(f"   - 팁 폭:             {tip_width:.3f} mm")
        report.append(f"   총: {cycle_length + tip_width:.3f} mm")
        
        report.append(f"\n3. 자동링크 {autolink_count}회:")
        for i in range(min(autolink_count, 10)):  # 최대 10개만 표시
            report.append(f"   [{i+1}회] 이송 → 가공 {cycle_length:.3f} mm (제품 {pieces_per_cycle}개)")
        
        if autolink_count > 10:
            report.append(f"   ... ({autolink_count - 10}회 더)")
        
        used_length = calc.calculate_used_length_after_autolink(autolink_count)
        report.append(f"\n총 사용 길이:          #528 = {used_length:.3f} mm")
        
        total_pieces = pieces_per_cycle * (autolink_count + 1)
        report.append(f"총 생산 제품:          {total_pieces} 개")
        
        # 남은 소재 분석
        report.append("\n" + "=" * 70)
        report.append("[자동링크 후 남은 소재 분석]")
        report.append("=" * 70)
        
        physical_remain = total_length - used_length
        report.append(f"\n물리적 남은 길이:      {physical_remain:.3f} mm")
        
        chuck_left = calc.params.get('chuck_left_limit', 0)
        chuck2jaw = calc.params.get('chuck2jaw', 0)
        safety_length = calc.params.get('safety_length', 0)
        margin = calc.params.get('margin', 0)
        
        report.append(f"\n차감 항목:")
        report.append(f"  - 척 왼쪽 한계:       {chuck_left:.3f} mm")
        report.append(f"  - 척투조:             {chuck2jaw:.3f} mm")
        report.append(f"  - 안전 길이:          {safety_length:.3f} mm")
        report.append(f"  - 팁 폭:              {tip_width:.3f} mm")
        report.append(f"  - 마진:               {margin:.3f} mm")
        total_deduct = chuck_left + chuck2jaw + safety_length + tip_width + margin
        report.append(f"  총 차감:              {total_deduct:.3f} mm")
        
        remaining, pieces = calc.calculate_remaining_material(autolink_count)
        report.append(f"\n★ 유효 가공 길이:     #538 = {remaining:.3f} mm")
        report.append(f"★ 추가 가능 제품:     #539 = {pieces} 개")
        
        # N220 로직 검증
        report.append("\n" + "=" * 70)
        report.append("[N220 로직 검증]")
        report.append("=" * 70)
        
        report.append(f"\n현재 코드:")
        report.append(f"  IF [#539 GT 0] GOTO 221;")
        report.append(f"  N221: GOTO 100;")
        
        if pieces > 0:
            report.append(f"\n⚠️  문제 발견!")
            report.append(f"  - #539 = {pieces} > 0")
            report.append(f"  - N221로 점프 → N100으로 복귀")
            report.append(f"  - N230(9003) 건너뛰기")
            report.append(f"  - 결과: 제품 {pieces}개분 소재 낭비! ({remaining:.3f} mm)")
            report.append(f"\n✅ 권장 수정:")
            report.append(f"  IF [#539 EQ 0] GOTO 221;")
            report.append(f"  GOTO 230;  ← 9003 호출하여 남은 소재 가공")
        else:
            report.append(f"\n✓ 문제 없음")
            report.append(f"  - #539 = 0 (남은 소재 없음)")
            report.append(f"  - 정상 종료")
        
        report.append("\n" + "=" * 70)
        
        return "\n".join(report)


def main():
    root = tk.Tk()
    app = GCodeGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
