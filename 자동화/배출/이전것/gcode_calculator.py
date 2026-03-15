#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
G-code 파라미터 자동 계산 프로그램
O0852/O9001/O9002 로직 분석 및 검증
"""

import math

class GCodeCalculator:
    def __init__(self):
        # 입력 파라미터
        self.params = {}
        
    def set_parameters(self, **kwargs):
        """파라미터 설정"""
        self.params.update(kwargs)
        
    def calculate_safety_override(self):
        """안전 잔량 계산 (#129)"""
        mat_type = self.params.get('mat_type', 1)  # #105
        raw_od = self.params.get('raw_od', 70)      # #109
        
        # RS40 또는 소재 직경 30mm 미만이면 안전 잔량 15
        if mat_type == 2 or raw_od < 30:
            return 15
        return 0
    
    def calculate_total_length(self):
        """총 소재 길이 계산 (#140)"""
        mat_oo = self.params.get('mat_oo', 0)       # #101
        mat_ooo = self.params.get('mat_ooo', 0)     # #102
        safety_res = self.calculate_safety_override()  # #129
        
        return mat_oo + mat_ooo - safety_res
    
    def calculate_unit_length(self):
        """제품 1개 길이 (#505)"""
        fin_length = self.params.get('fin_length', 0)  # #113
        tip_width = self.params.get('tip_width', 0)    # #114
        
        return fin_length + tip_width
    
    def calculate_effective_chuck_length(self):
        """척 유효 길이 (#500)"""
        chuck_length = self.params.get('chuck_length', 0)  # #103
        safety_length = self.params.get('safety_length', 0)  # #530
        
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
        total_length = self.calculate_total_length()  # #140
        chuck_left = self.params.get('chuck_left_limit', 0)  # #531
        chuck2jaw = self.params.get('chuck2jaw', 0)  # #532
        safety_length = self.params.get('safety_length', 0)  # #530
        cycle_length = self.calculate_cycle_length()  # #518
        tip_width = self.params.get('tip_width', 0)  # #114
        init_face = self.params.get('init_face_cut', 0)  # #104
        margin = self.params.get('margin', 0)  # #123
        
        # 분자 계산
        numerator = (total_length - chuck_left - chuck2jaw - safety_length - 
                    cycle_length - tip_width - init_face)
        
        # 분모 계산
        denominator = cycle_length + tip_width + margin
        
        if denominator == 0:
            return 0
        
        return int(numerator / denominator)
    
    def calculate_used_length_after_autolink(self, autolink_count):
        """자동링크 N회 후 사용된 총 길이"""
        init_face = self.params.get('init_face_cut', 0)  # #104
        cycle_length = self.calculate_cycle_length()  # #518
        tip_width = self.params.get('tip_width', 0)  # #114
        margin = self.params.get('margin', 0)  # #123
        
        # 첫 사이클
        first_cycle = cycle_length + tip_width
        
        # 자동링크 사이클들
        autolink_cycles = autolink_count * (cycle_length + tip_width + margin)
        
        # 총 사용 길이 (#528)
        return init_face + first_cycle + autolink_cycles
    
    def calculate_remaining_material(self, autolink_count):
        """자동링크 후 남은 유효 소재 (#538, #539)"""
        total_length = self.calculate_total_length()  # #140
        used_length = self.calculate_used_length_after_autolink(autolink_count)  # #528
        chuck_left = self.params.get('chuck_left_limit', 0)  # #531
        chuck2jaw = self.params.get('chuck2jaw', 0)  # #532
        safety_length = self.params.get('safety_length', 0)  # #530
        tip_width = self.params.get('tip_width', 0)  # #114
        margin = self.params.get('margin', 0)  # #123
        
        # 남은 유효 길이 (#538)
        remaining = (total_length - used_length) - (chuck_left + chuck2jaw + safety_length) - (tip_width + margin)
        
        # 만들 수 있는 제품 개수 (#539)
        unit_length = self.calculate_unit_length()
        if unit_length == 0:
            pieces = 0
        else:
            pieces = int(remaining / unit_length)
        
        return remaining, pieces
    
    def print_report(self):
        """계산 결과 리포트 출력"""
        print("=" * 70)
        print("G-CODE 파라미터 계산 결과")
        print("=" * 70)
        
        print("\n[입력 파라미터]")
        print(f"  소재 길이 (OO):        #101 = {self.params.get('mat_oo', 0)}")
        print(f"  소재 길이 (OOO):       #102 = {self.params.get('mat_ooo', 0)}")
        print(f"  척 길이:               #103 = {self.params.get('chuck_length', 0)}")
        print(f"  초기 면 절삭:          #104 = {self.params.get('init_face_cut', 0)}")
        print(f"  재질 타입:             #105 = {self.params.get('mat_type', 0)}")
        print(f"  완제품 직경 (OD):      #111 = {self.params.get('fin_od', 0)}")
        print(f"  완제품 내경 (ID):      #112 = {self.params.get('fin_id', 0)}")
        print(f"  완제품 길이:           #113 = {self.params.get('fin_length', 0)}")
        print(f"  팁 폭:                 #114 = {self.params.get('tip_width', 0)}")
        print(f"  마진:                  #123 = {self.params.get('margin', 0)}")
        
        print("\n[시스템 파라미터]")
        print(f"  안전 길이:             #530 = {self.params.get('safety_length', 0)}")
        print(f"  척 왼쪽 한계:          #531 = {self.params.get('chuck_left_limit', 0)}")
        print(f"  척투조:                #532 = {self.params.get('chuck2jaw', 0)}")
        
        print("\n" + "=" * 70)
        print("[계산된 값]")
        print("=" * 70)
        
        safety_res = self.calculate_safety_override()
        print(f"\n  안전 잔량:             #129 = {safety_res}")
        
        total_length = self.calculate_total_length()
        print(f"  총 소재 길이:          #140 = {total_length:.3f} mm")
        
        unit_length = self.calculate_unit_length()
        print(f"  제품 1개 길이:         #505 = {unit_length:.3f} mm")
        
        eff_chuck = self.calculate_effective_chuck_length()
        print(f"  척 유효 길이:          #500 = {eff_chuck:.3f} mm")
        
        pieces_per_cycle = self.calculate_pieces_per_cycle()
        print(f"  사이클당 제품 개수:    #517 = {pieces_per_cycle} 개")
        
        cycle_length = self.calculate_cycle_length()
        print(f"  사이클 가공 길이:      #518 = {cycle_length:.3f} mm")
        
        autolink_count = self.calculate_autolink_count()
        print(f"\n  ★ 자동링크 횟수:      #519 = {autolink_count} 회")
        
        print("\n" + "=" * 70)
        print("[가공 프로세스 시뮬레이션]")
        print("=" * 70)
        
        init_face = self.params.get('init_face_cut', 0)
        tip_width = self.params.get('tip_width', 0)
        margin = self.params.get('margin', 0)
        
        print(f"\n  1. 초기 면 절삭:        {init_face:.3f} mm")
        print(f"  2. 첫 사이클 가공:      {cycle_length:.3f} mm (제품 {pieces_per_cycle}개)")
        print(f"     - 제품 길이:         {cycle_length:.3f} mm")
        print(f"     - 팁 폭:             {tip_width:.3f} mm")
        print(f"     총: {cycle_length + tip_width:.3f} mm")
        
        print(f"\n  3. 자동링크 {autolink_count}회:")
        for i in range(autolink_count):
            print(f"     [{i+1}회] 이송 → 가공 {cycle_length:.3f} mm (제품 {pieces_per_cycle}개)")
        
        used_length = self.calculate_used_length_after_autolink(autolink_count)
        print(f"\n  총 사용 길이:          #528 = {used_length:.3f} mm")
        
        total_pieces = pieces_per_cycle * (autolink_count + 1)
        print(f"  총 생산 제품:          {total_pieces} 개")
        
        print("\n" + "=" * 70)
        print("[자동링크 후 남은 소재 분석]")
        print("=" * 70)
        
        physical_remain = total_length - used_length
        print(f"\n  물리적 남은 길이:      {physical_remain:.3f} mm")
        
        chuck_left = self.params.get('chuck_left_limit', 0)
        chuck2jaw = self.params.get('chuck2jaw', 0)
        safety_length = self.params.get('safety_length', 0)
        
        print(f"\n  차감 항목:")
        print(f"    - 척 왼쪽 한계:       {chuck_left:.3f} mm")
        print(f"    - 척투조:             {chuck2jaw:.3f} mm")
        print(f"    - 안전 길이:          {safety_length:.3f} mm")
        print(f"    - 팁 폭:              {tip_width:.3f} mm")
        print(f"    - 마진:               {margin:.3f} mm")
        total_deduct = chuck_left + chuck2jaw + safety_length + tip_width + margin
        print(f"    총 차감:              {total_deduct:.3f} mm")
        
        remaining, pieces = self.calculate_remaining_material(autolink_count)
        print(f"\n  ★ 유효 가공 길이:     #538 = {remaining:.3f} mm")
        print(f"  ★ 추가 가능 제품:     #539 = {pieces} 개")
        
        print("\n" + "=" * 70)
        print("[N220 로직 검증]")
        print("=" * 70)
        
        print(f"\n  현재 코드:")
        print(f"    IF [#539 GT 0] GOTO 221;")
        print(f"    N221: GOTO 100;")
        
        if pieces > 0:
            print(f"\n  ⚠️  문제 발견!")
            print(f"  - #539 = {pieces} > 0")
            print(f"  - N221로 점프 → N100으로 복귀")
            print(f"  - N230(9003) 건너뛰기")
            print(f"  - 결과: 제품 {pieces}개분 소재 낭비! ({remaining:.3f} mm)")
            print(f"\n  ✅ 권장 수정:")
            print(f"    IF [#539 EQ 0] GOTO 221;")
            print(f"    GOTO 230;  ← 9003 호출하여 남은 소재 가공")
        else:
            print(f"\n  ✓ 문제 없음")
            print(f"  - #539 = 0 (남은 소재 없음)")
            print(f"  - 정상 종료")
        
        print("\n" + "=" * 70)


def main():
    """메인 함수"""
    calc = GCodeCalculator()
    
    # 파라미터 설정 (사용자 입력값)
    calc.set_parameters(
        # A. SETTINGS
        mat_oo=60,           # #101
        mat_ooo=500,         # #102
        chuck_length=75,     # #103
        init_face_cut=1,     # #104
        mat_type=1,          # #105 (1:CN 2:RS 3:CM)
        proc_type=3,         # #106
        rough=0,             # #107
        single=1,            # #108
        
        # B. DIMENSIONS
        raw_od=70,           # #109
        raw_id=56,           # #110
        fin_od=64.90,        # #111
        fin_id=58.30,        # #112
        fin_length=7.823,    # #113
        tip_width=2.02,      # #114
        id_r=0.34,           # #115
        od_r=0.36,           # #116
        id_l=0.37,           # #117
        od_l=0.60,           # #118
        
        # C. CUTTING CONDITIONS
        t01_rpm=1700,        # #120
        t02_rpm=1700,        # #121
        pull_dist=1,         # #122
        margin=0.2,          # #123
        
        # SYSTEM VARS
        safety_length=12,    # #530
        chuck_left_limit=15, # #531
        chuck2jaw=1.03,      # #532
        z_rough_max=55,      # #506
        z_rough_free=1,      # #507
        t01_nose_r=0.2,      # #508
        rough_int=0.4,       # #509
        chamfer_int=1,       # #510
    )
    
    # 계산 및 리포트 출력
    calc.print_report()


if __name__ == "__main__":
    main()
