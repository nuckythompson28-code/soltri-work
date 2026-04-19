"""
validate_ingest.py — legacy_records.json vs 원본 '총지급수량' 열 교차검증.

피복지급대장의 각 행에는 (직원, 품목, 총지급수량) 정보가 있음.
JSON에서 같은 (직원, 품목) 조합의 qty를 합산 → 원본 총지급수량과 일치해야 함.

불일치 건을 리포트합니다.

Usage:
    python validate_ingest.py
"""
from __future__ import annotations
import json
import re
import sys
from pathlib import Path
from collections import defaultdict

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

ROOT = Path(__file__).parent
SRC = ROOT / '작업복_지급내역.md'
JSON_FILE = ROOT / 'legacy_records.json'

# ingest_legacy.py와 동일해야 함
ITEM_NAME_MAP = {
    '동 상의(점퍼)': '동 점퍼',
    '동 하의(바지)': '동 바지',
    '동 조끼': '동 조끼',
    '춘추 상의(점퍼)': '춘추 점퍼',
    '춘추 티': '춘추 티',
    '춘추 바지(여름겸용)': '춘추 바지',
    '하복 티': '하복 티',
    '4인치 안전화': '4인치 안전화',
    '6인치 안전화': '6인치 안전화',
    '깔창': '깔창',
    '덮개': '덮개',
    '보안경': '보안경',
    '앞치마': '앞치마',
}
SHORT_MAP = {
    '동상': '동 점퍼', '동하': '동 바지',
    '춘추상': '춘추 점퍼', '춘추하': '춘추 바지',
    '춘추 상': '춘추 점퍼', '춘추 하': '춘추 바지',
    '하복상': '하복 티', '하복하': '하복 바지',
    '반팔티': '하복 티', '긴팔티': '춘추 티',
    '겨울점퍼': '동 점퍼', '겨울바지': '동 바지',
    '춘추점퍼': '춘추 점퍼',
    '안전화': '4인치 안전화', 'K2 안전화': '6인치 안전화',
}
NAME_ALIAS = {
    '아자할,17.6.16': '아자할',
    '아자할(방글라데시)': '아자할',
    '사루알방글라야간조': '사루알',
    '사루알(방글라데시)': '사루알',
    '조이방글라2016.3.22': '조이',
    '조이(방글라데시)': '조이',
    '라틴송림': '라틴', '수몬송림': '수몬',
    '호세인(송림)': '호세인', '김대호(송림)': '김대호',
    '김윤호(송림)': '김윤호', '모탈렙(송림)': '모탈렙',
    '자카리아(송림)': '자카리아', '알리송림': '알리',
    '강현진(송림)': '강현진',
    '문경호(송림동) 퇴사': '문경호',
    '이선봉퇴사': '이선봉',
    '배인석퇴사': '배인석',
    '마셀로퇴사': '마셀로',
    '일용직홍금표파견직': '홍금표',
}


def norm_name(s):
    if s is None: return None
    s = s.strip()
    if not s or s.lower() == 'nan': return None
    return NAME_ALIAS.get(s, s)


def norm_item(s):
    if s is None: return None
    s = s.replace('\\n', '').replace('\n', '').strip()
    if not s or s.lower() == 'nan': return None
    if s in ITEM_NAME_MAP: return ITEM_NAME_MAP[s]
    if s in SHORT_MAP: return SHORT_MAP[s]
    return None


def split_row(line):
    parts = line.strip().strip('|').split('|')
    return [p.strip() for p in parts]


def is_separator(line):
    return bool(re.match(r'^\|\s*-+', line))


def parse_expected_totals(lines):
    """피복지급대장에서 (직원, 품목) → 총지급수량 추출."""
    expected = defaultdict(int)  # (name, item) -> 총
    row_details = []  # [(line, name, item, total_from_source)]

    in_section = False
    data_started = False
    current_name = None
    skip_retired = False

    for idx, line in enumerate(lines):
        if line.startswith('## 피복지급대장'):
            in_section = True
            continue
        if in_section and line.startswith('## '):
            break
        if not in_section or not line.strip().startswith('|') or is_separator(line):
            continue

        cols = split_row(line)
        if len(cols) < 10:
            continue

        if cols[0] in ('NO.', '번 호') or cols[0].startswith('피복지급대장'):
            data_started = True
            continue
        if not data_started:
            continue

        # 퇴사자 구분자
        if cols[0] == '퇴사자' or (cols[0] == '' and cols[1] == '퇴사자'):
            skip_retired = True
            continue

        # ingest_legacy.py와 동일한 규칙:
        #  - 현역 섹션: NO가 숫자일 때만 새 직원
        #  - 퇴사자 섹션: NO가 숫자거나, NO 비어도 이름 있으면 새 직원
        no_col, name_col = cols[0], cols[1]
        name_nonempty = name_col and name_col.lower() != 'nan'

        if no_col and no_col.isdigit() and name_nonempty:
            current_name = norm_name(name_col) or name_col
        elif skip_retired and name_nonempty:
            current_name = norm_name(name_col) or name_col

        if current_name is None:
            continue

        item_raw = cols[4] if len(cols) > 4 else ''
        item_name = norm_item(item_raw)
        if not item_name:
            continue

        # 총지급수량 (마지막 컬럼)
        total_raw = cols[-1] if cols else ''
        try:
            total = int(total_raw)
        except (ValueError, TypeError):
            continue

        # 동일 (직원, 품목)이 여러 행에 분산된 경우 합산 (실제로 퇴사자 섹션에서 중복 등장)
        expected[(current_name, item_name)] += total
        row_details.append((idx + 1, current_name, item_name, total))

    return expected, row_details


def compare(expected, records):
    """JSON 레코드 vs 기대 총량 비교."""
    actual = defaultdict(int)
    for r in records:
        actual[(r['empName'], r['item'])] += r['qty']

    mismatches = []
    matches = 0
    only_in_expected = []
    only_in_actual = []

    all_keys = set(expected.keys()) | set(actual.keys())
    for k in sorted(all_keys):
        exp = expected.get(k, 0)
        act = actual.get(k, 0)
        if exp == act:
            if exp > 0:
                matches += 1
        elif exp > 0 and act == 0:
            only_in_expected.append((k, exp))
        elif exp == 0 and act > 0:
            only_in_actual.append((k, act))
        else:
            mismatches.append((k, exp, act, act - exp))

    return matches, mismatches, only_in_expected, only_in_actual


def main():
    raw = SRC.read_text(encoding='utf-8')
    lines = raw.splitlines()

    with open(JSON_FILE, encoding='utf-8') as f:
        data = json.load(f)
    records = data['records']

    expected, row_details = parse_expected_totals(lines)
    matches, mismatches, only_exp, only_act = compare(expected, records)

    total_expected_sum = sum(expected.values())
    total_actual_sum = sum(r['qty'] for r in records)

    lines_out = [
        '# 레거시 인제스트 검증 리포트',
        '',
        f'- 원본 (피복지급대장) 행 수: **{len(row_details)}행** ((직원,품목) 조합)',
        f'- 원본 총지급수량 합계: **{total_expected_sum}개**',
        f'- JSON 레코드 수: **{len(records)}건**',
        f'- JSON qty 합계: **{total_actual_sum}개**',
        '',
        '## 일치 현황',
        '',
        f'- ✅ **정확히 일치**: {matches}개 (직원×품목 조합)',
        f'- ❌ **숫자 불일치**: {len(mismatches)}개',
        f'- 🟡 **원본에는 있고 JSON엔 없음**: {len(only_exp)}개 (파싱 누락 의심)',
        f'- 🟣 **JSON에만 있고 원본엔 없음**: {len(only_act)}개 (품목 매핑 변경 영향?)',
        '',
    ]

    diff = total_actual_sum - total_expected_sum
    if diff == 0:
        lines_out.append('### ✅ 총 수량 합계 일치')
    else:
        lines_out.append(f'### ⚠️ 총 수량 차이: {diff:+d}개 (JSON {total_actual_sum} vs 원본 {total_expected_sum})')
    lines_out.append('')

    if mismatches:
        lines_out += ['## 숫자 불일치 상세', '']
        lines_out.append('| 직원 | 품목 | 원본 총량 | JSON 합 | 차이 |')
        lines_out.append('|---|---|---|---|---|')
        for (name, item), exp, act, diff in mismatches[:50]:
            lines_out.append(f'| {name} | {item} | {exp} | {act} | {diff:+d} |')
        if len(mismatches) > 50:
            lines_out.append(f'| ... | | | | ({len(mismatches)-50}건 더) |')
        lines_out.append('')

    if only_exp:
        lines_out += ['## 파싱 누락 (원본에는 있지만 JSON 없음)', '']
        lines_out.append('| 직원 | 품목 | 원본 총량 |')
        lines_out.append('|---|---|---|')
        for (name, item), exp in only_exp[:30]:
            lines_out.append(f'| {name} | {item} | {exp} |')
        if len(only_exp) > 30:
            lines_out.append(f'| ... | | ({len(only_exp)-30}건 더) |')
        lines_out.append('')

    if only_act:
        lines_out += ['## JSON만 있는 항목 (원본 총량=0)', '']
        lines_out.append('| 직원 | 품목 | JSON 합 |')
        lines_out.append('|---|---|---|')
        for (name, item), act in only_act[:30]:
            lines_out.append(f'| {name} | {item} | {act} |')
        lines_out.append('')

    out = ROOT / 'legacy_validation.md'
    out.write_text('\n'.join(lines_out), encoding='utf-8')

    print(f'[OK] validation complete -> {out.name}')
    print(f'     matched: {matches}')
    print(f'     mismatches: {len(mismatches)}')
    print(f'     missing in JSON: {len(only_exp)}')
    print(f'     extra in JSON: {len(only_act)}')
    print(f'     qty total: JSON={total_actual_sum} vs Source={total_expected_sum} ({diff:+d})')


if __name__ == '__main__':
    main()
