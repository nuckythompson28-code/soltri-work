"""
작업복_지급내역.md → legacy_records.json + legacy_issues.md

피복지급대장(메인 원장) + 단체피복지급내역서(보조) + 조끼/춘추티 대장(보조)을 파싱하여
Apps Script 일괄 import용 JSON과 검토용 이슈 리포트를 생성합니다.

Usage:
    python ingest_legacy.py
"""
from __future__ import annotations
import json
import re
from pathlib import Path
from collections import Counter

ROOT = Path(__file__).parent
SRC = ROOT / '작업복_지급내역.md'
OUT_JSON = ROOT / 'legacy_records.json'
OUT_ISSUES = ROOT / 'legacy_issues.md'


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
    # 단체피복지급내역서 variants
    '춘추잠바사이즈': '춘추 점퍼',
    '춘추바지사이즈': '춘추 바지',
    '춘추긴팔티': '춘추 티',
    '여름반팔티': '하복 티',
    '조끼사이즈': '동 조끼',
    '기모바지': '동 바지',
    '춘추티사이즈': '춘추 티',
}

# 이름 중복/변형 정규화
NAME_ALIAS = {
    '아자할,17.6.16': '아자할',
    '아자할(방글라데시)': '아자할',
    '사루알방글라야간조': '사루알',
    '사루알(방글라데시)': '사루알',
    '조이방글라2016.3.22': '조이',
    '조이(방글라데시)': '조이',
    '라틴송림': '라틴',
    '수몬송림': '수몬',
    '호세인(송림)': '호세인',
    '김대호(송림)': '김대호',
    '김윤호(송림)': '김윤호',
    '모탈렙(송림)': '모탈렙',
    '자카리아(송림)': '자카리아',
    '알리송림': '알리',
    '강현진(송림)': '강현진',
    '문경호(송림동) 퇴사': '문경호',
    '이선봉퇴사': '이선봉',
    '배인석퇴사': '배인석',
    '마셀로퇴사': '마셀로',
    '일용직홍금표파견직': '홍금표',
}


def normalize_name(s: str | None) -> str | None:
    if s is None:
        return None
    s = s.strip()
    if not s or s.lower() == 'nan':
        return None
    return NAME_ALIAS.get(s, s)


def normalize_item(s: str | None) -> tuple[str | None, str | None]:
    """(정규화된 품목명, 미매핑 원본) — 매핑 실패 시 두 번째가 원본."""
    if s is None:
        return None, None
    # 개행과 공백 정리
    s = s.replace('\\n', '').replace('\n', '').strip()
    if not s or s.lower() == 'nan':
        return None, None
    if s in ITEM_NAME_MAP:
        return ITEM_NAME_MAP[s], None
    # 변형 처리: 퇴사자 섹션 등에서 쓰이는 축약 이름
    short_map = {
        '동상': '동 점퍼',
        '동하': '동 바지',
        '춘추상': '춘추 점퍼',
        '춘추하': '춘추 바지',
        '춘추 상': '춘추 점퍼',
        '춘추 하': '춘추 바지',
        '하복상': '하복 티',
        '하복하': '하복 바지',
        '반팔티': '하복 티',
        '긴팔티': '춘추 티',
        '겨울점퍼': '동 점퍼',
        '겨울바지': '동 바지',
        '춘추점퍼': '춘추 점퍼',
        '안전화': '4인치 안전화',      # 구형 표기, 대체로 4인치
        'K2 안전화': '6인치 안전화',    # 브랜드+6인치 대부분
    }
    if s in short_map:
        return short_map[s], None
    return None, s  # 미매핑


def normalize_size(s: str | None) -> tuple[str | None, str | None]:
    """(정규화 사이즈, 경고) — 불확실한 경우 경고 반환."""
    if s is None:
        return None, None
    s = s.strip()
    if not s or s.lower() == 'nan':
        return None, None
    # "XL-100" → "XL"  /  "XXXL-110" → "XXXL"
    m = re.match(r'^([xXsmlSML]+|\d+)-\d+$', s)
    if m:
        return m.group(1).upper() if m.group(1).isalpha() else m.group(1), None
    # "L,M" 복수 → 첫 번째
    if ',' in s:
        first = s.split(',')[0].strip()
        return first.upper() if first.isalpha() else first, f'multi-size: "{s}" → "{first}"'
    # "xL" / "xXL" → "XL" / "XXL"
    if re.match(r'^[xXlLsSmM0-9]+$', s):
        return s.upper() if s.isalpha() else s, None
    # "3~5" — 범위 표기, 애매
    if '~' in s:
        return None, f'range-size: "{s}"'
    # 사이즈 자리에 날짜가 들어간 경우 (예: "22.10.21")
    if re.match(r'^\d{1,2}\.\d{1,2}\.\d{1,2}$', s):
        return None, f'date-in-size: "{s}"'
    # 기본: 대문자 정규화
    return s.upper(), None


def parse_date(s: str | None) -> tuple[str | None, str | None]:
    """(YYYY-MM-DD, 이슈) — 파싱 실패 시 이슈 반환."""
    if s is None:
        return None, 'empty'
    s = s.strip()
    if not s or s.lower() == 'nan':
        return None, 'empty'
    # ISO 2018-12-17 또는 2018-12-17 00:00:00
    m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 2000:
            return None, f'suspicious_year_{y}: "{s}" (엑셀 serial date 오류 추정)'
        return f'{y:04d}-{mo:02d}-{d:02d}', None
    # 2자리 연도: 18.12.17 / 22.6.17 / 2016.3.22
    m = re.match(r'^(\d{1,4})\.(\d{1,2})\.(\d{1,2})$', s)
    if m:
        y = int(m.group(1))
        if y < 50:
            y += 2000
        elif y < 100:
            y += 1900
        return f'{y:04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}', None
    # "10월 지급" / "3월 지급" — 예정
    if '월' in s or '지급' in s:
        return None, 'future_schedule'
    return None, f'unparseable: "{s}"'


def parse_qty(s: str | None) -> tuple[int | None, str | None]:
    if s is None:
        return None, None
    s = s.strip()
    if not s or s.lower() == 'nan':
        return None, None
    if s.isdigit():
        return int(s), None
    # "OK" 등 → 1로 추정하되 경고
    if s.lower() in ('ok', 'o', 'y'):
        return 1, f'qty_from_confirm: "{s}" assumed 1'
    return None, f'unparseable_qty: "{s}"'


def split_row(line: str) -> list[str]:
    """| a | b | c | → ['a', 'b', 'c']."""
    parts = line.strip().strip('|').split('|')
    return [p.strip() for p in parts]


def is_separator(line: str) -> bool:
    return bool(re.match(r'^\|\s*-+', line))


def parse_main_ledger(lines: list[str]) -> tuple[list[dict], list[dict]]:
    """피복지급대장: l368-1344 파싱."""
    records: list[dict] = []
    issues: list[dict] = []
    in_section = False
    data_started = False
    current_emp = None  # {'name':..., 'dept':..., 'joinDate':...}
    skip_retired_section = False

    for idx, line in enumerate(lines):
        if line.startswith('## 피복지급대장'):
            in_section = True
            continue
        if in_section and line.startswith('## '):
            break  # 다음 섹션 시작
        if not in_section:
            continue
        if not line.strip().startswith('|'):
            continue
        if is_separator(line):
            continue

        cols = split_row(line)
        if len(cols) < 10:
            continue

        # 헤더 행 (NO. / 이름 / ...) 건너뛰기
        if cols[0] in ('NO.', '번 호') or cols[0].startswith('피복지급대장'):
            data_started = True
            continue
        if not data_started:
            continue

        # "퇴사자" 구분자 → 이후 행은 계속 처리하되 섹션 표시
        if cols[0] == '퇴사자' or (cols[0] == '' and cols[1] == '퇴사자'):
            skip_retired_section = True
            continue

        # 직원 시작 판정:
        #  - 현역 섹션: NO.가 숫자일 때만 새 직원
        #  - 퇴사자 섹션: NO가 숫자거나, NO는 비어도 이름이 있으면 새 직원 (각 행이 별개 퇴사자)
        no_col = cols[0]
        name_col = cols[1]
        name_nonempty = name_col and name_col.lower() != 'nan'

        if no_col and no_col.isdigit() and name_nonempty:
            current_emp = {
                'no': int(no_col),
                'name': normalize_name(name_col) or name_col,
                'joinDate': cols[2] if len(cols) > 2 else '',
                'dept': cols[3] if len(cols) > 3 else '',
                'retired': skip_retired_section,
            }
        elif skip_retired_section and name_nonempty:
            # 퇴사자 섹션에서 NO 없이 이름만 있는 행 → 각 행이 별개 직원
            current_emp = {
                'no': None,
                'name': normalize_name(name_col) or name_col,
                'joinDate': cols[2] if len(cols) > 2 else '',
                'dept': cols[3] if len(cols) > 3 else '',
                'retired': True,
            }

        if current_emp is None:
            continue

        # 종류(품목) 파싱
        item_raw = cols[4] if len(cols) > 4 else ''
        item_name, item_unmapped = normalize_item(item_raw)
        if not item_name:
            if item_unmapped:
                issues.append({
                    'line': idx + 1,
                    'emp': current_emp.get('name'),
                    'type': 'unmapped_item',
                    'detail': item_unmapped,
                })
            continue

        # 15 사이클 순회: 사이즈(5) | 수량(6) | 지급일자(7) | 확인(8) — 반복
        for c in range(15):
            base = 5 + c * 4
            if base + 2 >= len(cols):
                break
            size_raw = cols[base]
            qty_raw = cols[base + 1]
            date_raw = cols[base + 2]

            size_norm, size_warn = normalize_size(size_raw)
            qty, qty_warn = parse_qty(qty_raw)
            date, date_issue = parse_date(date_raw)

            # 전부 빈 셀이면 조용히 스킵
            all_empty = all(
                v is None or (isinstance(v, str) and (not v.strip() or v.strip().lower() == 'nan'))
                for v in (size_raw, qty_raw, date_raw)
            )
            if all_empty:
                continue

            # 필수: 날짜
            if not date:
                if date_issue and date_issue != 'empty':
                    issues.append({
                        'line': idx + 1,
                        'emp': current_emp['name'],
                        'item': item_name,
                        'cycle': c + 1,
                        'type': 'bad_date',
                        'detail': date_issue,
                        'raw': {'size': size_raw, 'qty': qty_raw, 'date': date_raw},
                    })
                continue

            # 수량 없으면 1로 가정 + 이슈
            if qty is None:
                qty = 1
                issues.append({
                    'line': idx + 1,
                    'emp': current_emp['name'],
                    'item': item_name,
                    'cycle': c + 1,
                    'type': 'missing_qty',
                    'detail': f'qty 칸이 "{qty_raw}" → 1로 가정',
                    'raw': {'size': size_raw, 'qty': qty_raw, 'date': date_raw},
                })
            elif qty_warn:
                issues.append({
                    'line': idx + 1,
                    'emp': current_emp['name'],
                    'item': item_name,
                    'cycle': c + 1,
                    'type': 'qty_warn',
                    'detail': qty_warn,
                })

            # 사이즈 경고
            if size_warn:
                issues.append({
                    'line': idx + 1,
                    'emp': current_emp['name'],
                    'item': item_name,
                    'cycle': c + 1,
                    'type': 'size_warn',
                    'detail': size_warn,
                })

            rec = {
                'empName': current_emp['name'],
                'dept': current_emp['dept'] if current_emp['dept'].lower() != 'nan' else '',
                'date': date,
                'item': item_name,
                'size': size_norm or '',
                'qty': qty,
                'source': 'ledger',
                'retired': current_emp.get('retired', False),
            }
            records.append(rec)

    return records, issues


def main():
    raw = SRC.read_text(encoding='utf-8')
    lines = raw.splitlines()

    records, issues = parse_main_ledger(lines)

    # 통계
    by_item = Counter(r['item'] for r in records)
    by_year = Counter(r['date'][:4] for r in records)
    by_emp = Counter(r['empName'] for r in records)
    unique_emps = len(by_emp)
    retired = sum(1 for r in records if r.get('retired'))

    issue_by_type = Counter(i['type'] for i in issues)

    # 레코드 ID 부여
    for i, r in enumerate(records):
        r['id'] = f'legacy-{i+1:04d}'
        r['time'] = '00:00'

    OUT_JSON.write_text(
        json.dumps({'records': records, 'stats': {
            'total': len(records),
            'unique_employees': unique_emps,
            'retired_records': retired,
            'by_item': dict(by_item.most_common()),
            'by_year': dict(sorted(by_year.items())),
        }}, ensure_ascii=False, indent=2),
        encoding='utf-8',
    )

    # 이슈 리포트
    lines_out = [
        '# 레거시 지급내역 인제스트 리포트',
        '',
        f'- 원본: `작업복_지급내역.md`',
        f'- 파싱 섹션: 피복지급대장',
        f'- **생성된 레코드**: {len(records)}건',
        f'- **고유 직원**: {unique_emps}명 (퇴사자 포함)',
        f'- **퇴사자 레코드**: {retired}건',
        '',
        '## 품목별',
        '',
    ]
    for item, cnt in by_item.most_common():
        lines_out.append(f'- {item}: {cnt}건')
    lines_out += ['', '## 연도별', '']
    for y, cnt in sorted(by_year.items()):
        lines_out.append(f'- {y}년: {cnt}건')
    lines_out += ['', '## 직원별 (상위 20)', '']
    for emp, cnt in by_emp.most_common(20):
        lines_out.append(f'- {emp}: {cnt}건')
    lines_out += ['', f'## 이슈 ({len(issues)}건)', '']
    for itype, cnt in issue_by_type.most_common():
        lines_out.append(f'- **{itype}**: {cnt}건')
    lines_out += ['', '### 이슈 상세 (처음 100건)', '']
    for i in issues[:100]:
        detail = i.get('detail', '')
        raw = i.get('raw', '')
        lines_out.append(
            f"- L{i['line']} | {i.get('emp','?')} | {i.get('item','?')} | {i['type']}: {detail} {raw if raw else ''}"
        )

    OUT_ISSUES.write_text('\n'.join(lines_out), encoding='utf-8')

    import sys
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    print(f'[OK] {len(records)} records -> {OUT_JSON.name}')
    print(f'     unique employees: {unique_emps}, issues: {len(issues)}')
    print(f'     report -> {OUT_ISSUES.name}')


if __name__ == '__main__':
    main()
