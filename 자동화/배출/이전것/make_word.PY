import os
import random
import pandas as pd
from datetime import datetime, timedelta
from docx import Document
from docx.oxml.ns import qn

def replace_text_with_font(doc, old_text, new_text):
    all_paragraphs = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                all_paragraphs.extend(cell.paragraphs)
    for p in all_paragraphs:
        if old_text in p.text:
            for run in p.runs:
                if old_text in run.text:
                    val = "" if pd.isna(new_text) else str(new_text)
                    run.text = run.text.replace(old_text, val)
                    run.font.name = '바탕'
                    r = run._element.rPr
                    r.rFonts.set(qn('w:eastAsia'), '바탕')

def get_weather_data(month):
    temp_map = {1: (-5, 2), 2: (-3, 5), 3: (2, 11), 4: (8, 18), 5: (14, 24), 6: (19, 27),
                7: (23, 30), 8: (24, 31), 9: (18, 26), 10: (11, 20), 11: (4, 12), 12: (-2, 4)}
    weather_list = ["맑음", "맑음", "맑음", "흐림", "구름조금"]
    min_t, max_t = temp_map[month]
    return random.choice(weather_list), str(random.randint(min_t, max_t))

def generate_records():
    # ================= [ 날짜 설정 구역 ] =================
    # 만들고 싶은 시작일과 종료일을 입력하세요 (형식: YYYY-MM-DD)
    START_DATE_STR = "2026-01-01" 
    END_DATE_STR = "2026-12-31"   # 일단 넉넉히 잡아두셔도 됩니다.
    # =====================================================

    template_path = 'template.docx'
    csv_path = 'data.csv'
    base_dir = "운영기록부_생성결과"

    if not os.path.exists(csv_path):
        print(f"에러: CSV 파일이 없습니다.")
        return

    try:
        df = pd.read_csv(csv_path, encoding='cp949')
    except:
        df = pd.read_csv(csv_path, encoding='utf-8')

    df['날짜'] = pd.to_datetime(df['날짜']).dt.strftime('%Y-%m-%d')
    df.set_index('날짜', inplace=True)

    # 문자열을 날짜 객체로 변환
    current_date = datetime.strptime(START_DATE_STR, "%Y-%m-%d")
    end_limit = datetime.strptime(END_DATE_STR, "%Y-%m-%d")
    
    days_ko = ["월", "화", "수", "목", "금", "토", "일"]

    print("--- 생성을 시작합니다 ---")

    while current_date <= end_limit:
        date_key = current_date.strftime('%Y-%m-%d')
        
        # [중요] CSV에 오늘 날짜 데이터가 없으면 중단!
        if date_key not in df.index:
            print(f">>> {date_key} 데이터가 CSV에 없어 작업을 중단합니다.")
            break

        month_str = f"{current_date.month:02d}월"
        save_dir = os.path.join(base_dir, month_str)
        if not os.path.exists(save_dir): os.makedirs(save_dir)

        doc = Document(template_path)
        wth, tmp = get_weather_data(current_date.month)
        is_weekend = current_date.weekday() >= 5
        
        # 데이터 매칭
        row = df.loc[date_key]
        if isinstance(row, pd.DataFrame): row = row.iloc[0]
        vals = {f"VAL{i}": row[f"전력{i}"] for i in range(1, 9)}

        if is_weekend:
            t17, t18 = "", ""
        else:
            t17, t18 = "07:00 ~ 12:00, 13:00~17:00", "07:00 ~ 12:00, 13:00~18:00"

        # 텍스트 교체
        replace_text_with_font(doc, "YYYY", str(current_date.year))
        replace_text_with_font(doc, "MM", str(current_date.month))
        replace_text_with_font(doc, "DD", str(current_date.day))
        replace_text_with_font(doc, "WD", days_ko[current_date.weekday()])
        replace_text_with_font(doc, "WTH", wth)
        replace_text_with_font(doc, "TMP", tmp)
        replace_text_with_font(doc, "TIME17", t17)
        replace_text_with_font(doc, "TIME18", t18)
        for k, v in vals.items(): replace_text_with_font(doc, k, v)

        save_name = f"배출방지시설운영기록부({current_date.strftime('%y%m%d')}) - {days_ko[current_date.weekday()]}.docx"
        doc.save(os.path.join(save_dir, save_name))
        print(f"[{date_key}] 생성 완료")
        
        current_date += timedelta(days=1)

    print("\n--- 모든 작업이 끝났습니다 ---")

if __name__ == "__main__":
    generate_records()