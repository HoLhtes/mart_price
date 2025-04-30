import pandas as pd
import json
from collections import defaultdict

EXCEL_FILE_PATH = 'mart_price.xlsx'
JSON_OUTPUT_PATH = 'data.json'
SHEET_NAME = 0

def convert_excel_to_json_grouped(excel_path, json_path, sheet_name=0):
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=0)
        expected_columns = ['상품 이름', '마트 이름', '년도', '월', '일', '가격(원)', '비고']
        if not all(col in df.columns for col in expected_columns):
            missing = [col for col in expected_columns if col not in df.columns]
            print(f"오류: 다음 열이 없습니다: {missing}")
            return

        # 새 JSON 구조: { "상품이름": { "마트이름": [ {date, price, note}, ... ] } }
        grouped_data = defaultdict(lambda: defaultdict(list))

        for index, row in df.iterrows():
            try:
                item = str(row['상품 이름']).strip()
                mart = str(row['마트 이름']).strip()
                year, month, day = int(row['년도']), int(row['월']), int(row['일'])
                price = int(row['가격(원)'])
                note = str(row.get('비고', '')).strip()
                date_str = f"{year:04d}. {month:02d}. {day:02d}."

                grouped_data[item][mart].append({
                    "date": date_str,
                    "price": price,
                    "note": note if note else None
                })

            except Exception as e:
                print(f"⚠️ {index+2}번째 행 처리 중 오류: {e}")

        # 날짜 기준 내림차순 정렬
        from datetime import datetime
        for item, marts in grouped_data.items():
            for mart, entries in marts.items():
                entries.sort(key=lambda x: datetime.strptime(x['date'], '%Y. %m. %d.'), reverse=True)

        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(grouped_data, f, ensure_ascii=False, indent=4)

        print(f"✅ 변환 완료: '{json_path}'로 저장되었습니다.")

    except FileNotFoundError:
        print(f"❌ 파일을 찾을 수 없습니다: {excel_path}")
    except Exception as e:
        print(f"❌ 예외 발생: {e}")

# 실행
if __name__ == "__main__":
    convert_excel_to_json_grouped(EXCEL_FILE_PATH, JSON_OUTPUT_PATH, SHEET_NAME)
