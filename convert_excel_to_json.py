import pandas as pd
import json
from collections import defaultdict

# --- 설정 부분 ---
# 엑셀 파일 경로를 정확하게 입력해 주세요. (예: './우리집_마트_가격데이터.xlsx')
EXCEL_FILE_PATH = 'mart_price.xlsx'
# 생성될 JSON 파일 이름을 설정해 주세요. (예: './data.json')
JSON_OUTPUT_PATH = 'data.json'
# 엑셀 시트 이름 (기본값은 첫 번째 시트)
SHEET_NAME = 0 # 첫 번째 시트는 0으로 설정

# --- 스크립트 본문 ---
def convert_excel_to_json(excel_path, json_path, sheet_name=0):
    """
    엑셀 파일에서 가격 데이터를 읽어와 JSON 형식으로 변환합니다.
    JSON 형식: { "상품이름-마트이름": [ {date: "...", price: ..., note: "..."}, ... ], ... }
    """
    try:
        # 엑셀 파일 읽기
        # header=0 은 첫 번째 행을 열 이름으로 사용한다는 의미
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=0)

        # 필요한 열 이름 확인 (엑셀 파일의 실제 열 이름과 일치해야 합니다!)
        # 만약 엑셀의 열 이름이 다르다면 이 부분을 수정해주세요.
        expected_columns = ['상품 이름', '마트 이름', '년도', '월', '일', '가격(원)', '비고']
        if not all(col in df.columns for col in expected_columns):
            missing = [col for col in expected_columns if col not in df.columns]
            print(f"오류: 엑셀 파일에 필요한 열이 없습니다. 다음 열을 확인해주세요: {missing}")
            print(f"엑셀 파일의 실제 열 이름: {list(df.columns)}")
            return

        # 데이터를 저장할 딕셔너리 초기화
        # defaultdict를 사용하면 키가 없을 때 자동으로 빈 리스트를 생성해 줍니다.
        items_data = defaultdict(list)

        # 엑셀 데이터 순회하며 JSON 구조에 맞게 데이터 가공
        for index, row in df.iterrows():
            try:
                item_name = str(row['상품 이름']).strip()
                store_name = str(row['마트 이름']).strip()
                year = int(row['년도'])
                month = int(row['월'])
                day = int(row['일'])
                price = int(row['가격(원)'])
                note = str(row.get('비고', '')).strip() # '비고' 열이 없을 수도 있으니 get() 사용

                # 상품의 고유 키 생성
                item_key = f"{item_name}-{store_name}"

                # 날짜 형식 맞추기 (예: "2025. 4. 30.")
                # 월과 일이 한 자릿수일 때 앞에 0을 붙일지 여부는 필요에 따라 조정
                date_str = f"{year}. {month}. {day}."

                # 해당 기록을 나타내는 딕셔너리 생성
                entry = {
                    "date": date_str,
                    "price": price,
                    "note": note if note else None # 비고가 빈 문자열이면 None으로 저장
                }

                # 해당 상품 키의 리스트에 기록 추가
                items_data[item_key].append(entry)

            except ValueError as ve:
                print(f"경고: {index+2}번째 행 데이터 변환 오류 - {ve}. 이 행은 건너뜁니다.")
            except KeyError as ke:
                 print(f"경고: {index+2}번째 행에서 예상치 못한 열 오류 - {ke}. 엑셀 열 이름을 다시 확인해주세요.")
            except Exception as e:
                 print(f"경고: {index+2}번째 행 처리 중 알 수 없는 오류 발생 - {e}. 이 행은 건너뜁니다.")


        # 각 상품별 기록을 날짜 순서대로 정렬 (년 -> 월 -> 일 순으로)
        # 'date' 문자열을 파싱해서 비교하는 것이 더 정확하지만,
        # 현재 형식 "YYYY. M. D."은 문자열 정렬로도 대부분 맞습니다.
        # 더 정확한 정렬이 필요하면 datetime 객체로 변환 후 정렬해야 합니다.
        # 예: sorted_entries = sorted(entries, key=lambda x: datetime.strptime(x['date'], '%Y. %m. %d.'))
        # 여기서는 간단하게 문자열 정렬을 사용합니다.
        # 엑셀에서 이미 날짜순으로 입력했다면 이 정렬 과정은 필요 없을 수도 있습니다.
        for key in items_data:
            # 'date' 문자열을 기준으로 오름차순 정렬
            # 예: "2025. 4. 30." < "2025. 5. 1."
            items_data[key].sort(key=lambda x: x['date']) # 간단한 문자열 정렬

        # JSON 파일로 저장
        with open(json_path, 'w', encoding='utf-8') as f:
            # ensure_ascii=False: 한글이 깨지지 않도록 설정
            # indent=4: JSON 파일을 읽기 좋게 들여쓰기 설정
            json.dump(items_data, f, ensure_ascii=False, indent=4)

        print(f"'{excel_path}' 파일이 성공적으로 '{json_path}'로 변환되었습니다.")

    except FileNotFoundError:
        print(f"오류: 엑셀 파일을 찾을 수 없습니다. 경로를 확인해주세요: {excel_path}")
    except Exception as e:
        print(f"파일 처리 중 오류 발생: {e}")

# 스크립트 실행 부분
if __name__ == "__main__":
    convert_excel_to_json(EXCEL_FILE_PATH, JSON_OUTPUT_PATH, SHEET_NAME)
