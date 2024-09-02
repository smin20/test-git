import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import pandas as pd
import time
import threading
import os
import sys
import requests
import datetime
import numpy as np

# 기본 창 생성
root = tk.Tk()
root.title("네이버 부동산 크롤링")
root.option_add("*Font","맑은고딕 12")

# 로그 출력을 위한 Text 위젯 추가
log_text = tk.Text(root, height=20, state='disabled')
log_text.grid(row=12, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')

# print 함수를 로그로 리디렉션하는 클래스
class Redirector:
    def __init__(self, widget):
        self.widget = widget

    def write(self, message):
        self.widget.config(state='normal')
        self.widget.insert(tk.END, message)
        self.widget.yview(tk.END)
        self.widget.config(state='disabled')

    def flush(self):
        pass

sys.stdout = Redirector(log_text)
sys.stderr = Redirector(log_text)

# 지역 선택
df = pd.read_csv('cortarNo.csv',encoding= 'cp949')
# 시/도 순서 지정
sido_order = ['서울시', '경기도', '인천시', '부산시', '대전시', '대구시', '울산시', '세종시', '광주시', '강원도', '충청북도', '충청남도', '경상북도', '경상남도', '전북도', '전라남도', '제주도']
sido = [sido for sido in sido_order if sido in df['시도'].unique().tolist()]

# 레이블, 콤보박스 및 버튼을 그리드에 배치
tk.Label(root, text="시/도 선택:").grid(row=0, column=0, padx=10, pady=5, sticky='w')
combo1 = ttk.Combobox(root, values=sido)
combo1.set("Select an option")
combo1.grid(row=0, column=1, padx=0, pady=10)

# 도시 선택
tk.Label(root, text="시/군/구 선택:").grid(row=1, column=0, padx=10, pady=5, sticky='w')
combo2 = ttk.Combobox(root)
combo2.set("Select a choice")
combo2.grid(row=1, column=1, padx=10, pady=10)

# 항목 선택
tk.Label(root, text="읍/면/동 선택:").grid(row=2, column=0, padx=10, pady=5, sticky='w')
combo3 = ttk.Combobox(root)
combo3.set("Select an item")
combo3.grid(row=2, column=1, padx=10, pady=10)

# 항목 선택
tk.Label(root, text="매물수").grid(row=2, column=2, padx=0, pady=5, sticky='w')
combo4 = ttk.Combobox(root, values=[20, 500, 1000, 2000, 9999])
combo4.set("500")
combo4.grid(row=2, column=2, padx=40, pady=10)

# 체크버튼 추가
sale_var = tk.BooleanVar()
rent_var = tk.BooleanVar()
lease_var = tk.BooleanVar()

tk.Checkbutton(root, text="매매", variable=sale_var).grid(row=4, column=0, padx=10, pady=5, sticky='w')
tk.Checkbutton(root, text="월세", variable=rent_var).grid(row=4, column=1, padx=10, pady=5, sticky='w')
tk.Checkbutton(root, text="전세", variable=lease_var).grid(row=4, column=2, padx=10, pady=5, sticky='w')

# 부동산 유형 체크버튼
rlet_types = [
    {'tagCd': 'APT', 'uiTagNm': '아파트'},
    {'tagCd': 'OPST', 'uiTagNm': '오피스텔'},
    {'tagCd': 'VL', 'uiTagNm': '빌라'},
    {'tagCd': 'ABYG', 'uiTagNm': '아파트분양권'},
    {'tagCd': 'OBYG', 'uiTagNm': '오피스텔분양권'},
    {'tagCd': 'JGC', 'uiTagNm': '재건축'},
    {'tagCd': 'JWJT', 'uiTagNm': '전원주택'},
    {'tagCd': 'DDDGG', 'uiTagNm': '단독/다가구'},
    {'tagCd': 'SGJT', 'uiTagNm': '상가주택'},
    {'tagCd': 'HOJT', 'uiTagNm': '한옥주택'},
    {'tagCd': 'JGB', 'uiTagNm': '재개발'},
    {'tagCd': 'OR', 'uiTagNm': '원룸'},
    {'tagCd': 'GSW', 'uiTagNm': '고시원'},
    {'tagCd': 'SG', 'uiTagNm': '상가'},
    {'tagCd': 'SMS', 'uiTagNm': '사무실'},
    {'tagCd': 'GJCG', 'uiTagNm': '공장/창고'},
    {'tagCd': 'GM', 'uiTagNm': '건물'},
    {'tagCd': 'TJ', 'uiTagNm': '토지'},
    {'tagCd': 'APTHGJ', 'uiTagNm': '지식산업센터'}
]

rlet_vars = {}
for idx, rlet_type in enumerate(rlet_types):
    var = tk.BooleanVar()
    rlet_vars[rlet_type['tagCd']] = var
    tk.Checkbutton(root, text=rlet_type['uiTagNm'], variable=var).grid(row=5+idx//3, column=idx%3, padx=0, pady=2, sticky='w')


def get_rlet_types():
    rlet_types = []
    for tagCd, var in rlet_vars.items():
        if var.get():
            rlet_types.append(tagCd)
    return rlet_types

# 이미지에서 추가된 버튼들
btn_search = tk.Button(root, text="조회", command=lambda: threading.Thread(target=btn_search_click).start())
btn_search.grid(row=0, column=2, padx=10, pady=10)

btn_open_file = tk.Button(root, text="파일열기", command = lambda: open_file())
btn_open_file.grid(row=3, column=0, padx=10, pady=10, sticky='s')

btn_open_file_location = tk.Button(root, text="파일위치열기", command = lambda: open_folder())
btn_open_file_location.grid(row=3, column=1, padx=10, pady=10, sticky='s')

version_label = tk.Label(root, text="made by jsm02115")
version_label.grid(row=3, column=2, padx=10, pady=10)

reset_label = tk.Button(root, text="초기화",  command = lambda: reset())
reset_label.grid(row=1, column=2, padx=10, pady=10)

def reset():
    combo1.set("Select an option")
    combo2.set("Select a choice")
    combo3.set("Select an item")
    combo4.set("500")
    sale_var.set(True)
    rent_var.set(True)
    lease_var.set(True)
    for tagCd, var in rlet_vars.items():
        var.set(tagCd == 'SG')
    print('시/도를 다시 선택하세요.')

def get_file_name():
    current_time = datetime.datetime.now().strftime("%Y%m%d")
    return f'{combo1.get()}_{combo2.get()}_{combo3.get()}_{current_time}.xlsx'

def get_cortarno():
    # Combine the values from combo1, combo2, and combo3 to form the name
    combined_name = f'{combo1.get()} {combo2.get()} {combo3.get()}'
    # Read the cortarNo.csv file
    cortarNo_df = pd.read_csv('cortarNo.csv', encoding='cp949')
    # Search for the code corresponding to the combined name
    try:
        cortar_no = cortarNo_df.loc[cortarNo_df['name'] == combined_name, 'code'].values[0]
        return cortar_no
    except IndexError:
        print(f"Error: No matching code found for {combined_name}")
        return None

def get_pages():
    pages = combo4.get()
    return(int(pages)/20)

def open_folder():
    current_time = datetime.datetime.now().strftime("%Y%m%d")
    file_name = f'{combo1.get()}_{combo2.get()}_{combo3.get()}_{current_time}.xlsx'
    file_path = os.path.join(os.getcwd(), file_name)
    os.system(f'explorer /select,"{file_path}"') 

def open_file():
    file_name = get_file_name()
    print(f"Opening file: {file_name}")
    if os.path.exists(file_name):
        os.system(f'start \"\" \"{file_name}\"')
    else:
        print("File does not exist")

# 콤보박스 선택 이벤트 핸들러
def on_combo_select(event):
    selected_combo = event.widget
    selected_value = selected_combo.get()
    
    if selected_combo == combo1:
        threading.Thread(target=update_combo2, args=(selected_value,)).start()

    if selected_combo == combo2:
        threading.Thread(target=update_combo3, args=(selected_value,)).start()

    if selected_combo == combo3:
        print('수정문의 : jsm02115@naver.com')

def get_gu(sido):
    return sorted(df[df['시도'] == sido]['구'].unique().tolist())

def get_dong(sido, gu):
    return sorted(df[(df['시도'] == sido) & (df['구'] == gu)]['동'].unique().tolist())

# 첫 번째 콤보박스의 선택에 따라 두 번째 콤보박스를 업데이트하는 함수
def update_combo2(selected_value):
    gu_list = get_gu(selected_value)
    combo2['values'] = gu_list
    combo2.set("Select a choice")

# 두 번째 콤보박스의 선택에 따라 세 번째 콤보박스를 업데이트하는 함수
def update_combo3(selected_value):
    sido_value = combo1.get()
    dong_list = get_dong(sido_value, selected_value)
    combo3['values'] = dong_list
    combo3.set("Select a choice")

def get_trade_types():
    trade_types = []
    if sale_var.get():
        trade_types.append('A1')
    if rent_var.get():
        trade_types.append('B2')
    if lease_var.get():
        trade_types.append('B1')
    return trade_types

# btn_search 클릭 시 실행되는 함수
def btn_search_click():
    print('btn click. 조회 시작.')
    # 전체 데이터를 저장할 DataFrame 초기화
    all_data = pd.DataFrame()
    cortarNo = get_cortarno()
    pages = int(get_pages())
    trade_types = get_trade_types()
    trade_type_param = '%3A'.join(trade_types)
    rlet_types = get_rlet_types()
    rlet_type_param = ':%3A'.join(rlet_types)

    if not trade_type_param:
        print("Please select at least one trade type.")
        return
    
    for idx in range(1, pages + 1):
        try:
            print(f"{idx} / {pages}page 조회 중")
            # 요청 URL
            url = f"https://m.land.naver.com/cluster/ajax/articleList?itemId=&mapKey=&lgeo=&showR0=&rletTpCd={rlet_type_param}&tradTpCd={trade_type_param}&z=14&cortarNo={cortarNo}&sort=dates&page={idx}"
            # 요청 헤더
            headers = {
                'Accept': 'application/json, text/javascript, */*; q=0.01',
                'Accept-Encoding': 'gzip, deflate, br, zstd',
                'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
                'Referer': 'https://m.land.naver.com/',
                'Sec-Ch-Ua': '"Google Chrome";v="125", "Chromium";v="125", "Not.A/Brand";v="24"',
                'Sec-Ch-Ua-Mobile': '?0',
                'Sec-Ch-Ua-Platform': '"Windows"',
                'Sec-Fetch-Dest': 'empty',
                'Sec-Fetch-Mode': 'cors',
                'Sec-Fetch-Site': 'same-origin',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36',
                'X-Requested-With': 'XMLHttpRequest',
                'Priority': 'u=1, i'
            }
            # GET 요청 보내기
            res = requests.get(url, headers=headers)
            res.raise_for_status()

            # JSON 데이터 파싱
            data = res.json()
            real_data = data.get('body', [])  # 실제 데이터 리스트

            if not real_data:
                print(f"No data found on page {idx}. Stopping search.")
                break

            # DataFrame으로 변환하여 기존 DataFrame에 추가
            print(real_data[0]['atclNo'])
            df = pd.DataFrame(real_data)
            all_data = pd.concat([all_data, df], ignore_index=True)
            time.sleep(0.5)

        except Exception as e:
            print(f"Error on page {idx}: {e}")
            break
    
    # 필요한 칼럼만 선택하고 한글로 이름 변경하기 위한 매핑 딕셔너리
    column_map = {
        'atclCfmYmd': '날짜',
        'atclNm' : '구분',
        'tradTpNm': '거래구분',
        'flrInfo': '층수',
        'prc': '매가(보증금)',
        'rentPrc': '월세',
        'spc1': '계약면적',
        'spc2': '전용면적',
        'direction': '향',
        'atclFetrDesc': '코멘트',
        'tagList': '태그',
        'cpNm': '제공',
        'rltrNm': '공인중개사',
        'atclNo' : '매물번호'
    }

    # 데이터프레임에서 필요한 칼럼만 선택하여 새 데이터프레임 생성
    selected_columns = list(column_map.keys())
    df_selected = all_data[selected_columns]

    # 칼럼 이름을 한글로 변경
    df_selected = df_selected.rename(columns=column_map)

    # 링크 열의 값을 URL 형식으로 변경
    base_url = "https://fin.land.naver.com/articles/"
    df_selected['Link'] = '=HYPERLINK("' + df_selected['매물번호'].apply(lambda x: f"{base_url}{x}") + '", "LINK")'

    base_url2 = "https://m.land.naver.com/near/article/"
    df_selected['위치'] = '=HYPERLINK("' + df_selected['매물번호'].apply(lambda x: f"{base_url2}{x}") + '", "위치")'

    # '전용면적'을 정수형으로 변환
    df_selected['전용면적'] = df_selected['전용면적'].astype(float)

    # 새로운 칼럼 '평당가'를 조건에 따라 계산하여 추가
    df_selected['평당가'] = np.where(df_selected['거래구분'] == '월세',
                                df_selected['월세'] / df_selected['전용면적'],
                                df_selected['매가(보증금)'] / df_selected['전용면적'])

    # '평당가'를 정수형으로 변환
    df_selected['평당가'] = df_selected['평당가'].astype(int)

    # 모든 DataFrame을 하나로 병합
    file_name = get_file_name()

    # 평당매가 계산

    # 수익률 계산
    df_selected.sort_values('날짜',ascending=False)
    df_selected.to_excel(excel_writer=file_name, index=False)
    print(f'Data saved to {file_name}')
    
# 콤보박스 선택 이벤트를 이벤트 핸들러에 바인딩
combo1.bind("<<ComboboxSelected>>", on_combo_select)
combo2.bind("<<ComboboxSelected>>", on_combo_select)
combo3.bind("<<ComboboxSelected>>", on_combo_select)

# Tkinter 애플리케이션 실행
reset()
root.mainloop()

