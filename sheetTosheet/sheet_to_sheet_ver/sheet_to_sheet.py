import openpyxl
import pandas as pd
import os

loop_n = 0
while loop_n < 100:
    # temp_file 폴더에 있는 파일 가져오기
    temp_path = './temp_file/'
    temp_file = os.listdir(temp_path)

    # temp 파일
    n = 100
    print('파일을 선택해주세요.')
    while n > len(temp_file):
        for i, data in enumerate(temp_file, 1):
            print('{}.{}'.format(i, data))
        try:
            n = int(input())
        except:
            print('숫자만 입력해주세요.')
            
        if n <= len(temp_file):
            continue
        else:
            print('숫자가 너무 큽니다.')
    print('\n')
    df = pd.read_excel(temp_path + temp_file[n-1], header=1)

    # data_file 폴더에 있는 파일 가져오기
    data_path = './data_file/'
    data_file = os.listdir(data_path)

    ### 데이터 파일
    n2 = 100
    print('파일을 선택해주세요.')
    while n2 > len(data_file):
        for i, data in enumerate(data_file, 1):
            print('{}.{}'.format(i, data))
        try:
            n2 = int(input())
        except:
            print('숫자만 입력해주세요.')
            
        if n2 <= len(data_file):
            break
        else:
            print('숫자가 너무 큽니다.')
        
    df2 = pd.read_csv(data_path + data_file[n2-1], encoding='cp949')

    for i in range(0, len(df)):
        if '제' in df.loc[i, '행정읍면동명']:
            df.loc[i, '행정읍면동명'] = df.loc[i, '행정읍면동명'].replace('제','')

    # 데이터 파일에 있는 지역명 '.', '제' 글자 제거
    for i in range(0, len(df2)):
        if '.' in df2.loc[i, '지역명']:
            df2.loc[i, '지역명'] = df2.loc[i, '지역명'].replace('.','')
        elif '제' in df2.loc[i, '지역명']:
            df2.loc[i, '지역명'] = df2.loc[i, '지역명'].replace('제','')

    for i in range(0, len(df)):
        for j in range(0, len(df2)):
            try:
                # 읍/면/동이 같으면 지역명 변경
                if df.loc[i, '행정읍면동명'].split()[1] == df2.loc[j, '지역명'].split()[3]:
                    df2.loc[j, '지역명'] = df.loc[i, '행정읍면동명']

                # 읍/면/동 앞 3글자만 같으면 지역명 변경 ex 용암1, 강서1
                elif df.loc[i, '행정읍면동명'].split()[1][:3] == df2.loc[j, '지역명'].split()[3][:3]:
                    df2.loc[j, '지역명'] = df.loc[i, '행정읍면동명']

                # 에러 발생 시 지역명 그대로
            except:
                df2.loc[j, '지역명'] = df2.loc[j, '지역명']

    for i in range(0, len(df)):
        for j in range(0, len(df2)):
            if df.iloc[i][0] in df2.iloc[j][0]:
                df.loc[i, '성별'] = df2.loc[j, '성별']
                df.loc[i, '연령'] = df2.loc[j, '연령']
                df.loc[i, '비표준화지표분자'] = df2.loc[j, '비표준화지표분자']
                df.loc[i, '비표준화지표분모'] = float(df2.loc[j, '비표준화지표분모'].replace(',',''))
                df.loc[i, '비표준화지표'] = df2.loc[j, '비표준화지표지표값']
                df.loc[i, '표준화지표'] = df2.loc[j, '표준화지표지표값']

    # 행정읍면동명, 빈 컬럼 제거
    df = df.drop(['행정읍면동명', 'Unnamed: 8'], axis=1)

    # 엑셀 파일 열기
    wb = openpyxl.load_workbook(temp_path + temp_file[n-1])

    # 워크시트 선택
    ws = wb['행정읍면동별 주제도']

    # 셀마다 df 값 입력 (B3부터 차례대로 값 입력)
    for i in range(len(df)):
        for j in range(len(df.loc[0])):
            ws.cell(row=3+i, column=j+2).value = df.loc[i].values[j]

    wb.save(temp_path + temp_file[n-1])
    wb.close()
    
    print('계속하시려면 1, 끝내시려면 2 를 입력해주세요.')
    if int(input()) == 1:
        loop_n = 0
    else:
        break