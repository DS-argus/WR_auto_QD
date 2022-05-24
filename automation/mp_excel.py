import pandas as pd
import xlwings as xw
from datetime import datetime
import warnings


warnings.filterwarnings('ignore')


def main():
    # DB에서 편입정보 sheet, database sheet을 df1, df2에 저장
    with xw.App(visible=False) as app:
        db = xw.Book(r'\\172.31.1.222\Deriva\자동화\DB\변액 DATABASE.xlsm')
        df1 = db.sheets("편입정보").range("A1").options(pd.DataFrame, index=False, expand='table', header=False).value
        df2 = db.sheets("database").range("A1").options(pd.DataFrame, index=False, expand='table', header=False).value
        db.close()

    # ID를 index로
    df1 = df1.set_index(df1[0])
    df1 = df1.rename(columns=df1.iloc[0])
    df1 = df1.drop(df1.columns[0], axis=1)
    df1 = df1.drop(df1.index[0])
    df1.index.name = 'ELS ID'

    df2 = df2.set_index(df2[0])
    df2 = df2.rename(columns=df2.iloc[0])
    df2 = df2.drop(df2.columns[0], axis=1)
    df2 = df2.drop(df2.index[0])
    df2.index.name = 'ID'

    # 상환된 투자내역 행 제외
    mask = df1['진행상태'].isin(['투자 중'])
    df1 = df1[mask]

    # 펀드정보_xlwings 시트 기존 내용 삭제
    file_schedule = xw.Book.caller()
    file_schedule.sheets['펀드정보_xlwings'].range("A2:E1000").clear_contents()

    # 월지급리스트 시트에 있는 평가일 받아오기
    search_date = file_schedule.sheets['월지급리스트'].range("A1").value
    search_date = datetime.combine(search_date, datetime.min.time())

    # 원하는 칼럼을 가진 빈 데이터프레임 생성
    df_fundlist_all = pd.DataFrame(columns=['펀드코드', '펀드명', '차수', '편입일',  '쿠폰'])
    df_fundlist_all.index.name = 'ID'

    # 만약 평가일에 해당하는 ELS 편입건 있으면 해당 ELS 편입한 펀드 찾아서 데이터프레임에 저장
    for i in range(len(df1.index)):
        for j in range(len(df1.iloc[:, 28:88].columns)):
            if df1.iloc[i, 28+j] == search_date:
                df_fundlist = pd.DataFrame(df2.iloc[:, 0][df2.iloc[:, 2] == df1.index[i]])
                df_fundlist['펀드코드'] = df2.iloc[:, 7][df2.iloc[:, 2] == df1.index[i]]
                df_fundlist['차수'] = df1.columns[28+j]
                df_fundlist['편입일'] = df1.iloc[i, 2]
                df_fundlist['쿠폰'] = df1.iloc[i, 15]
                df_fundlist_all = pd.concat([df_fundlist_all, df_fundlist])

    # 해당 기간에 펀드 없을경우 '정보 없음' 출력하고 종료
    if df_fundlist_all.empty:
        data = ["", '해당 펀드 없음', "", "", ""]
        df_fundlist_all.loc[0] = data
        file_schedule.sheets['펀드정보_xlwings'][0, 0].options(index=False).value = df_fundlist_all
        return

    # 펀드가 있을 경우 펀드명으로 정렬해서 ID부분제외하고 엑셀에 출력
    df_fundlist_all["ID"] = df_fundlist_all.index
    df_fundlist_all = df_fundlist_all.sort_values(by=['펀드명'])
    df_fundlist_all = df_fundlist_all[['펀드코드', '펀드명', '차수', '편입일', '쿠폰']]

    file_schedule.sheets['펀드정보_xlwings'][0, 0].options(index=False).value = df_fundlist_all

    # 펀드코드하고 편입일 일치하는 ELS종목 액면_FAS에서 모두 찾아서 월지급리스트 시트에 출력
    df_info = df_fundlist_all

    df_notional = file_schedule.sheets['액면_FAS'].range("A1").options(pd.DataFrame,
                                                                     index=False,
                                                                     expand='table',
                                                                     header=True).value

    df_notional.drop(['조회일자'], axis=1, inplace=True)

    df_result = pd.DataFrame(columns=['펀드코드', '펀드명', '종목코드', '종목명', '주수/계약수/액면', '발행일', '회차', '쿠폰'])

    for i in range(len(df_info)):
        code = df_info.iloc[i, 0]
        issue_date = df_info.iloc[i, 3]
        num = df_info.iloc[i, 2]
        coupon = df_info.iloc[i, 4]

        match_els = df_notional[(df_notional['펀드코드'] == code) & (df_notional['발행일'] == issue_date)]

        match_els['쿠폰'] = coupon
        match_els['회차'] = num[:-1]

        df_result = pd.concat([df_result, match_els])

    df_result.rename(columns={'주수/계약수/액면': '계약금액'}, inplace=True)
    df_result['MP베리어금액'] = df_result['계약금액'] * df_result['쿠폰'] / 12
    df_result['평가구분'] = "MP"
    df_result.drop(['쿠폰', '발행일'], axis=1, inplace=True)

    df_result = df_result[['펀드코드', '펀드명', '종목코드', '종목명', '평가구분', '회차', '계약금액', 'MP베리어금액']]

    # 달러형 아니면 회차가 6배수면 MP 추가, 달러형이면 회차가 6의 배수면 중도상환으로 변경
    df_add = pd.DataFrame(columns=['펀드코드', '펀드명', '종목코드', '종목명', '평가구분', '회차', '계약금액', 'MP베리어금액'])

    for i in range(len(df_result)):
        if int(df_result.iloc[i, 5]) % 6 == 0:   # 6의 배수면
            k = len(df_add)
            df_add.loc[k+1] = pd.Series(df_result.iloc[i, :])

    # df_add가 비어있지 않으면, 즉 6의 배수가 있으면
    if not df_add.empty:
        # 6의 배수인 종목 따로 모아서 중도상환으로 바꿔줌
        df_add['평가구분'] = '중도상환'

        # 기존에 달러형 & 6의 배수인 것들은 MP로 나와있으니 제거
        df_result = df_result[~((df_result['펀드명'].str.contains("달러")) & (df_result['회차'].astype(int) % 6 == 0))]
        df_result = pd.concat([df_result, df_add])

    # 달러형은 MP 배리어 금액에 회차만큼 곱해줘야하고 회차를 6으로 나눠줘야함 --> MP가 아니니까 남들 6차가 1차
    df_result["회차"] = df_result["회차"].astype(float)
    df_result['MP베리어금액'] = df_result['MP베리어금액'].astype(float)

    for i in range(len(df_result)):
        if df_result.iloc[i, 1][:2] == "달러":
            df_result.iloc[i, 7] = f'{df_result.iloc[i, 7] * df_result.iloc[i, 5]:.2f}'
            df_result.iloc[i, 5] = df_result.iloc[i, 5] // 6

    # 정렬
    df_result.sort_values(by=['종목명', '평가구분', '펀드명'], inplace=True)

    file_schedule.sheets['월지급리스트'].range("A2:H1000").clear_contents()

    file_schedule.sheets['월지급리스트'][1, 0].options(index=False).value = df_result

    return


if __name__ == "__main__":
    xw.Book(r"\\172.31.1.222\Deriva\자동화\자동화폴더\월지급내역서\월지급내역서.xlsm").set_mock_caller()
    main()
