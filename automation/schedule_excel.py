import pandas as pd
import numpy as np
import xlwings as xw
from datetime import datetime, timedelta


def main():
    # DB에서 편입정보 sheet, database sheet을 df1, df2에 저장
    app = xw.App()
    app.visible = False
    db = xw.Book(r"\\172.31.1.222\Deriva\자동화\DB\변액 DATABASE.xlsm")
    df1 = db.sheets("편입정보").range("A1").options(pd.DataFrame, index=False, expand='table', header=False).value
    df2 = db.sheets("database").range("A1").options(pd.DataFrame, index=False, expand='table', header=False).value
    db.close()
    app.quit()

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

    file_schedule = xw.Book.caller()
    file_schedule.sheets['Schedule'].range("A7:Q1000").clear_contents()

    # 검색기간 입력 받기
    start_date = file_schedule.sheets['Schedule'].range("K2").value
    start_date = datetime.combine(start_date, datetime.min.time())
    end_date = file_schedule.sheets['Schedule'].range("K3").value
    end_date = datetime.combine(end_date, datetime.min.time())
    btw_date = end_date - start_date

    df_fundlist_all = pd.DataFrame(columns=['펀드코드', '펀드명', '차수', '편입일',
                                            '발행사1', '발행사2', '발행사3', '발행사4',
                                            '쿠폰', '일자', 'Size', 'Worst', 'Level of Worst', '구조'])
    df_fundlist_all.index.name = 'ID'

    for k in range(btw_date.days+1):
        search_date = start_date + timedelta(days=k)

        for i in range(len(df1.index)):

            for j in range(len(df1.iloc[:, 28:88].columns)):

                if df1.iloc[i, 28+j] == search_date:
                    df_fundlist = pd.DataFrame(df2.iloc[:, 0][df2.iloc[:, 2] == df1.index[i]])
                    df_fundlist['펀드코드'] = df2.iloc[:, 7][df2.iloc[:, 2] == df1.index[i]]
                    df_fundlist['차수'] = df1.columns[28+j]
                    df_fundlist['편입일'] = df1.iloc[i, 2]
                    df_fundlist['발행사1'] = str(df1.iloc[i, 3]) + " " + str(int(df1.iloc[i, 4] or 0))
                    df_fundlist['발행사2'] = str(df1.iloc[i, 5]) + " " + str(int(df1.iloc[i, 6] or 0))
                    df_fundlist['발행사3'] = str(df1.iloc[i, 7]) + " " + str(int(df1.iloc[i, 8] or 0))
                    df_fundlist['발행사4'] = str(df1.iloc[i, 9]) + " " + str(int(df1.iloc[i, 10] or 0))
                    df_fundlist['쿠폰'] = df1.iloc[i, 15]
                    df_fundlist['일자'] = search_date.date()
                    df_fundlist['Size'] = df2.iloc[:, 8][df2.iloc[:, 2] == df1.index[i]]
                    df_fundlist['Worst'] = df2.iloc[:, 38][df2.iloc[:, 2] == df1.index[i]]
                    df_fundlist['Level of Worst'] = df2.iloc[:, 37][df2.iloc[:, 2] == df1.index[i]]
                    df_fundlist['구조'] = df1.iloc[i, 16]
                    df_fundlist_all = pd.concat([df_fundlist_all, df_fundlist])
                    break

    # 해당 기간에 펀드 없을경우 '정보 없음' 출력
    if df_fundlist_all.empty:
        data = ['해당 펀드 없음', "", "", "", "", "", "", "", "", "", "", ""]
        df_fundlist_all.loc[0] = data
        df_fundlist_all["ID"] = df_fundlist_all.index
        df_fundlist_all = df_fundlist_all.sort_index().sort_values(by=['일자', '발행사1', '펀드명'])
        df_fundlist_all.set_index(df_fundlist_all["일자"], inplace=True)
        del df_fundlist_all['일자']
        df_fundlist_all = df_fundlist_all[['ID', '펀드코드', '펀드명', '차수',
                                           '편입일', '발행사1', '발행사2', '발행사3', '발행사4',
                                           '쿠폰', 'Size', 'Worst', 'Level of Worst', '구조']]
        file_schedule.sheets['Schedule'][5, 0].options(index=True).value = df_fundlist_all
        return

    # 펀드가 있을 경우
    df_fundlist_all["ID"] = df_fundlist_all.index
    df_fundlist_all = df_fundlist_all.sort_index().sort_values(by=['일자', '발행사1', '펀드명'])
    df_fundlist_all.set_index(df_fundlist_all["일자"], inplace=True)
    del df_fundlist_all['일자']
    df_fundlist_all = df_fundlist_all[['ID', '펀드코드', '펀드명', '차수',
                                       '편입일', '발행사1', '발행사2', '발행사3', '발행사4',
                                       '쿠폰', 'Size', 'Worst', 'Level of Worst', '구조']]
    df_fundlist_all = df_fundlist_all.replace("None 0", "")

    file_schedule.sheets['Schedule'][5, 0].options(index=True).value = df_fundlist_all
    file_schedule.sheets['Schedule'].range("K7:K1000").number_format = '0.000%'
    file_schedule.sheets['Schedule'].range("L7:L1000").number_format = '#,##0'
    file_schedule.sheets['Schedule'].range("N7:N1000").number_format = '0.00%'

    date_list = np.array(list(set(df_fundlist_all.index.values)))
    date_list = np.sort(date_list)
    date_list = date_list.reshape(len(date_list), 1)
    file_schedule.sheets['Schedule'].range("Q7").value = date_list

    return


if __name__ == "__main__":
    xw.Book(r"\\172.31.1.222\Deriva\자동화\자동화폴더\이벤트 내역서\변액 내역서 자동화.xlsm").set_mock_caller()
    main()
