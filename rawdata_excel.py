from idxdata.historical_data import get_hist_data_from_sql

import xlwings as xw
import pandas as pd

from datetime import date, timedelta


def main():

    def get_data() -> pd.DataFrame:
        start = date(2018, 1, 1)
        end = date.today() - timedelta(days=1)

        underlyings = [
            'KOSPI200', 'HSCEI', 'HSI', 'NIKKEI225', 'S&P500', 'EUROSTOXX50', 'CSI300',
            'S&P500(Q)', 'EUROSTOXX50(Q)',
            'S&P500(KRW)', 'EUROSTOXX50(KRW)', 'HSCEI(KRW)'
        ]

        df = get_hist_data_from_sql(start, end, underlyings, type='w', ffill=True)

        return df

    def update_excel(dataframe: pd.DataFrame):
        with xw.App(visible=False)as app:
            ex = xw.Book(r'\\172.31.1.222\Deriva\자동화\DB\DB 종가데이터 엑셀\DB종가.xlsx')
            sh1 = ex.sheets['DB종가']

            sh1.range("A1:M8000").clear_contents()

            sh1[0, 0].options(index=True).value = dataframe

            ex.save()
            ex.close()

    df = get_data()
    update_excel(df)

    return


if __name__ == "__main__":
    main()
