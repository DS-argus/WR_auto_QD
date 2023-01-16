from idxdata.historical_data import get_price_from_sql

import numpy as np
import pandas as pd
import xlwings as xw
import matplotlib.pyplot as plt

from dateutil.relativedelta import relativedelta
from datetime import date

def price_plot(df):

    # plot 크기 및 subplot 개수, 간격
    f, axes = plt.subplots(6, 1)
    f.set_size_inches((10, 16))
    plt.subplots_adjust(hspace=0.8)

    # 그래프 x축 범위 설정용
    start_date = df.index[0] - relativedelta(years=1)
    end_date = df.index[-1] + relativedelta(years=1)

    # 그래프 그리기
    axes[0].plot(df["KOSPI200"])
    axes[1].plot(df["HSCEI"])
    axes[2].plot(df["NIKKEI225"])
    axes[3].plot(df["S&P500"])
    axes[4].plot(df["EUROSTOXX50"])
    axes[5].plot(df["CSI300"])

    # 축범위 설정
    axes[0].set_xlim([start_date, end_date])
    axes[1].set_xlim([start_date, end_date])
    axes[2].set_xlim([start_date, end_date])
    axes[3].set_xlim([start_date, end_date])
    axes[4].set_xlim([start_date, end_date])
    axes[5].set_xlim([start_date, end_date])

    # 위, 오른쪽 경계 없애기
    axes[0].spines['right'].set_visible(False)
    axes[1].spines['right'].set_visible(False)
    axes[2].spines['right'].set_visible(False)
    axes[3].spines['right'].set_visible(False)
    axes[4].spines['right'].set_visible(False)
    axes[5].spines['right'].set_visible(False)
    axes[0].spines['top'].set_visible(False)
    axes[1].spines['top'].set_visible(False)
    axes[2].spines['top'].set_visible(False)
    axes[3].spines['top'].set_visible(False)
    axes[4].spines['top'].set_visible(False)
    axes[5].spines['top'].set_visible(False)

    # 마지막 종가 출력
    axes[0].text(0.9, 1.1, f'    {df.index[-1]}: {df["KOSPI200"][-1]}', fontsize=8, transform=axes[0].transAxes)
    axes[1].text(0.9, 1.1, f'    {df.index[-1]}: {df["HSCEI"][-1]}', fontsize=8, transform=axes[1].transAxes)
    axes[2].text(0.9, 1.1, f'    {df.index[-1]}: {df["NIKKEI225"][-1]}', fontsize=8, transform=axes[2].transAxes)
    axes[3].text(0.9, 1.1, f'    {df.index[-1]}: {df["S&P500"][-1]}', fontsize=8, transform=axes[3].transAxes)
    axes[4].text(0.9, 1.1, f'    {df.index[-1]}: {df["EUROSTOXX50"][-1]}', fontsize=8, transform=axes[4].transAxes)
    axes[5].text(0.9, 1.1, f'    {df.index[-1]}: {df["CSI300"][-1]}', fontsize=8, transform=axes[5].transAxes)

    # 범례 출력
    axes[0].legend(labels=["KOSPI200"], markerfirst=False, fontsize='x-small', loc=9, bbox_to_anchor=(0.1, 1.3))
    axes[1].legend(labels=["HSCEI"], markerfirst=False, fontsize='x-small', loc=9, bbox_to_anchor=(0.1, 1.3))
    axes[2].legend(labels=["NIKKEI225"], markerfirst=False, fontsize='x-small', loc=9, bbox_to_anchor=(0.1, 1.3))
    axes[3].legend(labels=["S&P500"], markerfirst=False, fontsize='x-small', loc=9, bbox_to_anchor=(0.1, 1.3))
    axes[4].legend(labels=["EUROSTOXX50"], markerfirst=False, fontsize='x-small', loc=9, bbox_to_anchor=(0.1, 1.3))
    axes[5].legend(labels=["CSI300"], markerfirst=False, fontsize='x-small', loc=9, bbox_to_anchor=(0.1, 1.3))

    return f


def cagr(df_data, index_name):

    df_result = pd.DataFrame(index=index_name, columns=["CAGR"])

    for i in index_name:
        df = df_data[i].dropna()

        cagr_val = (df[-1] / df[0]) ** (365/(df.index[-1] - df.index[0]).days)
        df_result.loc[i] = f'{(cagr_val - 1) * 100:.2f}%'

    return df_result


def vol(df_data, index_name):

    df_result = pd.DataFrame(index=index_name, columns=["Vol"])

    for i in index_name:
        df = df_data[i].dropna()

        ar = np.array(df).astype(float)
        ar_return = np.log(ar[1:] / ar[:-1])
        df_result.loc[i] = f'{np.std(ar_return) * np.sqrt(252) * 100:.2f}%'

    return df_result


def mdd(df_data, index_name):

    df_result = pd.DataFrame(index=index_name, columns=["MDD", "MDD_Date"])

    for i in index_name:
        sr = df_data[i].dropna()
        mdd_list = [min(sr[i:]) / sr[i] for i in range(len(sr))]
        min_mdd = min(mdd_list)
        mdd_index = mdd_list.index(min_mdd)
        min_value = min(sr[mdd_index:])
        min_date = sr.index[sr == min_value]
        min_date = min_date[0].strftime("%Y-%m-%d")
        df_result.loc[i] = [f'{(min(mdd_list) - 1) * 100:.2f}%', f'{min_date}']

    return df_result


def print_to_excel():

    wb = xw.Book.caller()
    ws = wb.sheets['main']
    ws1 = wb.sheets['idxinfo']
    ws2 = wb.sheets['chart']

    index = ['KOSPI200', 'HSCEI', 'NIKKEI225', 'S&P500', 'EUROSTOXX50', 'CSI300']

    df_price = get_price_from_sql(date(2000, 1, 1), date.today(), index, type='w')

    df_CAGR = cagr(df_price, index)
    df_Vol = vol(df_price, index)
    df_MDD = mdd(df_price, index)

    df_total = df_CAGR.join(df_Vol.join(df_MDD))

    fig = price_plot(df_price)

    ws1.range("A1").value = date.today()
    ws1.range("A5").value = df_total
    ws2.pictures.add(fig, name='Historical', update=True)


if __name__ == "__main__":
    xw.Book(r"\\172.31.1.222\Deriva\자동화\자동화폴더\구조화증권위험분석\구조화증권위험분석.xlsm").set_mock_caller()
    print_to_excel()

    # index = ['KOSPI200', 'HSCEI', 'NIKKEI225', 'S&P500', 'EUROSTOXX50', 'CSI300']
    #
    # df_price = get_price_from_sql(date(2000, 1, 1), date.today(), index, type='w')
    #
    # df_CAGR = cagr(df_price, index)
    # df_Vol = vol(df_price, index)
    # df_MDD = mdd(df_price, index)
    #
    # df_total = df_CAGR.join(df_Vol.join(df_MDD))
    #
    # fig = price_plot(df_price)
    #
    # new_excel = xw.Book()
    # new_excel.sheets["Sheet1"].range("A5").value = df_total
    # new_excel.sheets["Sheet1"].pictures.add(fig, name='Historical', update=True)
