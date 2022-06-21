import xlwings as xw
import pandas as pd


# 분기마다 파생결합증권 한도 업데이트 할 때 현재 우리회사가 갖고 있는 파생결합증권 액면 알아야함

def main():
    with xw.App(visible=False) as app:
        db = xw.Book(r'\\172.31.1.222\Deriva\자동화\DB\변액 DATABASE.xlsm')
        df = db.sheets("편입정보").range("A1").options(pd.DataFrame, index=False, expand='table', header=False).value
        db.close()

    df = df.set_index(df[0])
    df = df.rename(columns=df.iloc[0])
    df = df.drop(df.columns[0], axis=1)
    df = df.drop(df.index[0])
    df.index.name = 'ELS ID'

    # 상환된 투자내역 행 제외
    mask = df['진행상태'].isin(['투자 중'])
    df = df[mask]

    df1 = df[['발행사1', '액면금액1', '통화']]
    df1 = df1.rename(columns={'발행사1': '발행사', '액면금액1': '액면금액'})
    df2 = df[['발행사2', '액면금액2', '통화']]
    df2 = df2.rename(columns={'발행사2': '발행사', '액면금액2': '액면금액'})
    df3 = df[['발행사3', '액면금액3', '통화']]
    df3 = df3.rename(columns={'발행사3': '발행사', '액면금액3': '액면금액'})
    df4 = df[['발행사4', '액면금액4', '통화']]
    df4 = df4.rename(columns={'발행사4': '발행사', '액면금액4': '액면금액'})

    df = pd.concat([df1, df2, df3, df4], ignore_index=True)

    df_krw = df[['발행사', '액면금액']][df['통화'] != '달러']
    df_usd = df[['발행사', '액면금액']][df['통화'] == '달러']

    krw_sum = df_krw.groupby('발행사')['액면금액'].sum()

    usd_sum = df_usd.groupby('발행사')['액면금액'].sum()

    return krw_sum, usd_sum


if __name__ == "__main__":
    xw.view(main()[0])
    xw.view(main()[1])
