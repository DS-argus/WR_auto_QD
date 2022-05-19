import pandas as pd
import xlwings as xw


def main():
    app = xw.App()
    app.visible = False
    db = xw.Book(r"\\172.31.1.222\Deriva\자동화\DB\변액 DATABASE.xlsm")
    df = db.sheets("database").range("A1").options(pd.DataFrame, index=False, expand='table', header=False).value
    db.close()
    app.quit()

    df = df.set_index(df[0])
    df = df.rename(columns=df.iloc[0])
    df = df.drop(df.columns[0], axis=1)
    df = df.drop(df.index[0])
    df.index.name = 'ID'

    columns = ['운용사코드', '펀드명', '기준지수결정일', '편입일', '발행사1', '발행사2', '발행사3', '발행사4', '기초자산1', '기초자산2', '기초자산3', '기초자산4']

    df = df[columns]

    # 적립형 안넣으니까 일단 다 없애버리는 형태로
    df = df.drop_duplicates(['펀드명'], keep=False)
    df = df.sort_values(by=['펀드명'])

    file_status = xw.Book.caller()
    sh = file_status.sheets['펀드정보_xlwings']
    sh.range("A2:Q1000").clear_contents()
    sh.range("A1").value = df

    sh.tables.add(source=sh.range("A1").expand('table'), name='펀드정보')


if __name__ == "__main__":
    xw.Book(r"\\172.31.1.222\Deriva\박성민(팀폴더)\본부업무\1. 일간_일간 데일리 현황\ELS 펀드현황\status_excel.xlsx").set_mock_caller()
    main()
