from idxdata.historical_data import *


start = date(2001, 1, 1)
end = date(2005, 2, 10)

underlying = ['KOSPI200', "EUROSTOXX50"]

print(underlying)

df = get_hist_data_from_sql(start, end, underlying, type="w")

print(df)