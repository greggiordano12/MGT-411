import pandas as pd
import pandas_datareader.data as pdr
from datetime import date, timedelta
import calendar

spy = pdr.DataReader("SPY","yahoo","2015-01-01")
spy.head()


def next_third_friday(d):
    """ Given a third friday find next third friday"""
    d += timedelta(weeks=4)
    return d if d.day >= 15 else d + timedelta(weeks=1)

def third_fridays(d, n):
    """Given a date, calculates n next third fridays. Will use for option expiration because
    every third friday is the main option expiration date

    """

    # Find closest friday to 15th of month
    s = date(d.year, d.month, 15)
    result = [s + timedelta(days=(calendar.FRIDAY - s.weekday()) % 7)]

    # This month's third friday passed. Find next.
    if result[0] < d:
        result[0] = next_third_friday(result[0])

    for i in range(n - 1):
        result.append(next_third_friday(result[-1]))

    return result
d = pd.to_datetime("2020-10-05")
third_fridays(d,3)[-1]


def returns_matrix(tickers, start_date, end_date):
    #NEED TO CHANGE FOR FACT THAT TICKERS HAVE VERY DIFFERENT DATA THAT THEY GO BACK TO#
    '''
    Creates a daily returns dataframe where each column is a different stock
    '''
    min_week = pd.to_datetime("1900-01-01") #sets initial week as today
    rmat = []
    data_dic = {}
    count = 0
    for tick in tickers:
        temp_df = pdr.DataReader(tick, "yahoo", start_date, end_date)
        temp_dates = temp_df.index
        data_dic[count] = temp_df

        if temp_dates[0] > min_week:
            min_week = temp_dates[0]
        count +=1
    for i in range(len(tickers)):
        temp_df = data_dic[i]
        temp_dates = temp_df.index
        if temp_dates[0] != min_week:
            drop_index = list(temp_dates).index(min_week)
            temp_df = temp_df.drop(temp_dates[:drop_index])
        temp_returns = np.log(temp_df.Close) - np.log(temp_df.Close.shift(1))
        temp_returns = list(temp_returns[1:])
        rmat = rmat + [temp_returns]
    rmat = np.array(rmat)
    dates = temp_df.index
    rdf=pd.DataFrame({tickers[i]:rmat[i] for i in range(len(tickers))})
    rdf.index = dates[1:]
    return rdf

writer = pd.ExcelWriter("HW1_Data.xlsx")
stock_data.to_excel(writer,sheet_name=tick+"_spot")
calls_df.to_excel(writer, sheet_name= tick + "_calls")
puts_df.to_excel(writer, sheet_name = tick+"_puts")
writer.save()

def bloombergExcel(ticker, strike, exp_date, start_date, end_date, C_P):
    strike = str(strike)[:-2]
    exp_date = str(exp_date.month) +"/" + str(exp_date.day) + "/" + str(exp_date.year)[:-2]
    m,d = str(start_date.month), str(start_date.day)
    if len(str(start_date.month)) < 2 or len(str(start_date.day))<2:
        if len(str(start_date.month)) < 2:
            m = "0" +str(start_date.month)
        if len(str(start_date.day))<2:
            d = "0" +str(start_date.day)
    start_date = str(start_date.year) + m + d

    m,d = str(end_date.month), str(end_date.day)
    if len(str(end_date.month)) < 2 or len(str(end_date.day))<2:
        if len(str(end_date.month)) < 2:
            m = "0" +str(end_date.month)
        if len(str(end_date.day))<2:
            d = "0" +str(end_date.day)
    end_date = str(end_date.year) + m + d

    s = f"""=BDH("{ticker} {exp_date} {C_P}{strike} Equity","PX_LAST",{start_date},{end_date})"""
    return s

s=bloombergExcel("SPY",300.0,pd.to_datetime("2020-10-16"),pd.to_datetime("2020-07-16"),pd.to_datetime("2020-10-16"),"C")
s
s[-51:-37]
s = "hello"
a = "Hi"
''+" s "+a

var_a = 10

f"""This is my quoted variable: "{var_a}". "{s}" """
class excelOption:

    def __init__(self, tickers, start_date, end_date=None):
        self.tickers = tickers
        self.start_date = start_date
        self.end_date = end_date
        self.price_matrix = pdr.DataReader(tickers, "yahoo",start_date,end_date)["Close"].dropna()


    def generateData(self,strike_pct,delta_pct = 0,C_P = "C"):
        '''
        Every 3-month ATM option. This works
        '''
        writer = pd.ExcelWriter("Option_Data.xlsx")
        dates = self.price_matrix.index
        last_month = 0
        data_dic = {tick:pd.DataFrame({}) for tick in self.tickers}
        for date in dates:

            if last_month != date.month:
                temp_3m_exp = third_fridays(date,3)[-1]
                if temp_3m_exp > pd.to_datetime("today"):
                    break
                if temp_3m_exp not in dates:
                    temp_3m_exp = temp_3m_exp - timedelta(days=1)

                atm_prices = self.price_matrix.loc[date]//1

                for tick in self.tickers:
                    temp_strike = atm_prices[tick]
                    bloom_str = bloombergExcel(tick, temp_strike,temp_3m_exp,date,temp_3m_exp,C_P)
                    col_name = bloom_str[-51:-37]
                    data_dic[tick][col_name] = [bloom_str]
                    df_index = data_dic[tick].columns.get_loc(col_name)
                    data_dic[tick].insert(df_index+1,"","",True)

            last_month = date.month

        for tick in self.tickers:
            data_dic[tick].to_excel(writer,sheet_name=tick+"_"+C_P)

        writer.save()

        return data_dic



s = 'hello "greg" how are you'




data = excelOption(["SPY","XLK","AAPL"],"2020-01-01")
d=data.generateData(0)
d["SPY"]



df.to_excel(writer)
writer.save()

import os
print(os.getcwd())
str(x[0])[:-2]
dic = {}
lst = ["hello","bye","g"]
dic = {t:{} for t in lst}
dic["hello"]["h"] = [1,2,3]
dic["hello"]["h"]
df = pd.DataFrame({"a":[1,2,3],"b":[1,2,3],"c":[3,4,5]})
df.insert(2,"newcol","",allow_duplicates=True)
df.insert(1,"","",True)
df["aa"] = [1,2,3]
df = pd.DataFrame({})
df["a"] = [1,2,3]
df
