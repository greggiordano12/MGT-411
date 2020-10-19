import pandas as pd
import pandas_datareader.data as pdr
from datetime import date, timedelta
import calendar




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
third_fridays(d,3)


# def returns_matrix(tickers, start_date, end_date):
#     #NEED TO CHANGE FOR FACT THAT TICKERS HAVE VERY DIFFERENT DATA THAT THEY GO BACK TO#
#     '''
#     Creates a daily returns dataframe where each column is a different stock
#     '''
#     min_week = pd.to_datetime("1900-01-01") #sets initial week as today
#     rmat = []
#     data_dic = {}
#     count = 0
#     for tick in tickers:
#         temp_df = pdr.DataReader(tick, "yahoo", start_date, end_date)
#         temp_dates = temp_df.index
#         data_dic[count] = temp_df
#
#         if temp_dates[0] > min_week:
#             min_week = temp_dates[0]
#         count +=1
#     for i in range(len(tickers)):
#         temp_df = data_dic[i]
#         temp_dates = temp_df.index
#         if temp_dates[0] != min_week:
#             drop_index = list(temp_dates).index(min_week)
#             temp_df = temp_df.drop(temp_dates[:drop_index])
#         temp_returns = np.log(temp_df.Close) - np.log(temp_df.Close.shift(1))
#         temp_returns = list(temp_returns[1:])
#         rmat = rmat + [temp_returns]
#     rmat = np.array(rmat)
#     dates = temp_df.index
#     rdf=pd.DataFrame({tickers[i]:rmat[i] for i in range(len(tickers))})
#     rdf.index = dates[1:]
#     return rdf



def bloombergExcel(ticker, strike, exp_date, start_date, end_date, C_P):
    strike = str(strike)[:-2]
    exp_date = str(exp_date.month) +"/" + str(exp_date.day) + "/" + str(exp_date.year)[-2:]
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

class excelOption:

    def __init__(self, tickers, start_date, end_date=None):
        self.tickers = tickers
        self.start_date = start_date
        self.end_date = end_date
        self.price_matrix = pdr.DataReader(tickers, "yahoo",start_date,end_date)["Close"].dropna()


    def generateData(self,strike_pct=0,delta_pct = 0,C_P = "C"):
        '''
        Every 3-month ATM option. This works
        '''
        # Create excel file for bloomberg to download data
        writer = pd.ExcelWriter("Option_Data.xlsx")
        # All the dates that the tickers were traded
        dates = self.price_matrix.index
        last_month = 0
        # Dictionary where all data will be stored (dictionary of dataframes for each ticker)
        data_dic = {tick:pd.DataFrame({}) for tick in self.tickers}

        # loop through each date to find the start of a new month to find the next
        # 3m atm option
        for date in dates:
            # when we hit a new month, find the third friday in 3months (where option expiries happen)
            if last_month != date.month:
                temp_3m_exp = third_fridays(date,3)[-1] # finds the next 3m atm option
                if temp_3m_exp > pd.to_datetime("today"):
                    break
                # if the third friday wasn't traded on then go to the thursday
                if temp_3m_exp not in dates:
                    temp_3m_exp = temp_3m_exp - timedelta(days=1)
                # finds what the atm strikes were 3m before
                atm_prices = self.price_matrix.loc[date]//1

                for tick in self.tickers:
                    temp_strike = atm_prices[tick]
                    # creates the bloomberg function for the given ticker and dates
                    bloom_str = bloombergExcel(tick, temp_strike,temp_3m_exp,date,temp_3m_exp,C_P)
                    col_name = bloom_str[6:-37] # look more into on column
                    data_dic[tick][col_name] = [bloom_str]
                    df_index = data_dic[tick].columns.get_loc(col_name)
                    data_dic[tick].insert(df_index+1,"","",True)

            last_month = date.month

        for tick in self.tickers:
            data_dic[tick].to_excel(writer,sheet_name=tick+"_"+C_P)

        writer.save()

        return data_dic








data = excelOption(["SPY","XLK","MSFT"],"2015-01-01")
d=data.generateData()
d["SPY"]
