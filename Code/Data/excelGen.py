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

# d = pd.to_datetime("2020-10-05")
# third_fridays(d,3)

def bloombergExcel(ticker, strike, exp_date, start_date, end_date, C_P):
    '''
    Use inputs to create the bloomberg function BDH string for Bloomberg Excel option data
    C_P = "C" or "P" to flag call or put
    '''
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

# s=bloombergExcel("SPY",300.0,pd.to_datetime("2020-10-16"),pd.to_datetime("2020-07-16"),pd.to_datetime("2020-10-16"),"C")
# s
date1 = pd.to_datetime("11-28-2020")

third_fridays(date1,2)

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
        data_dic, data_dic1 = {tick+" Calls":pd.DataFrame({}) for tick in self.tickers},{tick+" Puts":pd.DataFrame({}) for tick in self.tickers}
        data_dic.update(data_dic1)

        # loop through each date to find the start of a new month to find the next
        # 3m atm option
        for date in dates:
            # when we hit a new month, find the third friday in 3months (where option expiries happen)
            if last_month != date.month:
                temp_3m_exp = third_fridays(date,2)[-1] # finds the next 3m atm option
                if temp_3m_exp > pd.to_datetime("today"):
                    break
                # if the third friday wasn't traded on then go to the thursday
                if temp_3m_exp not in dates:
                    temp_3m_exp = temp_3m_exp - timedelta(days=1)
                # finds what the atm strikes were 3m before
                atm_prices = self.price_matrix.loc[date]//1

                for tick in self.tickers:
                    call_str, put_str = tick+" Calls",tick+" Puts"
                    temp_strike = atm_prices[tick]
                    # creates the bloomberg function for the given ticker and dates
                    bloom_str_call = bloombergExcel(tick, temp_strike,temp_3m_exp,date,temp_3m_exp,"C")
                    bloom_str_put = bloombergExcel(tick, temp_strike,temp_3m_exp,date,temp_3m_exp,"P")

                    col_name_call,col_name_put = bloom_str_call[6:-37],bloom_str_put[6:-37]
                    data_dic[call_str][col_name_call],data_dic[put_str][col_name_put] = [bloom_str_call], [bloom_str_put]
                    df_index_call, df_index_put = data_dic[call_str].columns.get_loc(col_name_call),data_dic[put_str].columns.get_loc(col_name_put)
                    data_dic[call_str].insert(df_index_call+1,"","",True)
                    data_dic[put_str].insert(df_index_put+1,"","",True)

            last_month = date.month

        for tick in self.tickers:
            call_str, put_str = tick+" Calls",tick+" Puts"
            data_dic[call_str].to_excel(writer,sheet_name=call_str)
            data_dic[put_str].to_excel(writer,sheet_name=put_str)


        writer.save()

        return data_dic


    def generateData1(self,strike_pct=0,delta_pct = 0,C_P = "C"):
        '''
        Every 3-month ATM option. This works
        '''

        # Create excel file for bloomberg to download data
        writer = pd.ExcelWriter("Option_Data.xlsx")
        # All the dates that the tickers were traded
        dates = self.price_matrix.index
        last_month = 0
        # Dictionary where all data will be stored (dictionary of dataframes for each ticker)
        data_dic, data_dic1 = {tick+" Calls":pd.DataFrame({}) for tick in self.tickers},{tick+" Puts":pd.DataFrame({}) for tick in self.tickers}
        data_dic.update(data_dic1)

        # loop through each date to find the start of a new month to find the next
        # 3m atm option

        temp_3m_exp = 0
        for date in dates:
            # when we hit a new month, find the third friday in 3months (where option expiries happen)
            if date == dates[0] or date > temp_3m_exp:
                temp_3m_exp = third_fridays(date,2)[0] # finds the next 3m atm option
                if temp_3m_exp == pd.to_datetime("2015-01-16"):
                    temp_3m_exp = pd.to_datetime("2015-01-17")

                if temp_3m_exp > pd.to_datetime("today"):
                    break
                # if the third friday wasn't traded on then go to the thursday
                if (temp_3m_exp not in dates) and (temp_3m_exp != pd.to_datetime("2015-01-17")):
                    temp_3m_exp = temp_3m_exp - timedelta(days=1)

                # finds what the atm strikes were 3m before
                atm_prices = round(self.price_matrix.loc[date],0)

                for tick in self.tickers:
                    call_str, put_str = tick+" Calls",tick+" Puts"
                    temp_strike = atm_prices[tick]
                    # creates the bloomberg function for the given ticker and dates
                    bloom_str_call = bloombergExcel(tick, temp_strike,temp_3m_exp,date,temp_3m_exp,"C")
                    bloom_str_put = bloombergExcel(tick, temp_strike,temp_3m_exp,date,temp_3m_exp,"P")

                    col_name_call,col_name_put = bloom_str_call[6:-37],bloom_str_put[6:-37]
                    data_dic[call_str][col_name_call],data_dic[put_str][col_name_put] = [bloom_str_call], [bloom_str_put]
                    df_index_call, df_index_put = data_dic[call_str].columns.get_loc(col_name_call),data_dic[put_str].columns.get_loc(col_name_put)
                    data_dic[call_str].insert(df_index_call+1,"","",True)
                    data_dic[put_str].insert(df_index_put+1,"","",True)



        for tick in self.tickers:
            call_str, put_str = tick+" Calls",tick+" Puts"
            data_dic[call_str].to_excel(writer,sheet_name=call_str)
            data_dic[put_str].to_excel(writer,sheet_name=put_str)


        writer.save()

        return data_dic





data = excelOption(["SPY","XLK","XLF","XLY","XLV","XLI"],"2015-01-01")



data.generateData1()
# d["SPY Puts"]


import os
os.getcwd()
dir ="C:/Users/home/OneDrive/Senior Year/MGT-411/Code/Data/Option_Data.xlsx"

pd.to_datetime("2020-01-05")>pd.to_datetime("2020-01-06")

round(2.6,0)
