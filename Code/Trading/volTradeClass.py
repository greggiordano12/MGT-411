import pandas as pd
import numpy as np
import pandas_datareader.data as pdr
from datetime import date, timedelta
import calendar
import scipy.stats as si
import time


def black_scholes(S, K, T, r, sigma, option_type="call"):
    # S: spot price
    # K: strike price
    # T: time to maturity
    # r: interest rate
    # sigma: volatility of underlying asset

    d1 = (np.log(S / K) + (r + 0.5 * sigma ** 2) * T) / (sigma * np.sqrt(T))
    d2 = (np.log(S / K) + (r - 0.5 * sigma ** 2) * T) / (sigma * np.sqrt(T))

    call = (S * si.norm.cdf(d1, 0.0, 1.0) - K * np.exp(-r * T) * si.norm.cdf(d2, 0.0, 1.0))
    put = (K * np.exp(-r * T) * si.norm.cdf(-d2, 0.0, 1.0) - S * si.norm.cdf(-d1, 0.0, 1.0))
    if option_type == "call":
        if call < .01:
            call = .01
        return call
    else:
        if put < .01:
            put = .01
        return put


def bsm_zero(S, K, T, r, sigma, option_price, option_type="call"):
    d1 = (np.log(S / K) + (r + 0.5 * sigma ** 2) * T) / (sigma * np.sqrt(T))
    d2 = (np.log(S / K) + (r - 0.5 * sigma ** 2) * T) / (sigma * np.sqrt(T))

    call = (S * si.norm.cdf(d1, 0.0, 1.0) - K * np.exp(-r * T) * si.norm.cdf(d2, 0.0, 1.0))
    put = (K * np.exp(-r * T) * si.norm.cdf(-d2, 0.0, 1.0) - S * si.norm.cdf(-d1, 0.0, 1.0))
    if option_type == "call":
        return (call - option_price)
    else:
        return (put - option_price)


def secant(S=100, K=100, T=1, r=.02, x1=0, x2=1, option_price=2, option_type="call", E=.001):
    n = 0;
    xm = 0;
    x0 = 0;
    c = 0;
    start = time.time()
    period_of_time = .01
    if (bsm_zero(S, K, T, r, x1, option_price, option_type) * bsm_zero(S, K, T, r, x2, option_price, option_type) < 0):
        while time.time() < start + period_of_time:

            # calculate the intermediate value
            x0 = ((x1 * bsm_zero(S, K, T, r, x2, option_price, option_type) - x2 * bsm_zero(S, K, T, r, x1,
                                                                                            option_price,
                                                                                            option_type)) /
                  (bsm_zero(S, K, T, r, x2, option_price, option_type) - bsm_zero(S, K, T, r, x1, option_price,
                                                                                  option_type)));

            # check if x0 is root of
            # equation or not
            c = bsm_zero(S, K, T, r, x1, option_price, option_type) * bsm_zero(S, K, T, r, x0, option_price,
                                                                               option_type);

            # update the value of interval
            x1 = x2;
            x2 = x0;

            # update number of iteration
            n += 1;

            # if x0 is the root of equation
            # then break the loop
            if (c == 0):
                break;
            xm = ((x1 * bsm_zero(S, K, T, r, x2, option_price, option_type) - x2 * bsm_zero(S, K, T, r, x1,
                                                                                            option_price,
                                                                                            option_type)) /
                  (bsm_zero(S, K, T, r, x2, option_price, option_type) - bsm_zero(S, K, T, r, x1, option_price,
                                                                                  option_type)));

            if (abs(xm - x0) < E):
                break;
        if x0 > 1:
            return -1
        return x0

    else:
        return -1
    return -1


class volTrade:
    def __init__(self, tickers, start_date, end_date=None,
                 data_directory="C:/Users/gregg/OneDrive/Senior Year/MGT-411/Code/Data/"):
        self.tickers = tickers
        self.start_date = start_date
        self.end_date = end_date
        self.option_data_directory = data_directory + "Option_Data.xlsx"
        self.sector_weights_directory = data_directory + "sector_weights.xlsx"
        self.data_directory = data_directory
        self.price_matrix = pdr.DataReader(tickers, "yahoo", start_date, end_date)["Close"].dropna()

        sector_weights_df = pd.read_excel(self.sector_weights_directory, usecols=self.tickers[1:] + ["Date"])
        sector_weights_df = sector_weights_df.set_index("Date").sort_index() / 100
        self.sector_weights_df = sector_weights_df.div(sector_weights_df.sum(axis=1), axis=0)

    def convertBloomExcelCol(self, col):
        ''' Given one of the columns that the our option_data sheet creates outputs strike and expiration date of option

        Inputs: column name from option_data sheet. Example: "SPY 3/20/15 P205"
        Outputs: exp_date, expiration date as pd.datetime object
                 strike: integer that is strike

        '''
        i = 0
        while i < len(col):
            if col[i].isnumeric():
                start_d_ind = i
                i += 1
                while i < len(col):
                    if col[i] == " ":
                        end_d_ind = i
                        i += 1
                        break
                    i += 1
                exp_date = pd.to_datetime(col[start_d_ind:end_d_ind + 1])
                strike = int(col[i + 1:])
                break
            i += 1
        return exp_date, strike

    def optionSeries(self):
        '''
        Reads from the data directory the option data downloaded from Bloomberg and creates
        a dictionary of dictionaries of dictionaries :0

        data[ticker]["Calls" or "Puts"]["Prices" or "Strikes"]["expiration-date-as-string"]

        Shown above will store a continuous time series of every option. Any dates that were missing in the Excel
        File were interpolated by using Black-Scholes ---> make minimum black-Scholes price 1 cent

        Along with option prices, another column of Implied Vols is shown for each as well

        Strikes option gives you strikes for each expiration date
        '''
        xls = pd.ExcelFile(self.option_data_directory)
        data_dic = {tick: {"Calls": {"Strikes": {}, "Prices": {}}, "Puts": {"Strikes": {}, "Prices": {}}} for tick in
                    self.tickers}
        for tick in self.tickers:
            print(tick)
            tick_prices = self.price_matrix[tick]  # get price series for given ticker
            tick_calls = pd.read_excel(xls, tick + " Calls")
            tick_puts = pd.read_excel(xls, tick + " Puts")

            i = 1
            while i < (tick_calls.shape[1] - 1):  # Loop through each expiration date
                temp_calls = tick_calls.iloc[:, i:i + 2]  # Isolate current expiration prices and dates
                temp_calls.columns = ["Date", temp_calls.columns[0]]
                temp_calls = temp_calls.set_index("Date").dropna()

                try:
                    temp_exp, temp_strike = self.convertBloomExcelCol(temp_calls.columns[0])
                    temp_prices = tick_prices.loc[temp_calls.index[0]:temp_exp]
                except:
                    i += 2  # skip one column
                    continue

                all_ivs = [.2]  # set default implied vol at 20% for those that never have an iv
                for j in range(len(temp_prices.index)):
                    temp_date = temp_prices.index[j]
                    # If the option data is missing trading days that the stock was traded, then we will replace the data
                    # with prices using the Black-Scholes Formula
                    if temp_date not in temp_calls.index:
                        S = temp_prices.loc[temp_date]
                        ttm = (temp_exp - temp_date).days / 365
                        iv = np.median(all_ivs)

                        c_price = black_scholes(S, temp_strike, ttm, .01, iv, "call")

                        other_df = pd.DataFrame({"Date": [temp_date], temp_calls.columns[0]: [c_price]})
                        other_df = other_df.set_index("Date")
                        temp_calls = temp_calls.append(other_df, ignore_index=False).sort_index()


                    else:
                        S = temp_prices.loc[temp_calls.index[j]]  # spot price
                        ttm = (temp_exp - temp_calls.index[j]).days / 365
                        c_price = temp_calls[temp_calls.columns[0]].loc[temp_calls.index[j]]
                        iv = secant(S, temp_strike, ttm, .01, 0, 1, c_price, "call", .001)
                        if j == 0 and iv != -1:
                            all_ivs = []

                    if iv == -1 or np.isnan(iv):
                        iv = np.median(all_ivs)  # if IV not found or real, then just use median IV from data
                    all_ivs.append(iv)

                if len(temp_calls.iloc[:, 0]) < len(all_ivs):
                    all_ivs = all_ivs[1:]
                temp_calls["IV"] = all_ivs
                try:
                    temp_calls.loc[pd.to_datetime(str(temp_exp)[:-9])] = max(
                        temp_prices.loc[pd.to_datetime(str(temp_exp)[:-9])] - temp_strike, 0)
                except:
                    greg = "it's ok"
                data_dic[tick]["Calls"]["Prices"][str(temp_exp)[:-9]] = temp_calls
                data_dic[tick]["Calls"]["Strikes"][str(temp_exp)[:-9]] = temp_strike

                i += 2  # skip one column

            ################ Puts Version Now ################################
            i = 1
            while i < (tick_puts.shape[1] - 1):  # Loop through each expiration date
                temp_puts = tick_puts.iloc[:, i:i + 2]  # Isolate current expiration prices and dates
                temp_puts.columns = ["Date", temp_puts.columns[0]]
                temp_puts = temp_puts.set_index("Date").dropna()
                try:
                    temp_exp, temp_strike = self.convertBloomExcelCol(temp_puts.columns[0])
                    temp_prices = tick_prices.loc[temp_puts.index[0]:temp_exp]
                except:
                    i += 2  # skip one column
                    continue

                all_ivs = [.2]  # set default implied vol at 20% for those that never have an iv
                for j in range(len(temp_prices.index)):
                    temp_date = temp_prices.index[j]
                    if temp_date not in temp_puts.index:
                        S = temp_prices.loc[temp_date]
                        ttm = (temp_exp - temp_date).days / 365
                        iv = np.median(all_ivs)

                        p_price = black_scholes(S, temp_strike, ttm, .01, iv, "put")

                        other_df = pd.DataFrame({"Date": [temp_date], temp_puts.columns[0]: [p_price]})
                        other_df = other_df.set_index("Date")
                        temp_puts = temp_puts.append(other_df, ignore_index=False).sort_index()


                    else:
                        S = temp_prices.loc[temp_puts.index[j]]  # spot price
                        ttm = (temp_exp - temp_puts.index[j]).days / 365
                        p_price = temp_puts[temp_puts.columns[0]].loc[temp_puts.index[j]]
                        iv = secant(S, temp_strike, ttm, .01, 0, 1, p_price, "put", .001)
                        if j == 0 and iv != -1:
                            all_ivs = []

                    if iv == -1 or np.isnan(iv):
                        iv = np.median(all_ivs)  # if IV not found or real, then just use median IV from data
                    all_ivs.append(iv)

                if len(temp_puts.iloc[:, 0]) < len(all_ivs):
                    all_ivs = all_ivs[1:]

                temp_puts["IV"] = all_ivs

                other_df = pd.DataFrame({"Date": [temp_date], temp_puts.columns[0]: [p_price]})
                other_df = other_df.set_index("Date")
                temp_puts = temp_puts.append(other_df, ignore_index=False).sort_index()
                try:
                    temp_puts.loc[pd.to_datetime(str(temp_exp)[:-9])] = max(
                        temp_strike - temp_prices.loc[pd.to_datetime(str(temp_exp)[:-9])], 0)
                except:
                    greg = "it's ok"
                data_dic[tick]["Puts"]["Prices"][str(temp_exp)[:-9]] = temp_puts
                data_dic[tick]["Puts"]["Strikes"][str(temp_exp)[:-9]] = temp_strike

                i += 2  # skip one column
        return data_dic

    def all_expTS(self, data=None):
        '''
        Uses data from optionSeries function to create dataframes for each expiration. These dataframes are continuous
        and include all the tickers calls and puts in one data frame. Also outputs a dataframe that does the same for
        implied vols
        '''
        if data == None:
            data = self.optionSeries()

        for tick in self.tickers:
            tick_calls, tick_puts = data[tick]["Calls"], data[tick]["Puts"]
            temp_call_exp_set, temp_put_exp_set = set(tick_calls["Prices"].keys()), set(tick_puts["Prices"].keys())

            if tick == self.tickers[0]:
                call_expiries, put_expiries = temp_call_exp_set, temp_put_exp_set
            else:
                # Find common expirations between total expirations and current ticker
                call_expiries = call_expiries.intersection(temp_call_exp_set)
                put_expiries = put_expiries.intersection(temp_put_exp_set)

        expiries = call_expiries.intersection(put_expiries)  # common set of both call and put expiries
        data_prices = {}
        data_ivs = {}
        for exp in expiries:
            if exp == "1/17/15":
                continue
            temp_option_matrix, temp_iv_matrix = pd.DataFrame({}), pd.DataFrame({})
            for tick in self.tickers:
                tick_calls, tick_puts = data[tick]["Calls"]["Prices"][exp].iloc[:, 0].drop_duplicates(keep="last"), \
                                        data[tick]["Puts"]["Prices"][exp].iloc[:, 0].drop_duplicates(keep='last')
                tick_call_ivs, tick_put_ivs = data[tick]["Calls"]["Prices"][exp].iloc[:, 1].drop_duplicates(
                    keep='last'), data[tick]["Puts"]["Prices"][exp].iloc[:, 1].drop_duplicates(keep='last')
                tick_calls, tick_puts = tick_calls[~tick_calls.index.duplicated(keep='first')], tick_puts[
                    ~tick_puts.index.duplicated(keep="first")]
                tick_call_ivs, tick_put_ivs = tick_call_ivs[~tick_call_ivs.index.duplicated(keep='first')], \
                                              tick_put_ivs[~tick_put_ivs.index.duplicated(keep="first")]
                curr_options, curr_ivs = pd.concat([tick_calls, tick_puts], axis=1), pd.concat(
                    [tick_call_ivs, tick_put_ivs], axis=1)
                temp_option_matrix, temp_iv_matrix = pd.concat([temp_option_matrix, curr_options], axis=1), pd.concat(
                    [temp_iv_matrix, curr_ivs], axis=1)

            temp_iv_matrix.columns = temp_option_matrix.columns
            data_prices[exp], data_ivs[
                exp] = temp_option_matrix.interpolate().dropna(), temp_iv_matrix.interpolate().dropna()

        return data_prices, data_ivs

    def dispersionTest(self, data=None):
        ''' Create simple dispersion backtest.
        Weight trade based on sector weighting file. Short SPY straddle and long sector straddles at the beginning of each month

        '''
        if data == None:
            data = self.all_expTS()

        exps = list(data.keys())
        exps.sort()
        return_series = {"Returns": {}, "Series": {}}
        all_returns = []
        for exp in exps:
            curr_time_series = data[exp]
            s_weights = self.sector_weights_df.loc[curr_time_series.index[0]]
            temp_weights = [-.5, -.5]
            for i in range(len(s_weights)):
                temp_weights.append(s_weights[i] / 2)
                temp_weights.append(s_weights[i] / 2)
            s_weights = np.array(temp_weights)

            daily_returns = curr_time_series.pct_change(1).dropna()
            cum_returns = (daily_returns + 1).cumprod() - 1
            weighted_cum = cum_returns.dot(s_weights)

            return_series["Returns"][exp] = weighted_cum[-1]
            daily_returns = (weighted_cum + 1).pct_change(1).dropna()
            return_series["Series"][exp] = daily_returns

        return return_series

    def dispersionPortfolio(self, option_prices=None, option_ivs=None):
        # slight adjustments to data1 makes this method possible. Data should be continuous now
        if data == None:
            option_prices, option_ivs = self.all_expTS()

        init_port_value = 10000000  # initial portfolio value is 10,000,000
        exps = list(option_prices.keys())
        exps.sort()
        for exp in exps:
            curr_time_series = option_prices[exp]
            s_weights = self.sector_weights_df.loc[curr_time_series.index[0]]
            temp_weights = [-.5, -.5]


tickers = ["SPY", "XLK", "XLF", "XLY", "XLV", "XLI"]
start_date = "2015-01-01"
try_vols = volTrade(tickers, start_date)

data1 = try_vols.optionSeries()

data1

option_prices, option_ivs = try_vols.all_expTS(data=data1)

option_prices.keys()

option_prices

all_portfolios = try_vols.dispersionTest(option_prices)
all_returns = list(all_portfolios["Returns"].values())

np.mean(all_returns)
np.std(all_returns)
all_returns

all_portfolios["Series"]['2020-10-16'].mean() * 31

all_returns
len(all_portfolios["Series"]["2020-10-16"])

#
# exps = list(option_prices.keys())
# exps.sort()
# writer = pd.ExcelWriter("Option_Price_Matrix.xlsx")
# for exp in exps:
#     option_prices[exp].to_excel(writer,sheet_name=exp)
#
# writer.save()
