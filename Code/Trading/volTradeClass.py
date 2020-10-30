import pandas as pd
import numpy as np
import pandas_datareader.data as pdr
from datetime import date, timedelta
import calendar
import scipy.stats as si

def black_scholes(S, K, T, r, sigma, option_type="call"):

    #S: spot price
    #K: strike price
    #T: time to maturity
    #r: interest rate
    #sigma: volatility of underlying asset

    d1 = (np.log(S / K) + (r + 0.5 * sigma ** 2) * T) / (sigma * np.sqrt(T))
    d2 = (np.log(S / K) + (r - 0.5 * sigma ** 2) * T) / (sigma * np.sqrt(T))

    call = (S * si.norm.cdf(d1, 0.0, 1.0) - K * np.exp(-r * T) * si.norm.cdf(d2, 0.0, 1.0))
    put = (K * np.exp(-r * T) * si.norm.cdf(-d2, 0.0, 1.0) - S * si.norm.cdf(-d1, 0.0, 1.0))
    if option_type == "call":
        return call
    else:
        return put



def bsm_zero(S, K, T, r, sigma, option_price, option_type="call"):
    d1 = (np.log(S / K) + (r + 0.5 * sigma ** 2) * T) / (sigma * np.sqrt(T))
    d2 = (np.log(S / K) + (r - 0.5 * sigma ** 2) * T) / (sigma * np.sqrt(T))

    call = (S * si.norm.cdf(d1, 0.0, 1.0) - K * np.exp(-r * T) * si.norm.cdf(d2, 0.0, 1.0))
    put = (K * np.exp(-r * T) * si.norm.cdf(-d2, 0.0, 1.0) - S * si.norm.cdf(-d1, 0.0, 1.0))
    if option_type == "call":
        return (call - option_price)
    else:
        return (put-option_price)

def secant(S=100,K=100,T=1,r=.02,x1=0, x2=1, option_price=2, option_type = "call",E=.001):
    n = 0; xm = 0; x0 = 0; c = 0;
    if (bsm_zero(S,K,T,r,x1,option_price,option_type) * bsm_zero(S,K,T,r,x2,option_price,option_type) < 0):
        while True:

            # calculate the intermediate value
            x0 = ((x1 * bsm_zero(S,K,T,r,x2,option_price,option_type) - x2 * bsm_zero(S,K,T,r,x1,option_price,option_type)) /
                            (bsm_zero(S,K,T,r,x2,option_price,option_type) - bsm_zero(S,K,T,r,x1,option_price,option_type)));

            # check if x0 is root of
            # equation or not
            c = bsm_zero(S,K,T,r,x1,option_price,option_type) * bsm_zero(S,K,T,r,x0,option_price,option_type);

            # update the value of interval
            x1 = x2;
            x2 = x0;

            # update number of iteration
            n += 1;

            # if x0 is the root of equation
            # then break the loop
            if (c == 0):
                break;
            xm = ((x1 * bsm_zero(S,K,T,r,x2,option_price,option_type) - x2 * bsm_zero(S,K,T,r,x1,option_price,option_type)) /
                            (bsm_zero(S,K,T,r,x2,option_price,option_type) - bsm_zero(S,K,T,r,x1,option_price,option_type)));

            if(abs(xm - x0) < E):
                break;

        return x0

    else:
        return -1

class volTrade:
    def __init__(self, tickers, start_date, end_date=None,data_directory = "C:/Users/home/OneDrive/Senior Year/MGT-411/Code/Data/Option_Data.xlsx"):
        self.tickers = tickers
        self.start_date = start_date
        self.end_date = end_date
        self.data_directory = data_directory
        self.price_matrix = pdr.DataReader(tickers, "yahoo",start_date,end_date)["Close"].dropna()


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
                i+=1
                while i < len(col):
                    if col[i] == " ":
                        end_d_ind = i
                        i+=1
                        break
                    i+=1
                exp_date = pd.to_datetime(col[start_d_ind:end_d_ind+1])
                strike = int(col[i+1:])
                break
        i+=1
        return exp_date, strike


    def atm_dispersion(self):
        xls = pd.ExcelFile(self.data_directory)

        for tick in tickers:
            tick_prices = self.price_matrix[tick] # get price series for given ticker
            tick_calls = pd.read_excel(xls,tick + " Calls")
            tick_puts = pd.read_excel(xls, tick+ " Puts")

            i = 1
            while i < (tick_calls.shape[1]-1): #Loop through each expiration date
                temp_calls = tick_calls.iloc[:,i:i+2] #Isolate current expiration prices and dates
                temp_calls.columns = ["Date", temp_calls.columns[0]]
                temp_calls= temp_calls.set_index("Date")
                temp_exp, temp_strike = self.convertBloomExcelCol(temp_calls.columns[0])
                temp_prices = tick_prices.loc[temp_calls.index[0]:exp_date]

                # for i in range(len(temp_prices.index)):
                #     if temp_prices.index[i] not in
                all_ivs = []
                for i in range(len(temp_calls.index)):
                    S = temp_prices.loc[temp_calls.index[i]] # spot price
                    ttm = (temp_exp - temp_calls.index[i]).days/365
                    c_price = temp_calls.loc[temp_calls.index[i]]
                    iv = secant(S,temp_strike,ttm,.01,0,1, c_price,"call",.001)
                    all_ivs.append(iv)


                temp_calls["IV"] = all_ivs
                return temp_calls



                i+=2 # skip one column







tickers = ["SPY","XLK","XLF","XLY","XLV","XLI"]
start_date = "2015-01-01"
try_vols = volTrade(tickers,start_date)

try_vols.atm_dispersion()



dir ="C:/Users/home/OneDrive/Senior Year/MGT-411/Code/Data/Option_Data.xlsx"
xls = pd.ExcelFile(dir)
df1=pd.read_excel(xls,"SPY Puts")

trial = df1.iloc[:,1:3]
trial = trial.dropna()
x = trial.set_index(trial.columns[0])
df1.shape

trial.columns = ["Date", trial.columns[0]]
trial = trial.set_index("Date")
trial.columns[0]

col = trial.columns[0]

col

SPY = pdr.DataReader("SPY","yahoo","2015-01-01")["Close"]
SPY.loc[pd.to_datetime("2015-01-01"):exp_date]

exp_date
strike
(exp_date - pd.to_datetime("2015-01-01")).days
