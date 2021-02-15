import pandas as pd
import pandas_datareader.data as pdr
import time


def daily_close_unadjusted(tickers, start, end=None):
    """
    Returns a dataframe of daily unadjusted close prices for a list of tickers. Index of df is the date and columns are the tickers.
    """
    df = pd.DataFrame()

    for i, ticker in enumerate(tickers):
        if i % 5 == 0 and i != 0:
            print("Sleep.")
            time.sleep(61)

        data = pdr.DataReader(ticker, "av-daily", start=start, end=end, api_key="5COZWJ1DAXKO1S4Y")
        df[ticker] = data["close"]

    return df


def getDeltaDf(ivs, prices, spotPrices):
    # Parsing option string for expiration
    cols = ivs.columns
    col = cols[1].split(" ")
    date = pd.to_datetime(col[1])

    # Adding a time to expiration col
    ivs['time'] = (ivs['Date'].apply(lambda x: len(pd.bdate_range(pd.to_datetime(x), date))) - 1) / 252
    prices['time'] = (prices['Date'].apply(lambda x: len(pd.bdate_range(pd.to_datetime(x), date))) - 1) / 252

    print(spotPrices.to_string())
    print(ivs.to_string())
    print(prices.to_string())

    strike = col[2][1:]


if __name__ == "__main__":
    ivPath = rf"C:\Users\John\Desktop\example_option_ivs.csv"
    pricesPath = fr"C:\Users\John\Desktop\example_option_prices.csv"
    spotPath = fr"C:\Users\John\Desktop\spotPrices.csv"

    ivDf = pd.read_csv(ivPath)
    priceDf = pd.read_csv(pricesPath)
    spotDf = pd.read_csv(spotPath)

    getDeltaDf(ivDf, priceDf, spotDf)

    # daily_close_unadjusted(['SPY', 'XLC', 'XLY', 'XLP', 'XLE', 'XLF', 'XLV', 'XLI', 'XLB', 'XLRE', 'XLK', 'XLU'],
    #                        dt.date(2015, 1, 1)).to_csv(rf"C:\Users\John\Desktop\spotPrices.csv")
