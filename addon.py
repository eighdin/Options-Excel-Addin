from datetime import *
import xlwings as xw
import yfinance as yf
import numpy as np
import pandas as pd

@xw.func
@xw.ret(index=False, header=False)
def options_data(ticker, strike_price, exp_date, contract_price, volume):
    output_df = pd.DataFrame({"exp_date":"", "contract_price":0, "volume":0, "current_price":0.0, "percent_change":0.0})
    stock = yf.Ticker(ticker=ticker)
    exp_date_normal_format = datetime.strptime(exp_date, "%y%m%d").strftime("%Y-%m-%d")
    
    option_type = 'calls' if 'c' in strike_price.lower() else 'puts' if 'p' in strike_price.lower() else None
    frame = getattr(stock.option_chain(stock.options[stock.options.index(exp_date_normal_format)]), option_type)
    
    contract = frame[frame['contractSymbol'].str.contains(exp_date) & np.isclose(frame['strike'], strike_price)]


my_ticker = yf.Ticker(ticker="NVDA")
frame = my_ticker.option_chain(my_ticker.options[4]).calls
output_df = pd.DataFrame({"exp_date":"", "contract_price":0, "volume":0, "current_price":0.0, "percent_change":0.0})

print(
    # frame[np.isclose(frame['strike'], 230)]
    output_df
)
