from datetime import *
import xlwings as xw
import yfinance as yf
import numpy as np

def days_until_friday(date):
    # Get the current weekday (0=Monday, 1=Tuesday, ..., 6=Sunday)
    current_weekday = date.weekday()
    
    # Calculate how many days until Friday (4)
    days_until_friday = (4 - current_weekday) % 7
    
    return days_until_friday

@xw.func(forcefullcalc=True)
@xw.ret(index=False, header=False)
def options_data(date_purchased, ticker, strike_price, exp_date, contract_price, volume):
    stock = yf.Ticker(ticker=ticker)
    exp_date_normal_format = datetime.strptime(exp_date, "%y%m%d").strftime("%Y-%m-%d")
    contract_date = datetime.strptime(date_purchased, "%m/%d/%y %I:%M%p")
    
    option_type = 'calls' if 'c' in strike_price.lower() else 'puts' if 'p' in strike_price.lower() else None
    option_frame = getattr(stock.option_chain(stock.options[stock.options.index(exp_date_normal_format)]), option_type)
    
    contract = option_frame[option_frame['contractSymbol'].str.contains(exp_date) & np.isclose(option_frame['strike'], strike_price)]
    current_price = contract['lastPrice']
    percent_change = float(lastPrice-contract_price)/contract_price
    total = (contract_price*100)*volume
    dollar_change = total*percent_change

    # PUT LOGIC FOR CALCULATING PRICE AT THE END OF EACH WEEK 
    
    output_dict = {
        "date":date, "ticker":ticker, "strike_price":strike_price, "exp_date":exp_date_normal_format, 
        "contract_price":contract_price, "volume":volume, "total":total, "current_price":current_price, "percent_change":percent_change, "dollar_change":dollar_change, 
        "price_eow_1":0, "percent_eow_1":0, "price_eow_2":0, "percent_eow_2":0, "price_eow_3":0, "percent_eow_3":0, "price_eow_4":0, "percent_eow_4":0, "price_eow_5":0, "percent_eow_5":0
    }
    return list(output_dict.values)


my_ticker = yf.Ticker(ticker="NVDA")
option_frame = my_ticker.option_chain(my_ticker.options[0]).calls
contract = option_frame[np.isclose(option_frame['strike'], 230)]
lastPrice = contract['lastPrice']
thisdate = datetime.strptime("11/22/24 10:34am", "%m/%d/%y %I:%M%p")
print(
    days_until_friday(thisdate)
    
)
