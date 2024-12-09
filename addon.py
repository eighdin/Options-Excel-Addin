from datetime import *
import xlwings as xw
import yfinance as yf
import numpy as np
import stored_values as const

def days_until_friday(date) -> int:
    # Get the current weekday (0=Monday, 1=Tuesday, ..., 6=Sunday)
    current_weekday = date.weekday()
    
    # Calculate how many days until Friday (4)
    days_until_friday = (4 - current_weekday) % 7
    
    return days_until_friday

def get_friday_from_date(start_date: datetime, which_friday: int) -> datetime:
    first_friday = start_date + timedelta(days=days_until_friday(start_date))
    nth_friday = first_friday + timedelta(weeks=which_friday-1)
    return nth_friday

def calculate_price_eow(contract, contract_date, current_price, week_num) -> float:
    fridaydatetime = get_friday_from_date(contract_date, week_num)
    stored_price_eow = const.stored_eow_prices[f'stored_price_eow_{week_num}']
    
    if fridaydatetime.date() == datetime.today().date() and np.isclose(stored_price_eow, 0.0):
        stored_price_eow = current_price
    elif fridaydatetime.date() < datetime.today().date() and np.isclose(stored_price_eow, 0.0) and fridaydatetime<=datetime.today():
        hist_data = yf.download(contract['contractSymbol'].iloc[0], start=fridaydatetime.date(), end=fridaydatetime.date()+timedelta(days=1))
        for entry in hist_data['Close'].iloc:
            if entry.name.date() == fridaydatetime.date():
                stored_price_eow = entry.iloc[0]
    return stored_price_eow

@xw.func(forcefullcalc=True)
@xw.ret(index=False, header=False)
def options_data(date_purchased, ticker, strike_price_in, exp_date, contract_price_in, volume_in):
    contract_price = np.float64(contract_price_in)
    volume = int(volume_in)
    strike_price = np.float64(strike_price_in[:-1][1:])
    stock = yf.Ticker(ticker=ticker)
    exp_date_normal_format_string = datetime.strptime(exp_date, "%y%m%d").strftime("%Y-%m-%d")
    exp_date_normal_format = datetime.strptime(exp_date, "%y%m%d")
    contract_date = datetime.strptime(date_purchased, "%m/%d/%y %I:%M%p")
    is_expired = datetime.today() > exp_date_normal_format
    
    option_type = 'calls' if 'c' in strike_price_in.lower() else 'puts' if 'p' in strike_price_in.lower() else None
    option_frame = getattr(stock.option_chain(stock.options[stock.options.index(exp_date_normal_format_string)]), option_type)
    
    contract = option_frame[option_frame['contractSymbol'].str.contains(exp_date) & np.isclose(option_frame['strike'], strike_price)]
    

    current_price = contract['lastPrice'].iloc[0]
    percent_change = (current_price-contract_price)/contract_price
    total = (contract_price*100)*volume
    dollar_change = total*percent_change
    
    price_eow_1 = calculate_price_eow(contract, contract_date, current_price, 1)
    price_eow_2 = calculate_price_eow(contract, contract_date, current_price, 2)
    price_eow_3 = calculate_price_eow(contract, contract_date, current_price, 3)
    price_eow_4 = calculate_price_eow(contract, contract_date, current_price, 4)
    price_eow_5 = calculate_price_eow(contract, contract_date, current_price, 5)
    
    percent_eow_1 = (price_eow_1-contract_price)/contract_price
    percent_eow_2 = (price_eow_2-contract_price)/contract_price
    percent_eow_3 = (price_eow_3-contract_price)/contract_price
    percent_eow_4 = (price_eow_4-contract_price)/contract_price
    percent_eow_5 = (price_eow_5-contract_price)/contract_price

    
    if is_expired and np.isclose(const.stored_price_exp, 0.0, 0.00001, 0.00001):
        const.stored_price_exp = current_price
    
    if is_expired:
        price_exp = const.stored_price_exp
        percent_change_exp = (price_exp-contract_price)/contract_price
        dollar_change_exp = total*percent_change_exp
    else:
        price_exp = const.stored_price_exp
        percent_change_exp = 0.0
        dollar_change_exp = 0.0
    
    output_dict = {
        "date":date_purchased, "ticker":ticker, "strike_price":strike_price_in, "exp_date":exp_date_normal_format_string, 
        "contract_price":contract_price, "volume":volume, "total":total, "current_price":current_price, "percent_change":percent_change, "dollar_change":dollar_change, 
        "price_eow_1":price_eow_1, "percent_eow_1":percent_eow_1, "price_eow_2":price_eow_2, "percent_eow_2":percent_eow_2, "price_eow_3":price_eow_3, "percent_eow_3":percent_eow_3, "price_eow_4":price_eow_4, "percent_eow_4":percent_eow_4, "price_eow_5":price_eow_5, "percent_eow_5":percent_eow_5,
        "price_exp":price_exp, "percent_change_exp":percent_change_exp, "dollar_change_exp":dollar_change_exp
    }
    return output_dict


my_ticker = yf.Ticker(ticker="NVDA")
option_frame = my_ticker.option_chain(my_ticker.options[0]).calls
contract = option_frame[np.isclose(option_frame['strike'], 230)]
lastPrice = contract['lastPrice']
thisdate = datetime.strptime("11/22/24 10:34am", "%m/%d/%y %I:%M%p")
data = options_data("11/08/24 12:40pm", "TSLA", "$290C", "250221", "54.96", "39600")
print(data)
# hist = yf.download('TSLA250221C00290000', start=thisdate.date(), end=thisdate.date()+timedelta(days=1))

# for entry in hist['Close'].iloc:
#     #if entry.name.strftime("%Y-%m-%d").__contains__("2024-11-15"):
#         print(entry.iloc[0])

# option_symbol = "AAPL230120C00150000"  # Example option symbol (modify as needed)

# # Download historical data for the specific option
# historical_data = yf.download(option_symbol, start="2023-01-01", end="2023-12-31")

# print(historical_data)