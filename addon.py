from datetime import *
import xlwings as xw
import yfinance as yf
import numpy as np
import pandas as pd
from diskcache import Cache

cache = Cache('hist_cache')

def days_until_friday(
    date: datetime
) -> int:
    # Get the current weekday (0=Monday, 1=Tuesday, ..., 6=Sunday)
    current_weekday = date.weekday()
    
    # Calculate how many days until Friday (4)
    days_until_friday = (4 - current_weekday) % 7
    
    return days_until_friday

def get_friday_from_date(
    start_date: datetime, 
    which_friday: int
) -> datetime:
    first_friday = start_date + timedelta(days=days_until_friday(start_date))
    nth_friday = first_friday + timedelta(weeks=which_friday-1)
    return nth_friday

def calculate_price_eow(
    hist_data: pd.DataFrame, 
    contract_datetime_obj: datetime, 
    current_price: np.float64,
    stored_nums: dict,
    week_num: int
) -> float:
    stored_eow_prices = stored_nums['stored_eow_prices']

    fridaydatetime = get_friday_from_date(contract_datetime_obj, week_num)
    stored_price_eow = stored_eow_prices[f'stored_price_eow_{week_num}']
    
    if fridaydatetime.date() == datetime.today().date() and np.isclose(stored_price_eow, 0.0):
        stored_price_eow = current_price
    elif fridaydatetime.date() < datetime.today().date() and np.isclose(stored_price_eow, 0.0) and fridaydatetime<=datetime.today():
        for entry in hist_data['Close'].iloc:
            if entry.name.date() == fridaydatetime.date():
                stored_price_eow = entry.iloc[0]
    return stored_price_eow

@cache.memoize()
def download_hist_info_cached(contract_symbol, start, end):
    return yf.download(contract_symbol, start, end)

@xw.func()
def options_data(
    date_purchased: str, 
    contract_symbol: str, 
    contract_price_in: float, 
    volume_in: int,
) -> str | float | int:
    stored_nums = {
        "stored_eow_prices":{
            "stored_price_eow_1":0.0, 
            "stored_price_eow_2":0.0, 
            "stored_price_eow_3":0.0, 
            "stored_price_eow_4":0.0, 
            "stored_price_eow_5":0.0
        },
        "stored_price_exp":0.0,
        "stored_percent_change_exp":0.0,
        "stored_dollar_change_exp":0.0
    }
    stored_price_exp = stored_nums["stored_price_exp"]
    stored_percent_change_exp = stored_nums["stored_percent_change_exp"]
    stored_dollar_change_exp = stored_nums["stored_dollar_change_exp"]
    
    # symbol slicing and basic variable definition
    strike_price = float(contract_symbol[-8:])/1000
    ticker = "".join(filter(lambda x: not x.isdigit(), contract_symbol[:6]))
    option_type_abbrev = contract_symbol[-9:][:1]
    option_type = 'calls' if option_type_abbrev.lower() == "c" else 'puts'
    exp_date_from_symbol = contract_symbol[ticker.__len__():ticker.__len__()+6]
    exp_datetime_obj = datetime.strptime(exp_date_from_symbol, "%y%m%d")
    exp_date_string = exp_datetime_obj.strftime("%Y-%m-%d")
    contract_datetime_obj = datetime.strptime(date_purchased, "%m/%d/%y %I:%M%p")
    is_expired = datetime.today() > exp_datetime_obj
    stock = yf.Ticker(ticker=ticker) # reference to stock that we're looking at
    
    if not is_expired:
        # grab all data for the stock from when it was purchased till today
        hist_data = download_hist_info_cached(contract_symbol, start=contract_datetime_obj.date(), end=datetime.today())
        # grabs call or put data depending on input provided in strike price in by accessing the stocks option chain
        option_frame = getattr(stock.option_chain(stock.options[stock.options.index(exp_date_string)]), option_type)
        # Gets a reference to the specific option contract that we want to get info for, downloads all historical data between purchase and expiration
        contract = option_frame[contract_symbol.__contains__(exp_date_from_symbol) & np.isclose(option_frame['strike'], strike_price)]
    else:
        hist_data = download_hist_info_cached(contract_symbol, start=contract_datetime_obj.date(), end=exp_datetime_obj.date())
    
    current_price = contract['lastPrice'].iloc[0] if not is_expired else hist_data['Close'].iloc[hist_data.__len__()-1].iloc[0]
    
    contract_price = np.float64(contract_price_in)
    volume = int(volume_in)
    percent_change = (current_price-contract_price)/contract_price
    total = (contract_price*100)*volume
    dollar_change = total*percent_change
    
    # Calculates the price at the end of the weeks using historical data grabbed between when the contract was bought and sold
    # This only goes up to 5 right now, can add more later
    price_eow_1 = calculate_price_eow(hist_data, contract_datetime_obj, current_price, stored_nums, 1)
    price_eow_2 = calculate_price_eow(hist_data, contract_datetime_obj, current_price, stored_nums, 2)
    price_eow_3 = calculate_price_eow(hist_data, contract_datetime_obj, current_price, stored_nums, 3)
    price_eow_4 = calculate_price_eow(hist_data, contract_datetime_obj, current_price, stored_nums, 4)
    price_eow_5 = calculate_price_eow(hist_data, contract_datetime_obj, current_price, stored_nums, 5)
    
    percent_eow_1 = (price_eow_1-contract_price)/contract_price
    percent_eow_2 = (price_eow_2-contract_price)/contract_price
    percent_eow_3 = (price_eow_3-contract_price)/contract_price
    percent_eow_4 = (price_eow_4-contract_price)/contract_price
    percent_eow_5 = (price_eow_5-contract_price)/contract_price
    
    # high maths and stuff
    high_post_buy_object = hist_data['High'].sort_values(by=contract_symbol, ascending=False).iloc[0]
    high_post_buy_dollar = high_post_buy_object.iloc[0]
    high_days_out = (high_post_buy_object.name.date() - contract_datetime_obj.date()).days
    percent_change_at_high = (high_post_buy_dollar-contract_price)/contract_price
    dollar_change_at_high = total*percent_change_at_high
    
    if is_expired and np.isclose(stored_price_exp, 0.0, 0.00001, 0.00001):
        stored_price_exp = current_price
        price_exp = stored_price_exp
        percent_change_exp = (price_exp-contract_price)/contract_price
        dollar_change_exp = total*percent_change_exp
    
    price_exp = stored_price_exp
    percent_change_exp = stored_percent_change_exp
    dollar_change_exp = stored_dollar_change_exp
    
    output_dict = {
        "date":date_purchased, "ticker":ticker, "strike_price":strike_price, "exp_date":exp_date_string, 
        "contract_price":contract_price, "volume":volume, "total":total, "current_price":current_price, "percent_change":percent_change, "dollar_change":dollar_change, 
        "price_eow_1":price_eow_1, "percent_eow_1":percent_eow_1, "price_eow_2":price_eow_2, "percent_eow_2":percent_eow_2, "price_eow_3":price_eow_3, "percent_eow_3":percent_eow_3, "price_eow_4":price_eow_4, "percent_eow_4":percent_eow_4, "price_eow_5":price_eow_5, "percent_eow_5":percent_eow_5,
        "price_exp":price_exp, "percent_change_exp":percent_change_exp, "dollar_change_exp":dollar_change_exp,
        "high_post_buy":high_post_buy_dollar, "high_days_out":high_days_out, "percent_change_at_high":percent_change_at_high, "dollar_change_at_high":dollar_change_at_high
    }
    return output_dict.values()
