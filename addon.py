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
    want_out: str
) -> str | float | int:
    """
    Returns a given options contract indicator, depending on input.
    ______________________________________________________________________________________________________________________
    
    ------------------------------------------------------------Parameters-------------------------------------------------------------------
    ______________________________________________________________________________________________________________________
    
        date_purchased (mm/dd/yy HH:MMam/pm): the date that the contract was purchased, formatted as shown
        contract_symbol (TSLA250221C00290000): example symbol provided, only accepts symbols in this format
        contract_price_in (69.69): price of the contract at purchase, please include the decimal, do not include any other symbols
        volume_in (2000): what is the volume of the contract purchased, whole number only
        
        want_out: what piece of data do you want? ONLY TAKES THESE VERY CASE AND TYPE SENSITIVE INPUTS:
            date: returns the date purchased
            ticker: returns the ticker of the contract
            strike_price: returns the strike price of the contract
            exp_date: returns the expiration date of the contract
            contract_price: returns the initial purchase price of the contract
            volume: returns the volume of the contract purchased
            total: total value of the contract at purchase
            current_price: current price of the contract
            percent_change: percent change from initial purchase price to current price
            dollar_change: dollar change from initial purchase price to current price
            price_eow_{1-5}: the price at the end of the week, from 1-5. examples: price_eow_1 ; price_eow_2 ; etc.
            percent_eow_{1-5}: percent change at the end of the week from 1-5, works the same as price_eow_(1-5)
            high_post_buy: returns the highest value in dollars of the contract after purchase
            high_days_out: returns how many days out that high was
            percent_change_high: percent change from initial purchase price to high price
            dollar_change_high: dollar change from initial purchase price to high price
    ______________________________________________________________________________________________________________________
    
    """
    parameters_dont_need_compute = [
        "date", "ticker", "strike_price", "exp_date", "contract_price", "volume", "total"
    ]
    doesnt_need_compute = any(parameter in want_out for parameter in parameters_dont_need_compute)
    
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
    ticker = "".join(filter(lambda x: not x.isdigit(), contract_symbol[:6]))
    strike_price = float(contract_symbol[-8:])/1000
    contract_price = np.float64(contract_price_in)
    volume = int(volume_in)
    total = (contract_price*100)*volume
    option_type_abbrev = contract_symbol[-9:][:1]
    option_type = 'calls' if option_type_abbrev.lower() == "c" else 'puts'
    exp_date_from_symbol = contract_symbol[ticker.__len__():ticker.__len__()+6]
    exp_datetime_obj = datetime.strptime(exp_date_from_symbol, "%y%m%d")
    exp_date_string = exp_datetime_obj.strftime("%Y-%m-%d")
    contract_datetime_obj = datetime.strptime(date_purchased, "%m/%d/%y %I:%M%p")
    is_expired = datetime.today() > exp_datetime_obj
    stock = yf.Ticker(ticker=ticker) # reference to stock that we're looking at
    today_to_the_hour = datetime.strptime(datetime.today().strftime("%Y-%m-%d %I:00%p"), "%Y-%m-%d %I:00%p")
    
    if not doesnt_need_compute: # if needs compute
        if not is_expired:
            # grab all data for the stock from when it was purchased till today
            hist_data = download_hist_info_cached(contract_symbol, start=contract_datetime_obj.date(), end=today_to_the_hour)
            # grabs call or put data depending on input provided in strike price in by accessing the stocks option chain
            option_frame = getattr(stock.option_chain(stock.options[stock.options.index(exp_date_string)]), option_type)
            # Gets a reference to the specific option contract that we want to get info for, downloads all historical data between purchase and expiration
            contract = option_frame[contract_symbol.__contains__(exp_date_from_symbol) & np.isclose(option_frame['strike'], strike_price)]
        else:
            try:
                hist_data = download_hist_info_cached(contract_symbol, start=contract_datetime_obj.date(), end=exp_datetime_obj.date())
            except:
                return f"{contract_symbol} is expired and data is not cached! Please enter data manually."
        current_price = contract['lastPrice'].iloc[0] if not is_expired else hist_data['Close'].iloc[hist_data.__len__()-1].iloc[0]
        percent_change = (current_price-contract_price)/contract_price
        dollar_change = total*percent_change

        # Calculates the price at the end of the weeks using historical data grabbed between when the contract was bought and sold
        # This only goes up to 5 right now, can add more later
        if want_out == "price_eow_1" or "percent_eow_1":
            price_eow_1 = calculate_price_eow(hist_data, contract_datetime_obj, current_price, stored_nums, 1)
            percent_eow_1 = (price_eow_1-contract_price)/contract_price
        if want_out == "price_eow_2" or "percent_eow_2":
            price_eow_2 = calculate_price_eow(hist_data, contract_datetime_obj, current_price, stored_nums, 2)
            percent_eow_2 = (price_eow_2-contract_price)/contract_price
        if want_out == "price_eow_3" or "percent_eow_3":
            price_eow_3 = calculate_price_eow(hist_data, contract_datetime_obj, current_price, stored_nums, 3)
            percent_eow_3 = (price_eow_3-contract_price)/contract_price
        if want_out == "price_eow_4" or "percent_eow_4":
            price_eow_4 = calculate_price_eow(hist_data, contract_datetime_obj, current_price, stored_nums, 4)
            percent_eow_4 = (price_eow_4-contract_price)/contract_price
        if want_out == "price_eow_5" or "percent_eow_5":
            price_eow_5 = calculate_price_eow(hist_data, contract_datetime_obj, current_price, stored_nums, 5)
            percent_eow_5 = (price_eow_5-contract_price)/contract_price
        
        if want_out.__contains__("high"):
            # high maths and stuff
            high_post_buy_object = hist_data['High'].sort_values(by=contract_symbol, ascending=False).iloc[0]
            high_post_buy_dollar = high_post_buy_object.iloc[0]
            high_days_out = (high_post_buy_object.name.date() - contract_datetime_obj.date()).days
            percent_change_high = (high_post_buy_dollar-contract_price)/contract_price
            dollar_change_high = total*percent_change_high
    
    if is_expired and np.isclose(stored_price_exp, 0.0000):
        stored_price_exp = current_price
        price_exp = stored_price_exp
        percent_change_exp = (price_exp-contract_price)/contract_price
        dollar_change_exp = total*percent_change_exp
    
    price_exp = stored_price_exp
    percent_change_exp = stored_percent_change_exp
    dollar_change_exp = stored_dollar_change_exp
    
    # want_out_case = {
    #     "date":date_purchased, "ticker":ticker, "strike_price":strike_price, "exp_date":exp_date_string, 
    #     "contract_price":contract_price, "volume":volume, "total":total, "current_price":current_price, "percent_change":percent_change, "dollar_change":dollar_change, 
    #     "price_eow_1":price_eow_1, "percent_eow_1":percent_eow_1, "price_eow_2":price_eow_2, "percent_eow_2":percent_eow_2, "price_eow_3":price_eow_3, "percent_eow_3":percent_eow_3, "price_eow_4":price_eow_4, "percent_eow_4":percent_eow_4, "price_eow_5":price_eow_5, "percent_eow_5":percent_eow_5,
    #     "price_exp":price_exp, "percent_change_exp":percent_change_exp, "dollar_change_exp":dollar_change_exp,
    #     "high_post_buy":high_post_buy_dollar, "high_days_out":high_days_out, "percent_change_at_high":percent_change_at_high, "dollar_change_at_high":dollar_change_at_high
    # }
    try:
        return locals()[want_out]
    except KeyError:
        return f"Wrong key! ({want_out}) Check what you requested to get out of the function."
    except Exception as e:
        return f"Something went wrong! Idk, ask eighdin. ERROR: {e}"

print(
    r"""
 _________  ___  ___  ________  ___       _______      
|\___   ___\\  \|\  \|\   __  \|\  \     |\  ___ \     
\|___ \  \_\ \  \\\  \ \  \|\  \ \  \    \ \   __/|    
     \ \  \ \ \   __  \ \  \\\  \ \  \    \ \  \_|/__  
      \ \  \ \ \  \ \  \ \  \\\  \ \  \____\ \  \_|\ \ 
       \ \__\ \ \__\ \__\ \_______\ \_______\ \_______\
        \|__|  \|__|\|__|\|_______|\|_______|\|_______|
    v0.2
    """,
    f"\n\ncache directory is in same directory as your spreadsheet. \n\n\tIt's called: {cache._directory}\n\nprobably don't touch it though, it'll mess up anything that expired and is still displaying info\n\n\nyou're welcome chase"
)