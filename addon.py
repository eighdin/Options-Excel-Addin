from datetime import *
import xlwings as xw
import yfinance as yf
import numpy as np

stored_eow_prices = {
    "stored_price_eow_1":0.0, 
    "stored_price_eow_2":0.0, 
    "stored_price_eow_3":0.0, 
    "stored_price_eow_4":0.0, 
    "stored_price_eow_5":0.0
}
stored_price_exp = 0.0
stored_percent_change_exp = 0.0
stored_dollar_change_exp = 0.0

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

def calculate_price_eow(hist_data, contract, contract_date, current_price, week_num) -> float:
    global stored_eow_prices
    fridaydatetime = get_friday_from_date(contract_date, week_num)
    stored_price_eow = stored_eow_prices[f'stored_price_eow_{week_num}']
    
    if fridaydatetime.date() == datetime.today().date() and np.isclose(stored_price_eow, 0.0):
        stored_price_eow = current_price
    elif fridaydatetime.date() < datetime.today().date() and np.isclose(stored_price_eow, 0.0) and fridaydatetime<=datetime.today():
        for entry in hist_data['Close'].iloc:
            if entry.name.date() == fridaydatetime.date():
                stored_price_eow = entry.iloc[0]
    return stored_price_eow

@xw.func(forcefullcalc=True)
def options_data(date_purchased, ticker, strike_price_in, exp_date, contract_price_in, volume_in):
    global stored_price_exp
    global stored_percent_change_exp
    global stored_dollar_change_exp
    global stored_eow_prices
    # get reference to stock in question that we're trading on, define strike price at the top because we need it for contract definition
    stock = yf.Ticker(ticker=ticker)
    strike_price = np.float64(strike_price_in[:-1][1:])
    
    # Handle all date stuff that we need in the function
    exp_date_normal_format_string = datetime.strptime(exp_date, "%y%m%d").strftime("%Y-%m-%d")
    exp_date_normal_format = datetime.strptime(exp_date, "%y%m%d")
    contract_date = datetime.strptime(date_purchased, "%m/%d/%y %I:%M%p")
    is_expired = datetime.today() > exp_date_normal_format
    
    # grabs call or put data depending on input provided in strike price in by accessing the stocks option chain
    option_type = 'calls' if 'c' in strike_price_in.lower() else 'puts' if 'p' in strike_price_in.lower() else None
    option_frame = getattr(stock.option_chain(stock.options[stock.options.index(exp_date_normal_format_string)]), option_type)
    
    # Gets a reference to the specific option contract that we want to get info for, downloads all historical data between purchase and expiration
    contract = option_frame[option_frame['contractSymbol'].str.contains(exp_date) & np.isclose(option_frame['strike'], strike_price)]
    contract_symbol = contract['contractSymbol'].iloc[0]
    hist_data = yf.download(contract['contractSymbol'].iloc[0], start=contract_date.date(), end=exp_date_normal_format.date())
    
    contract_price = np.float64(contract_price_in)
    volume = int(volume_in)
    current_price = contract['lastPrice'].iloc[0]
    percent_change = (current_price-contract_price)/contract_price
    total = (contract_price*100)*volume
    dollar_change = total*percent_change
    
    # Calculates the price at the end of the weeks using historical data grabbed between when the contract was bought and sold
    # This only goes up to 5 right now, can add more later
    price_eow_1 = calculate_price_eow(hist_data, contract, contract_date, current_price, 1)
    price_eow_2 = calculate_price_eow(hist_data, contract, contract_date, current_price, 2)
    price_eow_3 = calculate_price_eow(hist_data, contract, contract_date, current_price, 3)
    price_eow_4 = calculate_price_eow(hist_data, contract, contract_date, current_price, 4)
    price_eow_5 = calculate_price_eow(hist_data, contract, contract_date, current_price, 5)
    
    percent_eow_1 = (price_eow_1-contract_price)/contract_price
    percent_eow_2 = (price_eow_2-contract_price)/contract_price
    percent_eow_3 = (price_eow_3-contract_price)/contract_price
    percent_eow_4 = (price_eow_4-contract_price)/contract_price
    percent_eow_5 = (price_eow_5-contract_price)/contract_price
    
    # high maths and stuff
    high_post_buy_object = hist_data['High'].sort_values(by=contract_symbol, ascending=False).iloc[0]
    high_post_buy_dollar = high_post_buy_object.iloc[0]
    high_days_out = (high_post_buy_object.name.date() - contract_date.date()).days
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
        "date":date_purchased, "ticker":ticker, "strike_price":strike_price_in, "exp_date":exp_date_normal_format_string, 
        "contract_price":contract_price, "volume":volume, "total":total, "current_price":current_price, "percent_change":percent_change, "dollar_change":dollar_change, 
        "price_eow_1":price_eow_1, "percent_eow_1":percent_eow_1, "price_eow_2":price_eow_2, "percent_eow_2":percent_eow_2, "price_eow_3":price_eow_3, "percent_eow_3":percent_eow_3, "price_eow_4":price_eow_4, "percent_eow_4":percent_eow_4, "price_eow_5":price_eow_5, "percent_eow_5":percent_eow_5,
        "price_exp":price_exp, "percent_change_exp":percent_change_exp, "dollar_change_exp":dollar_change_exp,
        "high_post_buy":high_post_buy_dollar, "high_days_out":high_days_out, "percent_change_at_high":percent_change_at_high, "dollar_change_at_high":dollar_change_at_high
    }
    return output_dict.values()


