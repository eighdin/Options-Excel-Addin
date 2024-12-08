from datetime import *
import xlwings as xw
import yfinance as yf
import numpy as np
import stored_values as const

def days_until_friday(date):
    # Get the current weekday (0=Monday, 1=Tuesday, ..., 6=Sunday)
    current_weekday = date.weekday()
    
    # Calculate how many days until Friday (4)
    days_until_friday = (4 - current_weekday) % 7
    
    return days_until_friday

def calculate_price_eow(contract_date, current_price, week_num) -> float:
    if week_num == 1:
        if (datetime.today()-contract_date).days <= const.week_1_day_offset and days_until_friday(datetime.today()) == 0:
            const.stored_price_eow_1 = current_price
        return const.stored_price_eow_1
    if week_num == 2:
        if (datetime.today()-contract_date).days >= const.week_1_day_offset and (datetime.today()-contract_date).days <= const.week_2_day_offset and days_until_friday(datetime.today()) == 0:
            const.stored_price_eow_2 = current_price
        return const.stored_price_eow_2
    if week_num == 3:
        if (datetime.today()-contract_date).days >= const.week_2_day_offset and (datetime.today()-contract_date).days <= const.week_3_day_offset and days_until_friday(datetime.today()) == 0:
            const.stored_price_eow_3 = current_price
        return const.stored_price_eow_3
    if week_num == 4:
        if (datetime.today()-contract_date).days >= const.week_3_day_offset and (datetime.today()-contract_date).days <= const.week_4_day_offset and days_until_friday(datetime.today()) == 0:
            const.stored_price_eow_4 = current_price
        return const.stored_price_eow_4
    if week_num == 5:
        if (datetime.today()-contract_date).days >= const.week_4_day_offset and (datetime.today()-contract_date).days <= const.week_5_day_offset and days_until_friday(datetime.today()) == 0:
            const.stored_price_eow_5 = current_price
        return const.stored_price_eow_5

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
    
    option_type = 'calls' if 'c' in strike_price_in.lower() else 'puts' if 'p' in strike_price_in.lower() else None
    option_frame = getattr(stock.option_chain(stock.options[stock.options.index(exp_date_normal_format_string)]), option_type)
    
    contract = option_frame[option_frame['contractSymbol'].str.contains(exp_date) & np.isclose(option_frame['strike'], strike_price)]
    
    if datetime.today() < exp_date_normal_format:
        current_price = contract['lastPrice'].iloc[0]
        percent_change = (current_price-contract_price)/contract_price
        total = (contract_price*100)*volume
        dollar_change = total*percent_change
        
        price_eow_1 = calculate_price_eow(contract_date, current_price, 1)
        percent_eow_1 = (price_eow_1-contract_price)/contract_price
        price_eow_2 = calculate_price_eow(contract_date, current_price, 2)
        percent_eow_2 = (price_eow_2-contract_price)/contract_price
        price_eow_3 = calculate_price_eow(contract_date, current_price, 3)
        percent_eow_3 = (price_eow_3-contract_price)/contract_price
        price_eow_4 = calculate_price_eow(contract_date, current_price, 4)
        percent_eow_4 = (price_eow_4-contract_price)/contract_price
        price_eow_5 = calculate_price_eow(contract_date, current_price, 5)
        percent_eow_5 = (price_eow_5-contract_price)/contract_price
    
    if datetime.today() >= exp_date_normal_format and np.isclose(const.stored_price_exp, 0.0, 0.00001, 0.00001):
        const.stored_price_exp = current_price
    
    price_exp = const.stored_price_exp
    percent_change_exp = (price_exp-contract_price)/contract_price
    dollar_change_exp = total*percent_change_exp
    
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

for entry in data:
    print(
        f'{entry}\t{data[entry]}'
    )
