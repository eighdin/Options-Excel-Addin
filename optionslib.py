from datetime import *
import xlwings as xw
import yfinance as yf
import numpy as np
from pandas import DataFrame
from diskcache import Cache

import ContractInfo
from RepeatTimer import RepeatTimer

cache = Cache() # define cache that I use to cache the contract dict every time the prices are fetched
# GLOBAL VARS
contract_Info_Dict = cache.get('contract_Dict', {})
refresh_rate_mins = 60

def contract_init(
    contract_symbol: str
) -> ContractInfo:
    """
        Returns the contract object relating to the passed in contract symbol after creating it, if needed.
        Fills in the basic info that can be grabbed from the contract symbol.
    """
    if contract_Info_Dict.get(contract_symbol) is None:
        contract_Info_Dict[contract_symbol] = ContractInfo.ContractInfo()
    contract = contract_Info_Dict[contract_symbol]
    
    if contract.ticker is None: # create ticker info if it doesn't exist
        contract.ticker = "".join(filter(lambda x: not x.isdigit(), contract_symbol[:6]))
    if contract.strike_Price is None: # create strike price info if it doesn't exist
        contract.strike_Price = float(contract_symbol[-8:])/1000
        contract.contract_Type = "calls" if contract_symbol[-9:][:1].lower() == "c" else "puts"
    if contract.contract_Datetime_Obj is None and contract.contract_Date is not None:
        contract.contract_Datetime_Obj = datetime.strptime(contract.contract_Date, "%m/%d/%y %I:%M%p")
    if contract.exp_Datetime_Obj is None: # create expiration date information if it doesn't exist
        contract.exp_Date_From_Symbol = contract_symbol[contract.ticker.__len__():contract.ticker.__len__()+6]
        contract.exp_Datetime_Obj = datetime.strptime(contract.exp_Date_From_Symbol, "%y%m%d")
        contract.exp_Date_String = contract.exp_Datetime_Obj.strftime("%Y-%m-%d")
    if contract.datetime_Obj.date() > contract.exp_Datetime_Obj.date():
        contract.is_Expired = True
    return contract

def get_Contract_Symbol(caller) -> str:
    sheet = caller.sheet
    row = caller.row
    value = sheet.cells(row, 1).value
    return value

def fetch_curent_price(contract_symbol) -> float:
    '''Fetches the current price of a contract, and if it's expired attempts to retrieve historical data'''
    contract = contract_init(contract_symbol)
    stock = yf.Ticker(contract.ticker)
    if datetime.today().date() < contract.exp_Datetime_Obj.date():
        try:
            option_frame = getattr(stock.option_chain(stock.options[stock.options.index(contract.exp_Date_String)]), contract.contract_Type)
            curr_info = option_frame[contract_symbol.__contains__(contract.exp_Date_From_Symbol) & np.isclose(option_frame['strike'], contract.strike_Price)]
            current_Price = curr_info['lastPrice'].iloc[0]
            contract.current_Price = current_Price
            return contract.current_Price
        except Exception as e:
            print(f'Err: Failed to get current price for {contract_symbol}\n{e}')
    else:
        return f"EXP"

def fetch_price_Exp(contract_symbol) -> float:
    contract = contract_init(contract_symbol)
    if datetime.today().date() > contract.exp_Datetime_Obj.date():
        if contract.price_Exp == None:
            try:
                hist_data = yf.download(contract_symbol, start=contract.exp_Datetime_Obj.date()-2, end=contract.exp_Datetime_Obj.date()+1)
                contract.price_Exp= hist_data['Close'].iloc[hist_data.__len__()-1].iloc[0]
            except Exception as e:
                print(f'Failed to get exp price for contract: {contract_symbol}\n{e}')
        return contract.price_Exp
    else:
        return "N/A"

def fetch_High_Data(contract_symbol) -> int:
    contract = contract_init(contract_symbol)
    if (datetime.today().date() - contract.high_Day_Last_Refreshed.date()) > 1:
        contract.high_Day_Last_Refreshed = datetime.today()
        hist_Data = yf.download(contract_symbol, start=contract.datetime_Obj.date(), end=datetime.today().date())
        high_Obj = hist_Data['High'].sort_values(by=contract_symbol, ascending=False).iloc[0]
        contract.high_Post_Buy = high_Obj.iloc[0]
        contract.high_Days_Out = (high_Obj.name.date() - contract.datetime_Obj.date()).days
        contract.percent_Change_High = (contract.high_Post_Buy-contract.orig_Price)/contract.orig_Price
        contract.dollar_Change_High = contract.total*contract.percent_Change_High
        
        return 1
    else:
        return 0

def refresh_Func() -> None:
    '''Updates values for all contracts in the contract_Info_Dict, called by repeat timer object'''
    for symbol in contract_Info_Dict:
        try:
            curr_price = fetch_curent_price(symbol)
            print(f'Got current price of {symbol}: {curr_price:.2f}')
            
            high_Data_Obj = fetch_High_Data(symbol)
            
        except Exception as e:
            print(f'Failed to retrieve refresh info for: {symbol}!\nErr: {e}')
    cache.set('contract_Dict', contract_Info_Dict) # cache the contract info dictionary for use if the spreadsheet gets closed

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

# timer to fetch all prices for every refresh rate in mins
timer_Obj = RepeatTimer(refresh_rate_mins, refresh_Func)

# ====================================================================================================================================================#
# START OF EXCEL USER DEFINED FUNCTIONS (UDFs)
# ====================================================================================================================================================#
@xw.func()
def set_current_price_Refresh_Rate_Mins(new_refresh_rate):
    global refresh_rate_mins
    if new_refresh_rate > 30:
        refresh_rate_mins = new_refresh_rate
        return f"Refresh rate set to: {new_refresh_rate} mins"
    else:
        return f"Refresh rate must be > 30 mins!"

# contract symbol should be first column always

@xw.func()
def set_Contract_Date(contract_date, caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    contract.date = contract_date
    return contract_date

@xw.func()
def get_Ticker(caller):
    contract_symbol = caller.offset(0, -2).value
    contract = contract_init(contract_symbol)
    return contract.ticker


@xw.func()
def get_Strike_Price(caller):
    """
    returns the strike price and the contract type in this format:
    400C
    where the strike price is 400 dollars and it is a call
    """
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    return f'{contract.strike_Price}{contract.contract_Type}'

@xw.func()
def get_Exp_Date(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    return contract.exp_Date_String

@xw.func()
def set_Contract_Price(contract_price, caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    contract.orig_Price = contract_price
    return contract.orig_Price

@xw.func()
def set_Volume(volume, caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    contract.volume = volume
    return volume

@xw.func()
def get_Total(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    contract.total = (contract.orig_Price*100)*contract.volume
    return contract.total

@xw.func()
def get_Current_Price(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    if contract.current_Price is None or np.isclose(0):
        fetch_curent_price(contract_symbol)
    return contract.current_Price

@xw.func()
def get_Percent_Change(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    return (contract.current_Price-contract.orig_Price)/contract.orig_Price

@xw.func()
def get_Dollar_Change(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    contract.dollar_Change = (contract.current_Price-contract.orig_Price)/contract.orig_Price
    return contract.dollar_Change

@xw.func()
def get_Price_EOW(n_week, caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    prices_EOW = contract.price_EOW_List
    friday_Datetime = get_friday_from_date(contract.datetime_Obj, n_week)
    if datetime.today().date() >= friday_Datetime.date():
        if prices_EOW[n_week] == None:
            contract_Data = yf.download(contract_symbol, start=friday_Datetime.date(), end=friday_Datetime.date()+1)
            prices_EOW[n_week] = contract_Data['Close'][contract_symbol].iloc[0]
        else:
            return prices_EOW[n_week]
    else:
        return "N/A"

@xw.func()
def get_Percent_EOW(n_week, caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    percents_EOW = contract.percent_EOW_List
    prices_EOW = contract.price_EOW_List
    if prices_EOW[n_week] != None: # should only return true once price exists, which should only be once that friday has passed
        percents_EOW[n_week] = (prices_EOW[n_week]-contract.orig_Price)/contract.orig_Price
        return percents_EOW[n_week]
    else:
        return "N/A"

@xw.func()
def get_Price_At_Exp(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    if datetime.today().date() > contract.exp_Datetime_Obj.date():
        return fetch_price_Exp(contract_symbol)

@xw.func()
def get_Percent_Change_Exp(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    if contract.price_Exp is not None:
        contract_Price = contract.orig_Price
        price_Exp = contract.price_Exp
        if contract.percent_Change_Exp is None:
            contract.percent_Change_Exp = (price_Exp-contract_Price)/contract_Price
        return contract.percent_Change_Exp
    else:
        return "N/A"

@xw.func()
def get_Dollar_Change_Exp(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    if contract.percent_Change_Exp is not None:
        if contract.dollar_Change_Exp is None:
            contract.dollar_Change_Exp = contract.total * contract.percent_Change_Exp
        return contract.dollar_Change_Exp
    else:
        return "N/A"


@xw.func()
def get_High_Post_Buy(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    return contract.high_Post_Buy

@xw.func()
def get_High_Days_Out(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    return contract.high_Days_Out

@xw.func()
def get_Percent_Change_High(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    return contract.percent_Change_High

@xw.func()
def get_Dollar_Change_High(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    return contract.dollar_Change_High


print(
    r"""


                       /$$     /$$                                     /$$ /$$ /$$      
                      | $$    |__/                                    | $$|__/| $$      
  /$$$$$$   /$$$$$$  /$$$$$$   /$$  /$$$$$$  /$$$$$$$   /$$$$$$$      | $$ /$$| $$$$$$$ 
 /$$__  $$ /$$__  $$|_  $$_/  | $$ /$$__  $$| $$__  $$ /$$_____/      | $$| $$| $$__  $$
| $$  \ $$| $$  \ $$  | $$    | $$| $$  \ $$| $$  \ $$|  $$$$$$       | $$| $$| $$  \ $$
| $$  | $$| $$  | $$  | $$ /$$| $$| $$  | $$| $$  | $$ \____  $$      | $$| $$| $$  | $$
|  $$$$$$/| $$$$$$$/  |  $$$$/| $$|  $$$$$$/| $$  | $$ /$$$$$$$/      | $$| $$| $$$$$$$/
 \______/ | $$____/    \___/  |__/ \______/ |__/  |__/|_______//$$$$$$|__/|__/|_______/ 
          | $$                                                |______/                  
          | $$                                                                          
          |__/   v1.0a2
    """
    )