from datetime import *
import xlwings as xw
import yfinance as yf
import numpy as np
from diskcache import Cache
import ContractInfo
from RepeatTimer import RepeatTimer

# GLOBAL VARS
contract_Info_Dict = {}
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
        contract.contract_Type = contract_symbol[-9:][:1]
    if contract.datetime_Obj is None and contract.date is not None:
        contract.datetime_Obj = datetime.strptime(contract.date, "%m/%d/%y %I:%M%p")
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
    if datetime.today().date() > contract.exp_Datetime_Obj.date():
        try:
            hist_data = yf.download(contract_symbol, contract.datetime_Obj.date(), datetime.now())
            current_Price = hist_data['Close'].iloc[hist_data.__len__()-1].iloc[0]
        except Exception as e:
            print(f'Err: Failed to get hist data for {contract_symbol}\n{e}')
    else:
        # this line uses getattr to get the option chain for the specified contract type in the contract object: stock.options[exp_date].call/put
        option_frame = getattr(stock.option_chain(stock.options[stock.options.index(contract.exp_Date_String)]), contract.contract_Type)
        curr_info = option_frame[contract_symbol.__contains__(contract.exp_Date_From_Symbol) & np.isclose(option_frame['strike'], contract.strike_Price)]
        current_Price = curr_info['lastPrice'].iloc[0]
    contract.current_Price = current_Price
    return contract.current_Price

def fetch_all_curr_prices() -> None:
    '''Updates the current price for all contracts in the contract_Info_Dict'''
    for symbol in contract_Info_Dict:
        try:
            print(f'Got current price of: {symbol} --- {fetch_curent_price(symbol)}')
        except Exception as e:
            print(f'Failed to retrieve current price of: {symbol}!\nErr: {e}')

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

timer_obj = RepeatTimer(refresh_rate_mins, fetch_all_curr_prices)

@xw.func()
def set_Refresh_Rate_Mins(new_refresh_rate):
    global refresh_rate_mins
    if new_refresh_rate > 15:
        refresh_rate_mins = new_refresh_rate
        return f"Refresh rate set to: {new_refresh_rate} mins"
    else:
        return f"Refresh rate must be > 15 mins!"

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
        prices_EOW[n_week]
    else:
        return "N/A"

@xw.func()
def get_Percent_EOW(n_week, caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)

@xw.func()
def get_Percent_Change_Exp(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)

@xw.func()
def get_Dollar_Change_Exp(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)

@xw.func()
def get_High_Post_Buy(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)

@xw.func()
def get_High_Days_Out(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)

@xw.func()
def get_Percent_Change_High(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)

@xw.func()
def get_Dollar_Change_High(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)


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
          |__/   v1.0a1
    """
    )