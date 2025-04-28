from datetime import *
import xlwings as xw
import yfinance as yf
import numpy as np
from diskcache import Cache

from ContractInfo import ContractInfo
from RepeatTimer import RepeatTimer
from PluginConfig import PluginConfig

config = PluginConfig()
cache = Cache(
        'cache_dir',
        disk_min_file_size=0,
        size_limit=1073741824,
        cull_limit=100000
    ) # define cache that I use to cache the contract dict every time the prices are fetched

_already_printed_logo = False
if not _already_printed_logo:
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
             |__/   v2.0
    """
    )

def update_cache(contract_obj: ContractInfo, symbol: str):
    if contract_obj is not None:
        contract_Dict: dict[str, ContractInfo] = cache.get('contract_Dict', {})
        contract_Dict[symbol] = contract_obj
        cache.set('contract_Dict', contract_Dict)
    else:
        print("Didn't update cache with empty contract object.")

def contract_init(
    contract_symbol: str,
    date_in : str = None
) -> ContractInfo:
    """
        Returns the contract object relating to the passed in contract symbol after creating it, if needed.
        Fills in the basic info that can be grabbed from the contract symbol.
    """
    contract_Dict: dict[str, ContractInfo] = cache.get('contract_Dict', {})
    if contract_Dict.get(contract_symbol) is None:
        contract_Dict[contract_symbol] = ContractInfo()
    contract = contract_Dict[contract_symbol]
    
    try:
        if contract.ticker is None: # create ticker info if it doesn't exist
            contract.ticker = "".join(filter(lambda x: not x.isdigit(), contract_symbol[:6]))
        if contract.strike_Price is None: # create strike price info if it doesn't exist
            contract.strike_Price = float(contract_symbol[-8:])/1000
            contract.contract_Type = "calls" if contract_symbol[-9:][:1].lower() == "c" else "puts"
        if contract.datetime_Obj is None and date_in is not None:
            contract.date = date_in
            contract.datetime_Obj = datetime.strptime(contract.date, "%m/%d/%y %I:%M%p")
            contract.high_Day_Last_Refreshed = contract.datetime_Obj # set last day refreshed to the day the contract was bought so that it'll only refresh starting the day after
        if contract.exp_Datetime_Obj is None: # create expiration date information if it doesn't exist
            contract.exp_Date_From_Symbol = contract_symbol[contract.ticker.__len__():contract.ticker.__len__()+6]
            contract.exp_Datetime_Obj = datetime.strptime(contract.exp_Date_From_Symbol, "%y%m%d")
            contract.exp_Date_String = contract.exp_Datetime_Obj.strftime("%Y-%m-%d")
        if contract.high_Post_Buy is None: # set initial values for high if they're not set yet
            contract.high_Post_Buy = contract.orig_Price
            contract.high_Days_Out = 0
        if datetime.today().date() > contract.exp_Datetime_Obj.date():
            contract.is_Expired = True
        return contract
    except Exception as e:
        print(f"Error in contract init!\n{e}\n")

def get_Contract_Symbol(caller) -> str:
    sheet = caller.sheet
    row = caller.row
    value = sheet.cells(row, 1).value
    return value

def fetch_curent_price(contract_symbol) -> ContractInfo:
    '''Fetches the current price of a contract, and if it's expired attempts to retrieve historical data. Returns contract info object w current price updated'''
    contract = contract_init(contract_symbol)
    stock = yf.Ticker(contract.ticker)
    try:
        if datetime.today().date() < contract.exp_Datetime_Obj.date():
            option_frame = getattr(stock.option_chain(stock.options[stock.options.index(contract.exp_Date_String)]), contract.contract_Type)
            curr_info = option_frame[contract_symbol.__contains__(contract.exp_Date_From_Symbol) & np.isclose(option_frame['strike'], contract.strike_Price)]
            current_Price = curr_info['lastPrice'].iloc[0]
            contract.current_Price = current_Price
        return contract
    except Exception as e:
            print(f'Err: Failed to get current price for {contract_symbol}\n{e}\n')
    finally:
        update_cache(contract, contract_symbol)

def fetch_price_Exp(contract_symbol) -> float:
    contract = contract_init(contract_symbol)
    if datetime.today().date() > contract.exp_Datetime_Obj.date():
        if contract.price_Exp == None:
            try:
                hist_data = yf.download(contract_symbol, start=contract.exp_Datetime_Obj.date()-timedelta(days=2), end=contract.exp_Datetime_Obj.date()+timedelta(days=1))
                contract.price_Exp= hist_data['Close'].iloc[hist_data.__len__()-1].iloc[0]
            except Exception as e:
                print(f'Failed to get exp price for contract: {contract_symbol}\n{e}\n Setting value to current price if cached.')
                if contract.current_Price is not None:
                    contract.price_Exp = contract.current_Price
                else:
                    contract.price_Exp = 0.0
        update_cache(contract, contract_symbol)
        return contract.price_Exp
    else:
        return "N/A"

def fetch_High_Data(contract_symbol):
    contract = contract_init(contract_symbol)
    try:
        if not (contract.total, contract.orig_Price).__contains__(None):
            if (datetime.today().date() - contract.high_Day_Last_Refreshed.date()).days >= 1 or contract.current_Price > contract.high_Post_Buy: # if its been at least one day since high data was retrieved, retrieve it
                contract.high_Day_Last_Refreshed = datetime.today()
                hist_Data = yf.download(contract_symbol, start=contract.datetime_Obj, end=datetime.today())
                high_Obj = hist_Data['High'].sort_values(by=contract_symbol, ascending=False).iloc[0]
                contract.high_Post_Buy = high_Obj.iloc[0]
                contract.high_Days_Out = (high_Obj.name.date() - contract.datetime_Obj.date()).days
                # COMMENTED OUT BC CELL FORMULAS
                # contract.percent_Change_High = (contract.high_Post_Buy-contract.orig_Price)/contract.orig_Price
                # contract.dollar_Change_High = contract.total*contract.percent_Change_High
        else:
            print("Failed to fetch high data! Missing data")
    except Exception as e:
        print(f"Error getting high data!\n{e}\n")
    finally:
        update_cache(contract, contract_symbol)

def refresh_Func() -> None:
    '''Updates values for all contracts in the contract_Info_Dict, called by repeat timer object'''
    contract_Dict: dict[str, ContractInfo] = cache.get('contract_Dict')
    for symbol, contract in contract_Dict.items():
        if not contract.is_Expired:
            try:
                fetch_curent_price(symbol) # updates cache inside this method
                fetch_High_Data(symbol) # updates cache inside this method
                print(f"Refreshed data for {symbol}")
            except Exception as e:
                print(f'Failed to refresh data for: {symbol}!\nErr: {e}')

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
timer_Obj = RepeatTimer(config.refresh_rate_mins, refresh_Func)

# ====================================================================================================================================================#
# START OF EXCEL USER DEFINED FUNCTIONS (UDFs)
# ====================================================================================================================================================#

@xw.func()
def set_Refresh_Rate_Mins(new_refresh_rate):
    if new_refresh_rate > config.min_refresh_rate_mins:
        config.refresh_rate_mins = new_refresh_rate
        print(config.refresh_rate_mins)
        return f"Refresh rate set to: {new_refresh_rate} mins"
    else:
        return f"Refresh rate must be > {config.min_refresh_rate_mins} mins!"

# contract symbol should be first column always

@xw.func()
def set_Contract_Date(contract_date, caller):
    """expects contract date in format:%m/%d/%y %I:%M%p"""
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol, date_in=contract_date)
    try:
        return contract.date
    except Exception as e:
        print(f"Error setting contract date: {e}")
        return f"Error!"
    finally:
        update_cache(contract, contract_symbol)

@xw.func()
def get_Ticker(caller):
    contract_symbol = get_Contract_Symbol(caller)
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
    try:
        return f'{contract.strike_Price:.0f}{"C" if contract.contract_Type == "calls" else "P"}'
    except Exception as e:
        print(f"Error getting strike price!\n{e}\n")

@xw.func()
def get_Exp_Date(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    try:
        return datetime.strftime(contract.exp_Datetime_Obj, "%m/%d/%y")
    except Exception as e:
        print(f'Error getting exp date!\n{e}\n')
        return "Error! Check console!"

# DONT NEED BC CELL FORMULAS
@xw.func()
def set_Contract_Price(contract_price, caller):
    """Expects contract price in """
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    try:
        contract.orig_Price = np.float64(contract_price)
        return contract.orig_Price
    except Exception as e:
        print(f'Error setting contract price!\n{e}\n')
        return "Error setting contract price! Check console."
    finally:
        update_cache(contract, contract_symbol)

# @xw.func()
# def set_Volume(volume, caller):
#     contract_symbol = get_Contract_Symbol(caller)
#     contract = contract_init(contract_symbol)
#     try:
#         contract.volume = volume
#         return contract.volume
#     except Exception as e:
#         print(f"Error setting volume for {contract_symbol}!\n{e}\n")
#         return "Error!"
#     finally:
#         update_cache(contract, contract_symbol)

# @xw.func()
# def get_Total(caller):
#     contract_symbol = get_Contract_Symbol(caller)
#     contract = contract_init(contract_symbol)
#     try:
#         if contract.orig_Price is not None:
#             contract.total = (contract.orig_Price*100)*contract.volume
#             return contract.total
#         else:
#             return "Set contract price!"
#     except Exception as e:
#         print(f'Error getting total! (issue is probably one of the values not being set)\n{e}\n')
#         return "Error!"
#     finally:
#         update_cache(contract, contract_symbol)

@xw.func(volatile=True)
def get_Current_Price(caller):
    """
    Returns last cached current price.
    If not expired and no current price stored, gets that number from yf. 
    if expired, otherwise returns "EXP". 
    """
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    try:
        if not contract.is_Expired:
            if contract.current_Price is None:
                contract = fetch_curent_price(contract_symbol)
            return contract.current_Price
        else:
            if contract.current_Price is None:
                return "EXP"
            else:
                return contract.current_Price
    except Exception as e:
        print(f"Error getting current price!\n{e}\n")
        return "Error!"
    finally:
        update_cache(contract, contract_symbol)

# DONT NEED BC CELL FORMULAS
# @xw.func(volatile=True)
# def get_Percent_Change(caller):
#     """
#     Returns the percent change of the contract from original price to current price if cached or not expired.
#     If expired and no cached value returns Expired!
#     Otherwise, returns cached value.
#     """
#     contract_symbol = get_Contract_Symbol(caller)
#     contract = contract_init(contract_symbol)
#     try:
#         if not contract.is_Expired:
#             if contract.current_Price is not None and contract.orig_Price is not None:
#                 contract.percent_Change =  (contract.current_Price-contract.orig_Price)/contract.orig_Price
#                 return contract.percent_Change
#             else:
#                 return "Current price is None!"
#         else:
#             if contract.percent_Change is not None:
#                 return contract.percent_Change
#             else:
#                 return "Expired!"
#     except Exception as e:
#         print(f'Error getting percent change!\n{e}\n')
#         return "Error!"
#     finally:
#         update_cache(contract, contract_symbol)

# @xw.func(volatile=True)
# def get_Dollar_Change(caller):
#     """
#     Returns the dollar change of the contract if cached or not expired.
#     If expired and no cached value returns Expired!
#     Otherwise, returns cached value.
#     """
#     contract_symbol = get_Contract_Symbol(caller)
#     contract = contract_init(contract_symbol)
#     try:
#         if not contract.is_Expired:
#             if contract.percent_Change is not None:
#                 contract.dollar_Change = contract.total*contract.percent_Change
#                 return contract.dollar_Change
#             else:
#                 return "Percent change not calculated"
#         else:
#             if contract.dollar_Change is not None:
#                 return contract.dollar_Change
#             else:
#                 return "Expired!"
#     except Exception as e:
#         print(f"Error getting dollar change!\n{e}\n")
#         return "Error!"
#     finally:
#         update_cache(contract, contract_symbol)

@xw.func(volatile=True)
def get_Price_EOW(n_week, caller):
    """
    Returns the price of the contract at the end of n_week from when it was bought
    Returns N/A if not that week yet
    Returns EXP if no cached value and expired
    """
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    try:
        n_week = int(n_week)
        prices_EOW = contract.price_EOW_List
        price_EOW = prices_EOW[n_week-1] # -1 because zero index
        if not contract.is_Expired:
            friday_Datetime = get_friday_from_date(contract.datetime_Obj, n_week)
            if datetime.today().date() >= friday_Datetime.date():
                if price_EOW == 0:
                    contract_Data = yf.download(contract_symbol, start=friday_Datetime.date(), end=friday_Datetime.date() + timedelta(days=1))
                    price_EOW = contract_Data['Close'][contract_symbol].iloc[0]
                    prices_EOW[n_week-1] = price_EOW
                    return price_EOW
                else:
                    return price_EOW
            else:
                return "N/A"
        else:
            if not np.isclose(price_EOW, 0):
                return price_EOW
            else:
                return "EXP"
    except Exception as e:
        print(f"Error getting price eow!\n{e}\n")
        return "Error!"
    finally:
        update_cache(contract, contract_symbol)


# THIS COMMENTED OUT BC USING CELL FORMULAS INSTEAD
# @xw.func(volatile=True)
# def get_Percent_EOW(n_week, caller):
#     try:
#         n_week = int(n_week)
#         contract_symbol = get_Contract_Symbol(caller)
#         contract = contract_init(contract_symbol)
#         percents_EOW = contract.percent_EOW_List
#         prices_EOW = contract.price_EOW_List
#         price_EOW = prices_EOW[n_week-1] # zero index
#         if not np.isclose(price_EOW, 0): # should only return true once price exists, which should only be once that friday has passed
#             percents_EOW[n_week-1] = (price_EOW-contract.orig_Price)/contract.orig_Price
#             return percents_EOW[n_week-1]
#         else:
#             return "N/A"
#     except Exception as e:
#         print(f'Error getting percentage EOW {n_week}!!!\n{e}\n')
#         return "Error!"
#     finally:
#         update_cache(contract, contract_symbol)

@xw.func(volatile=True)
def get_Price_At_Exp(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    try:
        if datetime.today().date() > contract.exp_Datetime_Obj.date():
            if contract.current_Price not in (0, None):
                return contract.current_Price # return last stored current price
            else:
                return "No stored values"
        else:
            return "N/A"
    except Exception as e:
        print(f"Error getting price at exp!\n{e}\n")
        return "Error!"

# DON'T NEED BC CELL FORMULAS
# @xw.func(volatile=True)
# def get_Percent_Change_Exp(caller):
#     contract_symbol = get_Contract_Symbol(caller)
#     contract = contract_init(contract_symbol)
#     try:
#         if contract.price_Exp is not None:
#             contract_Price = contract.orig_Price
#             price_Exp = contract.price_Exp
#             if contract.percent_Change_Exp is None:
#                 contract.percent_Change_Exp = (price_Exp-contract_Price)/contract_Price
#             return contract.percent_Change_Exp
#         else:
#             return "N/A"
#     except Exception as e:
#         print(f"Error getting percent change at exp!\n{e}\n")
#         return "Error!"
#     finally:
#         update_cache(contract, contract_symbol)

# @xw.func(volatile=True)
# def get_Dollar_Change_Exp(caller):
#     contract_symbol = get_Contract_Symbol(caller)
#     contract = contract_init(contract_symbol)
#     try:
#         if contract.percent_Change_Exp is not None:
#             if contract.dollar_Change_Exp is None:
#                 contract.dollar_Change_Exp = contract.total * contract.percent_Change_Exp
#             return contract.dollar_Change_Exp
#         else:
#             return "N/A"
#     except Exception as e:
#         print(f"Error getting dollar change at exp!\n{e}\n")
#         return "Error!"
#     finally:
#         update_cache(contract, contract_symbol)

@xw.func(volatile=True)
def get_High_Post_Buy(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    try:
        if contract.high_Post_Buy in (None, 0):
            return "NO DATA"
        else:
            return contract.high_Post_Buy
    except Exception as e:
        print(f"Error getting high post buy!\n{e}\n")
        return "Error!"
    finally:
        update_cache(contract, contract_symbol)

@xw.func(volatile=True)
def get_High_Days_Out(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    try:
        return contract.high_Days_Out
    except Exception as e:
        print(f"Error getting high days out!\n{e}\n")
        return "Error!"
    finally:
        update_cache(contract, contract_symbol)

# @xw.func(volatile=True)
# def get_Percent_Change_High(caller):
#     contract_symbol = get_Contract_Symbol(caller)
#     contract = contract_init(contract_symbol)
#     try:
#         if contract.percent_Change_High is not None:
#             return contract.percent_Change_High
#         else:
#             return "Missing high data!"
#     except Exception as e:
#         print(f"Error getting percent change high!\n{e}\n")
#         return "Error!"
#     finally:
#         update_cache(contract, contract_symbol)

# @xw.func(volatile=True)
# def get_Dollar_Change_High(caller):
#     contract_symbol = get_Contract_Symbol(caller)
#     contract = contract_init(contract_symbol)
#     try:
#         if contract.dollar_Change_High is not None:
#             return contract.dollar_Change_High
#         else:
#             return "Missing high data!"
#     except Exception as e:
#         print(f"Error getting dollar change high!\n{e}\n")
#         return "Error!"
#     finally:
#         update_cache(contract, contract_symbol)

if __name__ == '__main__':
    xw.serve()