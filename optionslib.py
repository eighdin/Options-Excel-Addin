from datetime import *
import xlwings as xw
import yfinance as yf
import numpy as np
import pandas as pd
from diskcache import Cache
import ContractInfo

contract_Info_Dict = {}
# use vba to have current price and eow prices updated on a timer

def contract_init(
    contract_symbol: str
) -> ContractInfo:
    """
        Returns the contract object relating to the passed in contract symbol after creating it, if needed.
        Fills in the basic info that can be grabbed from the contract symbol.
    """
    if contract_Info_Dict.get(contract_symbol) is None:
        contract_Info_Dict[contract_symbol] = ContractInfo()
    contract = contract_Info_Dict[contract_symbol]
    if contract.ticker is None: # create ticker info if it doesn't exist
        contract.ticker = "".join(filter(lambda x: not x.isdigit(), contract_symbol[:6]))
    if contract.strike_Price is None: # create strike price info if it doesn't exist
        contract.strike_Price = float(contract_symbol[-8:])/1000
        contract.contract_Type = contract_symbol[-9:][:1]
    if contract.exp_Datetime_Obj is None: # create expiration date information if it doesn't exist
        exp_date_from_symbol = contract_symbol[contract.ticker.__len__():contract.ticker.__len__()+6]
        contract.exp_Datetime_Obj = datetime.strptime(exp_date_from_symbol, "%y%m%d")
        contract.exp_Date_String = contract.exp_Datetime_Obj.strftime("%Y-%m-%d")
    return contract

def get_Contract_Symbol(caller):
    sheet = caller.sheet
    row = caller.row
    value = sheet.cells(row, 1).value
    return value

@xw.func()
def set_Refresh_Rate_Mins(new_refresh_rate):
    global refresh_rate 
    if new_refresh_rate >= 60:
        refresh_rate = new_refresh_rate
        return f"Refresh rate set to: {new_refresh_rate} mins"
    else:
        return f"Refresh rate must be > 60 mins!"

# contract symbol should be first column always

@xw.func()
def set_Contract_Date(contract_date, caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    contract.contract_Date = contract_date
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
    contract.contract_Price = contract_price
    return contract.contract_Price

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
    contract.total = (contract.contract_Price*100)*contract.volume
    return contract.total

@xw.func()
def get_Current_Price(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    return contract.current_Price

@xw.func()
def get_Percent_Change(caller):
    contract_symbol = get_Contract_Symbol(caller)
    contract = contract_init(contract_symbol)
    return (contract.current_Price-contract.contract_Price)/contract.contract_Price

@xw.func()
def get_Dollar_Change(caller):
    contract_symbol = get_Contract_Symbol(caller)

@xw.func()
def get_Price_EOW(n_week, caller):
    contract_symbol = get_Contract_Symbol(caller)

@xw.func()
def get_Percent_EOW(n_week, caller):
    contract_symbol = get_Contract_Symbol(caller)

@xw.func()
def get_Percent_Change_Exp(caller):
    contract_symbol = get_Contract_Symbol(caller)

@xw.func()
def get_Dollar_Change_Exp(caller):
    contract_symbol = get_Contract_Symbol(caller)

@xw.func()
def get_High_Post_Buy(caller):
    contract_symbol = get_Contract_Symbol(caller)

@xw.func()
def get_High_Days_Out(caller):
    contract_symbol = get_Contract_Symbol(caller)

@xw.func()
def get_Percent_Change_High(caller):
    contract_symbol = get_Contract_Symbol(caller)

@xw.func()
def get_Dollar_Change_High(caller):
    contract_symbol = get_Contract_Symbol(caller)

if datetime.now()-
    