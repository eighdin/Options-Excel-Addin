from datetime import *
import xlwings as xw
import yfinance as yf
import numpy as np
import pandas as pd
from diskcache import Cache
import ContractInfo

contract_Info_Dict = {}

def contract_check(
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
    if contract.exp_Datetime_Obj is None: # create expiration date information if it doesn't exist
        exp_date_from_symbol = contract_symbol[contract.ticker.__len__():contract.ticker.__len__()+6]
        contract.exp_Datetime_Obj = datetime.strptime(exp_date_from_symbol, "%y%m%d")
        contract.exp_Date_String = contract.exp_Datetime_Obj.strftime("%Y-%m-%d")
    return contract

# contract symbol should be first column always

# contract date is entered manually into spreadsheet

@xw.func()
def get_Ticker(caller):
    contract_symbol = caller.offset(0, -2).value
    contract = contract_check(contract_symbol)
    return contract.ticker


@xw.func()
def get_Strike_Price(caller):
    contract_symbol = caller.offset(0, -3).value
    contract = contract_check(contract_symbol)
    return contract.strike_Price


@xw.func()
def get_Exp_Date(caller):
    contract_symbol = caller.offset(0, -4).value
    contract = contract_check(contract_symbol)
    return contract.exp_Date_String

def set_Contract_Price(contract_price, caller):
    contract_symbol = caller.offset(0, -4).value
    contract = contract_check(contract_symbol)
    contract.contract_Price = contract_price
    return contract.contract_Price

# volume entered manually into spreadsheet

@xw.func()
def get_Total(caller):
    pass

@xw.func()
def get_Current_Price(caller):
    contract_symbol = caller.offset(0, -8).value

@xw.func()
def get_Percent_Change(caller):
    pass

@xw.func()
def get_Dollar_Change(caller):
    pass

@xw.func()
def get_Price_EOW(caller):
    pass

@xw.func()
def get_Percent_EOW(caller):
    pass

@xw.func()
def get_Percent_Change_Exp(caller):
    pass

@xw.func()
def get_Dollar_Change_Exp(caller):
    pass

@xw.func()
def get_High_Post_Buy(caller):
    pass

@xw.func()
def get_High_Days_Out(caller):
    pass

@xw.func()
def get_Percent_Change_High(caller):
    pass

@xw.func()
def get_Dollar_Change_High(caller):
    pass