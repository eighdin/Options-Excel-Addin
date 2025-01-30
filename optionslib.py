from datetime import *
import xlwings as xw
import yfinance as yf
import numpy as np
import pandas as pd
from diskcache import Cache
import ContractInfo

contract_Info_Dict = {}

# contract symbol should be first column always

# contract date is entered manually into spreadsheet

@xw.func()
def get_Ticker(caller):
    contract_symbol = caller.offset(0, -2).value
    if contract_Info_Dict.get(contract_symbol) is None:
        

@xw.func()
def get_Strike_Price(caller):
    contract_symbol = caller.offset(0, -3).value

@xw.func()
def get_Exp_Date(caller):
    contract_symbol = caller.offset(0, -4).value

# contract price entered manually into spreadsheet

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