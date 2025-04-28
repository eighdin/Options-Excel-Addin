from dataclasses import dataclass, field
import datetime
from typing import List, Optional

EOW_LIST_SIZE = 5 

# anything commented out is no longer needed
@dataclass
class ContractInfo():
    # Date stuff for contract
    date : Optional[str] = None
    datetime_Obj : Optional[datetime.datetime] = None
    
    # general info
    ticker : Optional[str] = None
    strike_Price : Optional[int] = None
    contract_Type : Optional[str] = None
    
    # exp date info
    exp_Date_From_Symbol : Optional[str] = None
    exp_Datetime_Obj : Optional[datetime.datetime] = None
    exp_Date_String : Optional[str] = None
    
    #exp price info
    price_Exp : Optional[float] = None
    # percent_Change_Exp : Optional[float] = None
    # dollar_Change_Exp : Optional[float] = None
    
    # basic price info
    orig_Price : Optional[float] = None
    # volume : Optional[int] = None
    # total : Optional[float] = None
    current_Price : Optional[float] = None
    # percent_Change : Optional[float] = None
    # dollar_Change : Optional[float] = None
    
    # EOW lists
    price_EOW_List : List[float] = field(default_factory=lambda: [0.0]*EOW_LIST_SIZE)
    percent_EOW_List : List[float] = field(default_factory=lambda: [0.0]*EOW_LIST_SIZE)
    
    # high price values
    high_Day_Last_Refreshed : Optional[datetime.datetime] = None
    high_Post_Buy : Optional[float] = None
    high_Days_Out : Optional[float]= None
    # percent_Change_High : Optional[float] = None
    # dollar_Change_High : Optional[float] = None
    
    # duh
    is_Expired : bool = False