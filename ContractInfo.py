import datetime

class ContractInfo:
    def __init__(self):
        # Date stuff for contract
        self.date : str = None
        self.datetime_Obj : datetime = None
        
        # general info
        self.ticker : str = None
        self.strike_Price : int = 0
        self.contract_Type : str = None
        
        # exp date info
        self.exp_Date_From_Symbol : str = None
        self.exp_Datetime_Obj : datetime = None
        self.exp_Date_String : str = None
        
        #exp price info
        self.price_Exp : float = None
        self.percent_Change_Exp : float = None
        self.dollar_Change_Exp : float = None
        
        # basic price info
        self.orig_Price : float = 0
        self.volume : int = 0
        self.total : float = 0
        self.current_Price : float = 0
        self.percent_Change : float = 0
        self.dollar_Change : float = 0
        
        # EOW lists
        self.price_EOW_List : list = []
        self.percent_EOW_List : list = []
        
        # high price values
        self.high_Day_Last_Refreshed : datetime = None
        self.high_Post_Buy : float = 0
        self.high_Days_Out : float= None
        self.percent_Change_High : float = None
        self.dollar_Change_High : float = None
        
        # duh
        self.is_Expired : bool = False