import datetime

class ContractInfo:
    def __init__(self):
        self.date : str = None
        self.datetime_Obj : datetime = None
        self.ticker : str = None
        self.strike_Price : int = None
        self.contract_Type : str = None
        self.exp_Date_From_Symbol : str = None
        self.exp_Datetime_Obj : datetime = None
        self.exp_Date_String : str = None
        self.orig_Price : float = None
        self.volume : int = None
        self.total : float = None
        self.current_Price : float = None
        self.percent_Change : float = None
        self.dollar_Change : float = None
        self.price_EOW_List : list = []
        self.percent_EOW_List : list = []
        self.price_Exp : float = None
        self.percent_Change_Exp : float = None
        self.dollar_Change_Exp : float = None
        self.high_Post_Buy : float = None
        self.high_Days_Out : float= None
        self.percent_Change_High : float = None
        self.dollar_Change_High : float = None
        self.is_Expired : bool = False