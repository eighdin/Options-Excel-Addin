class ContractInfo:
    def __init__(self):
        self.date = None
        self.datetime_Obj = None
        self.ticker = None
        self.strike_Price = None
        self.contract_Type = None
        self.exp_Date_From_Symbol = None
        self.exp_Datetime_Obj = None
        self.exp_Date_String = None
        self.orig_Price = None
        self.volume = None
        self.total = None
        self.current_Price = None
        self.percent_Change = None
        self.dollar_Change = None
        self.price_EOW_List = []
        self.percent_EOW_List = []
        self.percent_Change_Exp = None
        self.dollar_Change_Exp = None
        self.high_Post_Buy = None
        self.high_Days_Out = None
        self.percent_Change_High = None
        self.dollar_Change_High = None
        self.is_Expired = False