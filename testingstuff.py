from datetime import *
import yfinance as yf
import addon

# my_ticker = yf.Ticker(ticker="NVDA")
# option_frame = my_ticker.option_chain(my_ticker.options[0]).calls
# contract = option_frame[np.isclose(option_frame['strike'], 230)]
# lastPrice = contract['lastPrice']
#thisdate = datetime.strptime("10/1/24 10:34am", "%m/%d/%y %I:%M%p")
# data = options_data("11/08/24 12:40pm", "TSLA250221C00290000", "54.96", "39600")
# print(data)
# hist = yf.download('TSLA250221C00290000', start=thisdate.date(), end=thisdate.date()+timedelta(days=100))
# print(hist['High'].sort_values(by='TSLA250221C00290000', ascending=False).iloc[0].name.date())
# for entry in hist['Close'].iloc:
#     #if entry.name.strftime("%Y-%m-%d").__contains__("2024-11-15"):
#         print(entry.iloc[0])

# option_symbol = "AAPL230120C00150000"  # Example option symbol (modify as needed)

# # Download historical data for the specific option

# print(historical_data)
#contract_symbol = 'TSLA250221C00290000'
#strike_price = float(contract_symbol[-8:])/1000
#ticker = "".join(filter(lambda x: not x.isdigit(), contract_symbol[:6]))
#option_type = contract_symbol[-9:][:1]
#exp_date = datetime.strptime(contract_symbol[ticker.__len__():ticker.__len__()+6], "%y%m%d")
#print(ticker, strike_price, option_type, exp_date)
#historical_data = yf.download(contract_symbol)
# print(historical_data.iloc[historical_data.__len__()-1])
#print(yf.Ticker(ticker).option_chain(exp_date.date().__str__()))

print(addon.options_data("11/08/24 01:51pm","ALAB241220P00090000","3.94","2990", "ticker"))