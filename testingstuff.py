# my_ticker = yf.Ticker(ticker="NVDA")
# option_frame = my_ticker.option_chain(my_ticker.options[0]).calls
# contract = option_frame[np.isclose(option_frame['strike'], 230)]
# lastPrice = contract['lastPrice']
# thisdate = datetime.strptime("10/1/24 10:34am", "%m/%d/%y %I:%M%p")
data = options_data("11/08/24 12:40pm", "TSLA", "$290C", "250221", "54.96", "39600")
print(data)
# hist = yf.download('TSLA250221C00290000', start=thisdate.date(), end=thisdate.date()+timedelta(days=100))
# print(hist['High'].sort_values(by='TSLA250221C00290000', ascending=False).iloc[0].name.date())
# for entry in hist['Close'].iloc:
#     #if entry.name.strftime("%Y-%m-%d").__contains__("2024-11-15"):
#         print(entry.iloc[0])

# option_symbol = "AAPL230120C00150000"  # Example option symbol (modify as needed)

# # Download historical data for the specific option
# historical_data = yf.download(option_symbol, start="2023-01-01", end="2023-12-31")

# print(historical_data)