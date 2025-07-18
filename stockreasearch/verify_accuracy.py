import yfinance as yf
import pandas as pd

# Get live data
stock = yf.Ticker('AAPL')
info = stock.info

print("=== LIVE YAHOO FINANCE DATA ===")
print(f"Current Price: ${info.get('currentPrice', 'N/A')}")
print(f"Market Cap: ${info.get('marketCap', 'N/A'):,}")
print(f"P/E Ratio: {info.get('trailingPE', 'N/A')}")
print(f"EPS: ${info.get('trailingEps', 'N/A')}")
print(f"Shares Outstanding: {info.get('sharesOutstanding', 'N/A'):,}")

# Get data from our generated file
print("\n=== OUR PROGRAM OUTPUT ===")
df = pd.read_excel('AAPL_stock_data.xlsx', sheet_name='Main', header=None)
for i, row in df.iterrows():
    if pd.notna(row[12]):
        print(f"{row[12]}: {row[13]}")

# Check historical data accuracy
print("\n=== HISTORICAL DATA CHECK ===")
hist_df = pd.read_excel('AAPL_stock_data.xlsx', sheet_name='Historical Data')
print(f"Historical data points: {len(hist_df)}")
print(f"Date range: {hist_df['Date'].min()} to {hist_df['Date'].max()}")
print(f"Latest close price: ${hist_df['Close'].iloc[-1]:.2f}")

# Get live historical data for comparison
live_hist = stock.history(period='5d')
print(f"Live latest close: ${live_hist['Close'].iloc[-1]:.2f}")