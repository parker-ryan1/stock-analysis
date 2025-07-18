import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta
import sys

def fetch_stock_data(ticker):
    """
    Fetch comprehensive stock data for a given ticker from Yahoo Finance
    """
    try:
        # Create ticker object
        stock = yf.Ticker(ticker)
        
        # Get stock info
        info = stock.info
        
        # Get historical data (5 years for more comprehensive data)
        end_date = datetime.now()
        start_date = end_date - timedelta(days=5*365)
        hist_data = stock.history(start=start_date, end=end_date, interval='1d')
        
        # Get quarterly historical data for financial modeling
        quarterly_data = stock.history(period='5y', interval='3mo')
        
        # Get financial data
        try:
            financials = stock.financials
            quarterly_financials = stock.quarterly_financials
            balance_sheet = stock.balance_sheet
            quarterly_balance_sheet = stock.quarterly_balance_sheet
            cashflow = stock.cashflow
            quarterly_cashflow = stock.quarterly_cashflow
        except:
            financials = pd.DataFrame()
            quarterly_financials = pd.DataFrame()
            balance_sheet = pd.DataFrame()
            quarterly_balance_sheet = pd.DataFrame()
            cashflow = pd.DataFrame()
            quarterly_cashflow = pd.DataFrame()
        
        # Get dividends and splits
        dividends = stock.dividends
        splits = stock.splits
        
        return {
            'info': info,
            'historical': hist_data,
            'quarterly_data': quarterly_data,
            'financials': financials,
            'quarterly_financials': quarterly_financials,
            'balance_sheet': balance_sheet,
            'quarterly_balance_sheet': quarterly_balance_sheet,
            'cashflow': cashflow,
            'quarterly_cashflow': quarterly_cashflow,
            'dividends': dividends,
            'splits': splits
        }
    
    except Exception as e:
        print(f"Error fetching data for {ticker}: {str(e)}")
        return None

def create_excel_file(ticker, data):
    """
    Create Excel file with comprehensive stock data in organized format
    """
    filename = f"{ticker}_stock_data.xlsx"
    
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        
        # Main Summary Sheet
        create_main_sheet(writer, ticker, data)
        
        # Historical Data Sheet with comprehensive data
        create_historical_sheet(writer, data)
        
        # Financial Model Sheet
        create_financial_model_sheet(writer, data)
        
        # Company Info Sheet
        if data['info']:
            info_df = pd.DataFrame(list(data['info'].items()), columns=['Metric', 'Value'])
            info_df.to_excel(writer, sheet_name='Company Info', index=False)
        
        # Quarterly Financials
        if not data['quarterly_financials'].empty:
            data['quarterly_financials'].to_excel(writer, sheet_name='Quarterly Financials')
        
        # Balance Sheet
        if not data['balance_sheet'].empty:
            data['balance_sheet'].to_excel(writer, sheet_name='Balance Sheet')
        
        # Cash Flow
        if not data['cashflow'].empty:
            data['cashflow'].to_excel(writer, sheet_name='Cash Flow')
        
        # Dividends & Splits
        create_dividends_splits_sheet(writer, data)
    
    print(f"Excel file created: {filename}")
    return filename

def create_main_sheet(writer, ticker, data):
    """Create main summary sheet with key metrics"""
    info = data['info']
    hist = data['historical']
    
    # Key metrics summary
    main_data = []
    
    if not hist.empty:
        current_price = hist['Close'].iloc[-1]
        main_data.extend([
            ['', '', '', '', '', '', '', '', '', '', '', '', 'Price', current_price, ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', 'Shares', info.get('sharesOutstanding', 'N/A'), 'Latest'],
            ['', '', '', '', '', '', '', '', '', '', '', '', 'Market Cap', info.get('marketCap', 'N/A'), ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', 'Cash', info.get('totalCash', 'N/A'), 'Latest'],
            ['', '', '', '', '', '', '', '', '', '', '', '', 'Debt', info.get('totalDebt', 'N/A'), 'Latest'],
            ['', '', '', '', '', '', '', '', '', '', '', '', 'Revenue', info.get('totalRevenue', 'N/A'), 'TTM'],
            ['', '', '', '', '', '', '', '', '', '', '', '', 'Net Income', info.get('netIncomeToCommon', 'N/A'), 'TTM'],
            ['', '', '', '', '', '', '', '', '', '', '', '', 'EPS', info.get('trailingEps', 'N/A'), 'TTM'],
            ['', '', '', '', '', '', '', '', '', '', '', '', 'P/E Ratio', info.get('trailingPE', 'N/A'), ''],
        ])
    
    main_df = pd.DataFrame(main_data)
    main_df.to_excel(writer, sheet_name='Main', index=False, header=False)

def create_historical_sheet(writer, data):
    """Create comprehensive historical data sheet"""
    hist = data['historical']
    
    if not hist.empty:
        # Reset index to make Date a column
        hist_clean = hist.reset_index()
        
        # Add calculated fields
        hist_clean['Daily Return'] = hist_clean['Close'].pct_change()
        hist_clean['20-Day MA'] = hist_clean['Close'].rolling(window=20).mean()
        hist_clean['50-Day MA'] = hist_clean['Close'].rolling(window=50).mean()
        hist_clean['200-Day MA'] = hist_clean['Close'].rolling(window=200).mean()
        hist_clean['Volatility'] = hist_clean['Daily Return'].rolling(window=20).std() * (252**0.5)
        
        # Format date column
        hist_clean['Date'] = hist_clean['Date'].dt.strftime('%Y-%m-%d')
        
        hist_clean.to_excel(writer, sheet_name='Historical Data', index=False)
    else:
        # Create empty sheet with headers
        empty_df = pd.DataFrame(columns=['Date', 'Open', 'High', 'Low', 'Close', 'Volume', 'Dividends', 'Stock Splits'])
        empty_df.to_excel(writer, sheet_name='Historical Data', index=False)

def create_financial_model_sheet(writer, data):
    """Create financial model sheet with quarterly data"""
    quarterly = data['quarterly_data']
    financials = data['quarterly_financials']
    
    model_data = []
    
    if not quarterly.empty:
        # Create date headers
        dates = quarterly.index.strftime('%Y-%m-%d').tolist()
        quarters = [f"Q{((pd.to_datetime(date).month-1)//3)+1}{pd.to_datetime(date).year%100}" for date in dates]
        
        # Add headers
        model_data.append(['Main'] + [''] + dates)
        model_data.append([''] + [''] + quarters)
        
        # Add price data
        model_data.append(['', 'Stock Price'] + quarterly['Close'].tolist())
        model_data.append(['', 'Volume'] + quarterly['Volume'].tolist())
        
        # Add financial data if available
        if not financials.empty:
            for metric in ['Total Revenue', 'Net Income', 'Total Assets', 'Total Debt']:
                if metric in financials.index:
                    values = financials.loc[metric].fillna('').tolist()
                    model_data.append(['', metric] + values)
    
    model_df = pd.DataFrame(model_data)
    model_df.to_excel(writer, sheet_name='Model', index=False, header=False)

def create_dividends_splits_sheet(writer, data):
    """Create dividends and splits sheet"""
    dividends = data['dividends']
    splits = data['splits']
    
    # Combine dividends and splits data
    div_split_data = []
    
    if not dividends.empty:
        div_df = dividends.reset_index()
        div_df['Type'] = 'Dividend'
        div_df = div_df.rename(columns={'Dividends': 'Amount'})
        div_split_data.append(div_df[['Date', 'Type', 'Amount']])
    
    if not splits.empty:
        split_df = splits.reset_index()
        split_df['Type'] = 'Stock Split'
        split_df = split_df.rename(columns={'Stock Splits': 'Amount'})
        div_split_data.append(split_df[['Date', 'Type', 'Amount']])
    
    if div_split_data:
        combined_df = pd.concat(div_split_data, ignore_index=True)
        combined_df = combined_df.sort_values('Date', ascending=False)
        combined_df['Date'] = combined_df['Date'].dt.strftime('%Y-%m-%d')
        combined_df.to_excel(writer, sheet_name='Dividends & Splits', index=False)
    else:
        # Create empty sheet
        empty_df = pd.DataFrame(columns=['Date', 'Type', 'Amount'])
        empty_df.to_excel(writer, sheet_name='Dividends & Splits', index=False)

def main():
    if len(sys.argv) != 2:
        print("Usage: python stock_data_fetcher.py <TICKER>")
        print("Example: python stock_data_fetcher.py AAPL")
        return
    
    ticker = sys.argv[1].upper()
    print(f"Fetching data for {ticker}...")
    
    # Fetch data
    data = fetch_stock_data(ticker)
    
    if data is None:
        print(f"Failed to fetch data for {ticker}")
        return
    
    # Create Excel file
    create_excel_file(ticker, data)
    print("Done!")

if __name__ == "__main__":
    main()