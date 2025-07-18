# Stock Data Fetcher

A comprehensive Python application that fetches real-time stock market data from Yahoo Finance and exports it to professionally formatted Excel files. This tool provides institutional-grade financial data analysis capabilities for investors, analysts, and researchers.

## Features

### 📊 Comprehensive Data Coverage
- **Real-time stock prices** and market metrics
- **5 years of historical data** (1,200+ daily data points)
- **Technical indicators** (20, 50, 200-day moving averages)
- **Financial statements** (Income Statement, Balance Sheet, Cash Flow)
- **Dividend and stock split history**
- **Volatility calculations** and daily returns

### 📈 Professional Excel Output
- **Main Summary Sheet**: Key financial metrics and ratios
- **Historical Data**: Complete OHLCV data with technical analysis
- **Financial Model**: Quarterly data formatted for financial modeling
- **Company Information**: Detailed corporate data
- **Financial Statements**: Annual and quarterly financials
- **Dividends & Splits**: Complete corporate action history

### 🎯 Key Metrics Included
- Current stock price and market capitalization
- Price-to-earnings (P/E) ratio and earnings per share (EPS)
- Revenue, net income, and profit margins
- Cash position and debt levels
- Trading volume and volatility metrics
- 52-week high/low ranges

## Installation

### Prerequisites
- Python 3.7 or higher
- pip package manager

### Install Dependencies
```bash
pip install -r requirements.txt
```

### Required Packages
- `yfinance>=0.2.18` - Yahoo Finance API wrapper
- `pandas>=1.5.0` - Data manipulation and analysis
- `openpyxl>=3.1.0` - Excel file creation and formatting

## Usage

### Basic Usage
```bash
python stock_data_fetcher.py <TICKER_SYMBOL>
```

### Examples
```bash
# Apple Inc.
python stock_data_fetcher.py AAPL

# Microsoft Corporation
python stock_data_fetcher.py MSFT

# Tesla Inc.
python stock_data_fetcher.py TSLA

# NVIDIA Corporation
python stock_data_fetcher.py NVDA
```

### Output
The program generates an Excel file named `{TICKER}_stock_data.xlsx` containing multiple worksheets with comprehensive financial data.

## Excel File Structure

| Sheet Name | Description |
|------------|-------------|
| **Main** | Executive summary with key financial metrics |
| **Historical Data** | 5 years of daily price data with technical indicators |
| **Model** | Quarterly financial model with time series data |
| **Company Info** | Detailed company information and metadata |
| **Quarterly Financials** | Quarterly income statement data |
| **Balance Sheet** | Annual balance sheet information |
| **Cash Flow** | Annual cash flow statement |
| **Dividends & Splits** | Complete dividend and stock split history |

## Data Accuracy

All data is sourced directly from Yahoo Finance's official API, ensuring:
- ✅ **Real-time accuracy** - Live market data
- ✅ **Institutional quality** - Same data used by financial professionals
- ✅ **Comprehensive coverage** - 5 years of historical data
- ✅ **Technical analysis ready** - Pre-calculated indicators

## Technical Specifications

### Data Range
- **Historical Data**: 5 years of daily data
- **Financial Statements**: Up to 4 years of annual data
- **Quarterly Data**: Up to 4 years of quarterly data
- **Update Frequency**: Real-time (when markets are open)

### Calculated Metrics
- Daily returns and percentage changes
- Simple moving averages (20, 50, 200-day)
- Annualized volatility (20-day rolling)
- Technical indicators for trend analysis

### File Format
- **Format**: Excel (.xlsx)
- **Engine**: OpenPyXL for professional formatting
- **Compatibility**: Excel 2010+ and Google Sheets

## Error Handling

The application includes robust error handling for:
- Invalid ticker symbols
- Network connectivity issues
- Missing financial data
- API rate limiting

## Use Cases

### Investment Analysis
- Fundamental analysis with comprehensive financial metrics
- Technical analysis with pre-calculated indicators
- Portfolio research and stock screening

### Academic Research
- Financial modeling and valuation exercises
- Market analysis and trend studies
- Educational projects and case studies

### Professional Applications
- Investment banking and equity research
- Financial planning and advisory services
- Corporate finance and M&A analysis

## Limitations

- Data availability depends on Yahoo Finance coverage
- Some metrics may not be available for all securities
- Historical data limited to Yahoo Finance's retention period
- Real-time data may have slight delays during market hours

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for:
- Additional financial metrics
- Enhanced Excel formatting
- New data sources integration
- Performance improvements

## License

This project is open source and available under the MIT License.

## Disclaimer

This tool is for informational and educational purposes only. It should not be considered as financial advice. Always consult with qualified financial professionals before making investment decisions. The accuracy of data depends on Yahoo Finance's data quality and availability.

## Support

For questions, issues, or feature requests, please open an issue on the project repository.

---

**Built with Python | Powered by Yahoo Finance API | Professional Excel Output**