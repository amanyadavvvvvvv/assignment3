# Options Analyzer

A Python script that analyzes stock prices and put options data to calculate IRR and effective returns.

## Features

- Fetches 1-year historical price data
- Calculates 52-week high/low and distance from low
- Retrieves nearest strike put options
- Calculates IRR and effective return (15% margin)
- Exports results to Excel

## Requirements

```bash
pip install -r requirements.txt
```

**Dependencies:**
- pandas
- yfinance
- openpyxl

## Usage

```bash
python main.py
```

## Output

The script generates:
- Console output with analysis summary
- Excel file: `Assessment3_Output.xlsx`

## Stocks Analyzed

- AAPL (Apple Inc.)
- MSFT (Microsoft Corporation)
- GOOG (Alphabet Inc.)

## Metrics Calculated

- Current Price
- 52-Week High/Low
- Distance from Low (%)
- Nearest Strike Price
- Premium
- IRR (Internal Rate of Return)
- Effective Return (with 15% margin)

## Customization

Edit the `tickers` list in the `main()` function:

```python
tickers = ["AAPL", "TSLA", "NVDA"]
```

## Notes

- Requires active internet connection
- Data source: Yahoo Finance
- Uses nearest expiry options
- US stocks used due to reliable options data availability
- Not all stocks have options data available