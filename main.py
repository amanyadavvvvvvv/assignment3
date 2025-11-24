import yfinance as yf
import pandas as pd
import warnings

# Suppress warnings
warnings.filterwarnings("ignore")

def fetch_price_data(ticker, period="1y"):
    try:
        data = yf.download(ticker, period=period, progress=False)["Close"]
        
        if data.empty:
            print(f"  ✗ No price data available for {ticker}")
            return None
        
        return data
        
    except Exception as e:
        print(f"  ✗ Error fetching price data: {str(e)}")
        return None

def calculate_price_metrics(data):
    try:
        current_price = float(data.iloc[-1])
        high_52 = float(data.max())
        low_52 = float(data.min())
        distance_from_low = (current_price - low_52) / low_52 * 100
        
        return {
            'current_price': current_price,
            'high_52': high_52,
            'low_52': low_52,
            'distance_from_low': distance_from_low
        }
        
    except Exception as e:
        print(f"  ✗ Error calculating price metrics: {str(e)}")
        return None

def fetch_options_data(ticker, current_price):
    try:
        stock = yf.Ticker(ticker)
        expiries = stock.options
        
        # Check if options are available
        if not expiries:
            print(f"  ⚠ No options data available for {ticker}")
            return {
                'strike': None,
                'premium': None,
                'irr': None,
                'effective_return': None
            }
        
        # Get nearest expiry
        expiry = expiries[0]
        options = stock.option_chain(expiry)
        puts = options.puts
        
        if puts.empty:
            print(f"  ⚠ No put options available for {ticker}")
            return {
                'strike': None,
                'premium': None,
                'irr': None,
                'effective_return': None
            }
        
        # Find nearest strike to current price
        nearest = puts.iloc[(puts['strike'] - current_price).abs().argsort()[:1]]
        strike = float(nearest['strike'].values[0])
        premium = float(nearest['lastPrice'].values[0])
        
        # Calculate IRR and effective return
        irr = premium / strike
        effective_return = irr / 0.15
        
        return {
            'strike': strike,
            'premium': premium,
            'irr': irr,
            'effective_return': effective_return
        }
        
    except Exception as e:
        print(f"  ✗ Error fetching options data: {str(e)}")
        return {
            'strike': None,
            'premium': None,
            'irr': None,
            'effective_return': None
        }

def process_stock(ticker):
    print(f"\nProcessing {ticker} ...")
    
    # Fetch price data
    price_data = fetch_price_data(ticker)
    
    if price_data is None:
        # Return default values if price data fetch failed
        return {
            "Stock": ticker,
            "Current Price": 0,
            "52-Week High": 0,
            "52-Week Low": 0,
            "Distance from Low (%)": 0,
            "Nearest Strike": None,
            "Premium": None,
            "IRR": None,
            "Effective Return (15% margin)": None
        }
    
    # Calculate price metrics
    price_metrics = calculate_price_metrics(price_data)
    
    if price_metrics is None:
        # Return default values if calculation failed
        return {
            "Stock": ticker,
            "Current Price": 0,
            "52-Week High": 0,
            "52-Week Low": 0,
            "Distance from Low (%)": 0,
            "Nearest Strike": None,
            "Premium": None,
            "IRR": None,
            "Effective Return (15% margin)": None
        }
    
    # Fetch options data
    options_data = fetch_options_data(ticker, price_metrics['current_price'])
    
    # Compile results
    result = {
        "Stock": ticker,
        "Current Price": round(price_metrics['current_price'], 2),
        "52-Week High": round(price_metrics['high_52'], 2),
        "52-Week Low": round(price_metrics['low_52'], 2),
        "Distance from Low (%)": round(price_metrics['distance_from_low'], 2),
        "Nearest Strike": options_data['strike'],
        "Premium": options_data['premium'],
        "IRR": None if options_data['irr'] is None else round(options_data['irr'], 4),
        "Effective Return (15% margin)": None if options_data['effective_return'] is None else round(options_data['effective_return'], 4)
    }
    
    return result

def save_to_excel(df, filename="Assessment3_Output.xlsx"):
    try:
        df.to_excel(filename, index=False)
        print(f"\n✓ Data saved to {filename}")
        return True
        
    except Exception as e:
        print(f"\n✗ Error saving to Excel: {str(e)}")
        return False

def main():
    """Main execution function"""
    try:
        # List of stocks
        tickers = ["AAPL", "MSFT", "GOOG"]
        
        print("="*60)
        print("  Options Analysis - Stock Price & Put Options Data")
        print("="*60)
        
        # Process each stock
        all_results = []
        for ticker in tickers:
            result = process_stock(ticker)
            all_results.append(result)
        
        # Convert to DataFrame
        df = pd.DataFrame(all_results)
        
        # Display results
        print("\n" + "="*60)
        print("  Results Summary")
        print("="*60)
        print(df)
        
        # Save to Excel
        save_to_excel(df)
        
        print("\n" + "="*60)
        print("  Analysis Complete!")
        print("="*60)
        
    except KeyboardInterrupt:
        print("\n\nProcess interrupted by user.")
    except Exception as e:
        print(f"\nUnexpected error in main execution: {str(e)}")

if __name__ == "__main__":
    main()