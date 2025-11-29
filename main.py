import yfinance as yf
import pandas as pd
import warnings
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
from datetime import datetime

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
            "Effective Return (15% margin)": None,
            "price_data": None
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
            "Effective Return (15% margin)": None,
            "price_data": None
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
        "Effective Return (15% margin)": None if options_data['effective_return'] is None else round(options_data['effective_return'], 4),
        "price_data": price_data  # Store for graphing
    }
    
    return result

def create_visualizations(df, all_results):
    """Create comprehensive visualization graphs"""
    try:
        # Set style
        plt.style.use('seaborn-v0_8-darkgrid')
        
        # Create figure with subplots
        fig = plt.figure(figsize=(20, 12))
        gs = gridspec.GridSpec(3, 3, figure=fig, hspace=0.3, wspace=0.3)
        
        # Color scheme
        colors = ['#2E86AB', '#A23B72', '#F18F01']
        
        # 1. Price Trends (Top row - spans 3 columns)
        ax1 = fig.add_subplot(gs[0, :])
        for idx, result in enumerate(all_results):
            if result['price_data'] is not None:
                ticker = result['Stock']
                price_data = result['price_data']
                ax1.plot(price_data.index, price_data.values, 
                        label=ticker, linewidth=2, color=colors[idx % 3])
                
                # Mark current price
                ax1.scatter(price_data.index[-1], price_data.values[-1], 
                           color=colors[idx % 3], s=100, zorder=5, 
                           edgecolors='white', linewidth=2)
        
        ax1.set_title('52-Week Price Trends', fontsize=16, fontweight='bold', pad=20)
        ax1.set_xlabel('Date', fontsize=12)
        ax1.set_ylabel('Price ($)', fontsize=12)
        ax1.legend(loc='best', fontsize=11)
        ax1.grid(True, alpha=0.3)
        
        # 2. Current vs High/Low (Middle left)
        ax2 = fig.add_subplot(gs[1, 0])
        stocks = df['Stock'].tolist()
        current = df['Current Price'].tolist()
        high = df['52-Week High'].tolist()
        low = df['52-Week Low'].tolist()
        
        x = range(len(stocks))
        width = 0.25
        
        ax2.bar([i - width for i in x], low, width, label='52W Low', 
               color='#E63946', alpha=0.8)
        ax2.bar(x, current, width, label='Current', 
               color='#2E86AB', alpha=0.8)
        ax2.bar([i + width for i in x], high, width, label='52W High', 
               color='#06A77D', alpha=0.8)
        
        ax2.set_title('Price Comparison', fontsize=14, fontweight='bold')
        ax2.set_ylabel('Price ($)', fontsize=11)
        ax2.set_xticks(x)
        ax2.set_xticklabels(stocks, fontsize=11)
        ax2.legend(fontsize=10)
        ax2.grid(True, alpha=0.3, axis='y')
        
        # 3. Distance from Low (Middle center)
        ax3 = fig.add_subplot(gs[1, 1])
        distances = df['Distance from Low (%)'].tolist()
        bars = ax3.barh(stocks, distances, color=colors, alpha=0.8)
        
        # Add value labels
        for i, (bar, val) in enumerate(zip(bars, distances)):
            ax3.text(val + 1, i, f'{val:.1f}%', 
                    va='center', fontsize=10, fontweight='bold')
        
        ax3.set_title('Distance from 52-Week Low', fontsize=14, fontweight='bold')
        ax3.set_xlabel('Distance (%)', fontsize=11)
        ax3.grid(True, alpha=0.3, axis='x')
        
        # 4. Options Premium (Middle right)
        ax4 = fig.add_subplot(gs[1, 2])
        premiums = [p if p is not None else 0 for p in df['Premium'].tolist()]
        bars = ax4.bar(stocks, premiums, color=colors, alpha=0.8)
        
        # Add value labels
        for bar, val in zip(bars, df['Premium'].tolist()):
            height = bar.get_height()
            if val is not None:
                ax4.text(bar.get_x() + bar.get_width()/2., height,
                        f'${val:.2f}', ha='center', va='bottom', 
                        fontsize=10, fontweight='bold')
        
        ax4.set_title('Put Option Premium', fontsize=14, fontweight='bold')
        ax4.set_ylabel('Premium ($)', fontsize=11)
        ax4.grid(True, alpha=0.3, axis='y')
        
        # 5. IRR Comparison (Bottom left)
        ax5 = fig.add_subplot(gs[2, 0])
        irrs = [i * 100 if i is not None else 0 for i in df['IRR'].tolist()]
        bars = ax5.bar(stocks, irrs, color=colors, alpha=0.8)
        
        # Add value labels
        for bar, val in zip(bars, df['IRR'].tolist()):
            height = bar.get_height()
            if val is not None:
                ax5.text(bar.get_x() + bar.get_width()/2., height,
                        f'{val*100:.2f}%', ha='center', va='bottom', 
                        fontsize=10, fontweight='bold')
        
        ax5.set_title('Internal Rate of Return (IRR)', fontsize=14, fontweight='bold')
        ax5.set_ylabel('IRR (%)', fontsize=11)
        ax5.grid(True, alpha=0.3, axis='y')
        
        # 6. Effective Return (Bottom center)
        ax6 = fig.add_subplot(gs[2, 1])
        eff_returns = [e * 100 if e is not None else 0 for e in df['Effective Return (15% margin)'].tolist()]
        bars = ax6.bar(stocks, eff_returns, color=colors, alpha=0.8)
        
        # Add value labels
        for bar, val in zip(bars, df['Effective Return (15% margin)'].tolist()):
            height = bar.get_height()
            if val is not None:
                ax6.text(bar.get_x() + bar.get_width()/2., height,
                        f'{val*100:.2f}%', ha='center', va='bottom', 
                        fontsize=10, fontweight='bold')
        
        ax6.set_title('Effective Return (15% Margin)', fontsize=14, fontweight='bold')
        ax6.set_ylabel('Return (%)', fontsize=11)
        ax6.grid(True, alpha=0.3, axis='y')
        
        # 7. Strike vs Current Price (Bottom right)
        ax7 = fig.add_subplot(gs[2, 2])
        strikes = [s if s is not None else 0 for s in df['Nearest Strike'].tolist()]
        
        x = range(len(stocks))
        width = 0.35
        
        ax7.bar([i - width/2 for i in x], current, width, 
               label='Current Price', color='#2E86AB', alpha=0.8)
        ax7.bar([i + width/2 for i in x], strikes, width, 
               label='Strike Price', color='#F18F01', alpha=0.8)
        
        ax7.set_title('Current Price vs Strike Price', fontsize=14, fontweight='bold')
        ax7.set_ylabel('Price ($)', fontsize=11)
        ax7.set_xticks(x)
        ax7.set_xticklabels(stocks, fontsize=11)
        ax7.legend(fontsize=10)
        ax7.grid(True, alpha=0.3, axis='y')
        
        # Add main title
        fig.suptitle('Options Analysis Dashboard', 
                    fontsize=20, fontweight='bold', y=0.995)
        
        # Save the figure
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Options_Analysis_Graphs_{timestamp}.png"
        plt.savefig(filename, dpi=300, bbox_inches='tight', facecolor='white')
        print(f"\n✓ Graphs saved to {filename}")
        
        # Also save as PDF for high quality
        pdf_filename = f"Options_Analysis_Graphs_{timestamp}.pdf"
        plt.savefig(pdf_filename, bbox_inches='tight', facecolor='white')
        print(f"✓ PDF graphs saved to {pdf_filename}")
        
        plt.close()
        return True
        
    except Exception as e:
        print(f"\n✗ Error creating visualizations: {str(e)}")
        return False

def save_to_excel(df, filename="Assessment3_Output.xlsx"):
    try:
        # Remove price_data column before saving to Excel
        df_to_save = df.drop(columns=['price_data'], errors='ignore')
        df_to_save.to_excel(filename, index=False)
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
        # Display without price_data column 
        df_display = df.drop(columns=['price_data'], errors='ignore')
        print(df_display)
        
        # Save to Excel
        save_to_excel(df)
        
        # Create visualizations
        print("\n" + "="*60)
        print("  Generating Visualizations...")
        print("="*60)
        create_visualizations(df, all_results)
        
        print("\n" + "="*60)
        print("  Analysis Complete!")
        print("="*60)
        
    except KeyboardInterrupt:
        print("\n\nProcess interrupted by user.")
    except Exception as e:
        print(f"\nUnexpected error in main execution: {str(e)}")

if __name__ == "__main__":
    main()