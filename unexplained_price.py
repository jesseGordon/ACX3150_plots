import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pandas as pd
import numpy as np
import yfinance as yf

# Download stock and index prices
webjet_prices = yf.download('WEB.AX', start="2023-03-24", end="2024-03-24")['Close'] # Webjet stock prices
index_prices = yf.download('^AORD', start="2023-03-24", end="2024-03-24")['Close'] # ASX 200 index prices

# Calculate daily returns
webjet_returns = webjet_prices.pct_change()
index_returns = index_prices.pct_change()

# Webjet's beta
beta_webjet = 1.87

window = 10

# Calculate expected returns for Webjet based on the market returns
expected_webjet_returns = index_returns * beta_webjet

# Determine the unexplained effect
unexplained_effect = webjet_returns - expected_webjet_returns

# Apply a 5-day rolling mean to the actual returns and the unexplained effect
rolling_actual_returns = webjet_returns.rolling(window).mean()
rolling_unexplained_effect = unexplained_effect.rolling(window).mean()

# Now plot these rolling values
fig, ax = plt.subplots(figsize=(15, 7))


# Rolling unexplained effect
ax.plot(rolling_unexplained_effect.index, rolling_unexplained_effect * 100, label=f'{window}-Day Rolling Unexplained Effect (%)', color='green')


# Event data
# Replace 'path_to_excel_file.xlsx' with the actual path to your Excel file
file_path = 'events.xlsx'

# Read the Excel file
events = pd.read_excel(file_path)

# remove any rows with missing values
events = events.dropna()

print(events)

webjet_base = webjet_prices.iloc[0]

"""
# Plotting the stock prices
ax.plot(dates, webjet_prices, label='Webjet Stock Price', color='blue')
ax.plot(dates, index_prices, label='Index Price', color='green', linestyle='--')
ax.plot(dates, another_stock_prices, label='Another Stock Price', color='orange', linestyle='-.')

"""

# Define a color map for the impact direction
impact_colors = {'Stock': 'red', 'Industry': 'orange', 'Market': 'blue'}

def find_closest_business_day(target_date, date_range):
    while target_date not in date_range:
        target_date -= pd.Timedelta(days=1)
    return target_date

# Adding event annotations with impact direction
for i, event in events.iterrows():
    if event['Impact'] == 'Stock':
        # Find the normalized price for the event date
        web_unexplained = rolling_unexplained_effect.loc[find_closest_business_day(event['Date'], rolling_unexplained_effect.index)]
        date_event = find_closest_business_day(event['Date'], rolling_unexplained_effect.index)
        if not np.isnan(web_unexplained):
            # Prepare the text with the event and its impact
            annotation_text = f"{event['Event']}\nImpact: {event['Impact']}"
            # Choose the color based on the impact direction
            annotation_color = impact_colors.get(event['Impact'], 'black')
            
            # Calculate the offset for the text to minimize overlap
            text_offset = (0, 10 + i * 5) if i % 3 == 0 else (0, -10 - i * 5)
            
            # Add the annotation with improved formatting
            ax.annotate(annotation_text, 
                    xy=(date_event, web_unexplained * 100), 
                    xytext=text_offset, 
                    textcoords='offset points', 
                    arrowprops=dict(arrowstyle='->', color=annotation_color),
                    ha='center', 
                    va='center',
                    bbox=dict(boxstyle='square,pad=0.5', fc='white', alpha=0.9),  # square padding
                    fontsize=8,
                    color=annotation_color)
# Formatting the date axis
ax.xaxis.set_major_locator(mdates.MonthLocator())
ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %Y'))
ax.xaxis.set_minor_locator(mdates.WeekdayLocator(byweekday=mdates.MO))

# Improve the plot with grid, labels, title, and legend
ax.grid(True, which='major', linestyle='--', linewidth='0.5', color='black')
ax.set_title(f'52 Week Unexplained WEB.AX Stock Returns Chart with Event Annotations: Beta={beta_webjet}, Window={window} days')
ax.set_xlabel('Date')
ax.set_ylabel('Unexplained Returns (%)')
ax.legend()


# Rotate date labels for better readability
fig.autofmt_xdate()

plt.tight_layout()
plt.show()