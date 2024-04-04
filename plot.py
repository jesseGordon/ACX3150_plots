import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pandas as pd
import numpy as np
import yfinance as yf



webjet_prices = yf.download('WEB.AX', start="2023-03-24", end="2024-03-24")['Close'] # Webjet stock prices

index_prices = yf.download('^AORD', start="2023-03-24", end="2024-03-24")['Close'] # ASX 200 index prices

another_stock_prices = yf.download('FLT.AX', start="2023-03-24", end="2024-03-24")['Close'] # Another stock prices

dates = webjet_prices.index

print(len(dates), len(webjet_prices), len(index_prices), len(another_stock_prices))



# Event data
file_path = 'events.xlsx'

events = pd.read_excel(file_path)

# remove any rows with missing values
events = events.dropna()

print(events)

# Convert dates to pandas datetime format
events['Date'] = pd.to_datetime(events['Date'])
# finds the closest previous business day for a given date that is not in our date range
def find_closest_business_day(target_date, date_range):
    while target_date not in date_range:
        target_date -= pd.Timedelta(days=1)
    return target_date


events['Date'] = events['Date'].apply(find_closest_business_day, date_range=dates)


# Choose a base date, which is common to all data
base_date = '2023-03-24'  

# Normalize the prices to start from the same point
webjet_base = webjet_prices.loc[base_date]
index_base = index_prices.loc[base_date]
flt_base = another_stock_prices.loc[base_date]

webjet_normalized = (webjet_prices / webjet_base) * 100
index_normalized = (index_prices / index_base) * 100
flt_normalized = (another_stock_prices / flt_base) * 100

# plot these normalised values
fig, ax = plt.subplots(figsize=(15, 7))

ax.plot(dates, webjet_normalized, label='Webjet - WEB.AX', color='red')
ax.plot(dates, index_normalized, label='Index - ^AORD', color='blue', linestyle='--')
ax.plot(dates, flt_normalized, label='Flight Centre - FLT.AX', color='orange', linestyle='-.')


# Define a color map 
impact_colors = {'Stock': 'red', 'Industry': 'orange', 'Market': 'blue'}

# Adding event annotations with impact direction
for i, event in events.iterrows():

    event_price = webjet_prices.get(event['Date'], np.nan)
    event_normalized_price = (event_price / webjet_base) * 100 if not np.isnan(event_price) else np.nan
    if not np.isnan(event_normalized_price):
        annotation_text = f"{event['Event']}\nImpact: {event['Impact']}"
        # Choose the colour based on  impact direction
        annotation_color = impact_colors.get(event['Impact'], 'black')
        
        text_offset = (0, int(event['Offset'])) if 'Offset' in event else (0, 0)
        
        # Add the annotation with improved formatting
        ax.annotate(annotation_text, 
                xy=(event['Date'], event_normalized_price), 
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

ax.grid(True, which='major', linestyle='--', linewidth='0.5', color='black')
ax.set_title('52 Week Stock Price Chart with Event Annotations')
ax.set_xlabel('Date')
ax.set_ylabel('Stock Price (%)')
ax.legend()



fig.autofmt_xdate()

plt.tight_layout()
plt.show()