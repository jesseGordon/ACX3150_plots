import pandas as pd
import openpyxl

# Load the dataset
file_path = 'WEB-2018-2023.xlsx'

stock_price = 8.56

# Load the workbook and check the sheet names
wb = openpyxl.load_workbook(file_path)
sheet_names = wb.sheetnames
print(sheet_names)

# Load the relevant sheets for the necessary calculations
balance_sheet = pd.read_excel(file_path, sheet_name='Balance Sheet')
profit_loss = pd.read_excel(file_path, sheet_name='Profit Loss')
per_share_stats = pd.read_excel(file_path, sheet_name='Per Share Statisticts')

# Extracting data from Balance Sheet
CA_cash = balance_sheet.loc[balance_sheet['Item'] == 'CA - Cash', '03/23'].values[0]
CA_receivables = balance_sheet.loc[balance_sheet['Item'] == 'CA - Receivables', '03/23'].values[0]
CA_prepaid_expenses = balance_sheet.loc[balance_sheet['Item'] == 'CA - Prepaid Expenses', '03/23'].values[0]
current_assets = CA_cash + CA_receivables + CA_prepaid_expenses

total_assets = balance_sheet.loc[balance_sheet['Item'] == 'Total Assets', '03/23'].values[0]
current_liabilities = balance_sheet.loc[balance_sheet['Item'] == 'Total Curr. Liabilities', '03/23'].values[0]
total_liabilities = balance_sheet.loc[balance_sheet['Item'] == 'Total Liabilities', '03/23'].values[0]

# Extracting data from Profit Loss
ebit = profit_loss.loc[profit_loss['Item'] == 'EBIT', '03/23'].values[0]
sales = profit_loss.loc[profit_loss['Item'] == 'Operating Revenue', '03/23'].values[0]

# Extracting data from Per Share Statistics
shares_outstanding = per_share_stats.loc[per_share_stats['Item'] == 'Shares Outstand. (EOP)', '03/23'].values[0]


market_value_of_equity = shares_outstanding * stock_price

# Calculating Working Capital (A)
working_capital = current_assets - current_liabilities

# Displaying all extracted and calculated values
print(f"Working Capital: {working_capital}")
print(f"Total Assets: {total_assets}")
print(f"Current Liabilities: {current_liabilities}")
print(f"Total Liabilities: {total_liabilities}")
print(f"EBIT: {ebit}")
print(f"Market Value of Equity: {market_value_of_equity}")
print(f"Sales: {sales}")

# Values
A = working_capital
B = total_assets
C = balance_sheet.loc[balance_sheet['Item'] == 'Retained Earnings', '03/23'].values[0]
D = ebit
E = market_value_of_equity
F = total_liabilities
G = sales
H = total_assets

# Altman Z-Score calculation
Z = 1.2 * (A / B) + 1.4 * (C / B) + 3.3 * (D / B) + 0.6 * (E / F) + 1.0 * (G / H)
print(f"Altman Z-Score: {Z}")


# Plot for different share prices (around the current stock price, += 20%)
stock_prices = [stock_price * (1 + i / 100) for i in range(-20, 21)]
z_scores = []

for price in stock_prices:
    market_value_of_equity = shares_outstanding * price
    E = market_value_of_equity
    Z = 1.2 * (A / B) + 1.4 * (C / B) + 3.3 * (D / B) + 0.6 * (E / F) + 1.0 * (G / H)
    z_scores.append(Z)

print(z_scores)

# Plotting the Z-Scores for different share prices
import matplotlib.pyplot as plt


#Plot current stock price
plt.axvline(x=stock_price, color='r', linestyle='--', label='Current Stock Price')

plt.plot(stock_prices, z_scores)
plt.xlabel('Stock Price ($)')
plt.ylabel('Altman Z-Score')
plt.title('Altman Z-Score vs. WEB Stock Price')
plt.grid()

import numpy as np
# Highlight sections based on Altman Z-Score for financial distress (grey), neutral (yellow), and safe (green). Values based on Altman's original paper.
plt.axhspan(0, 1.88, color='red', alpha=0.2)
plt.axhspan(1.88, 3.0, color='grey', alpha=0.2)
plt.axhspan(3.0, 6, color='green', alpha=0.2)



plt.show()