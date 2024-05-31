import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
import numpy as np

# Function to calculate Altman Z-Score
def calculate_altman_z_score(file_path, stock_price, fiscal_year_col):
    # Load the workbook and relevant sheets
    wb = openpyxl.load_workbook(file_path)
    sheet_names = wb.sheetnames

    balance_sheet = pd.read_excel(file_path, sheet_name='Balance Sheet')
    profit_loss = pd.read_excel(file_path, sheet_name='Profit Loss')
    per_share_stats = pd.read_excel(file_path, sheet_name='Per Share Statisticts')

    # Extracting data from Balance Sheet
    CA_cash = balance_sheet.loc[balance_sheet['Item'] == 'CA - Cash', fiscal_year_col].values[0]
    CA_receivables = balance_sheet.loc[balance_sheet['Item'] == 'CA - Receivables', fiscal_year_col].values[0]
    CA_prepaid_expenses = balance_sheet.loc[balance_sheet['Item'] == 'CA - Prepaid Expenses', fiscal_year_col].values[0]
    current_assets = CA_cash + CA_receivables + CA_prepaid_expenses

    total_assets = balance_sheet.loc[balance_sheet['Item'] == 'Total Assets', fiscal_year_col].values[0]
    current_liabilities = balance_sheet.loc[balance_sheet['Item'] == 'Total Curr. Liabilities', fiscal_year_col].values[0]
    total_liabilities = balance_sheet.loc[balance_sheet['Item'] == 'Total Liabilities', fiscal_year_col].values[0]

    # Extracting data from Profit Loss
    ebit = profit_loss.loc[profit_loss['Item'] == 'EBIT', fiscal_year_col].values[0]
    sales = profit_loss.loc[profit_loss['Item'] == 'Operating Revenue', fiscal_year_col].values[0]

    # Extracting data from Per Share Statistics
    shares_outstanding = per_share_stats.loc[per_share_stats['Item'] == 'Shares Outstand. (EOP)', fiscal_year_col].values[0]

    market_value_of_equity = shares_outstanding * stock_price

    # Calculating Working Capital (A)
    working_capital = current_assets - current_liabilities

    # Values
    A = working_capital
    B = total_assets
    C = balance_sheet.loc[balance_sheet['Item'] == 'Retained Earnings', fiscal_year_col].values[0]
    D = ebit
    E = market_value_of_equity
    F = total_liabilities
    G = sales
    print(f'Company: {file_path.split("/")[-1].split("-")[0]}')
    print(f'Working Capital: {working_capital}, Total Assets: {total_assets}, Retained Earnings: {C}, EBIT: {D}, Market Value of Equity: {E}, Total Liabilities: {F}, Sales: {G}, Total Assets: {B}')

    # Altman Z-Score calculation
    Z = 1.2 * (A / B) + 1.4 * (C / B) + 3.3 * (D / B) + 0.6 * (E / F) + 1.0 * (G / B)
    
    # Calculate Z-scores for different share prices
    stock_prices = [stock_price * (1 + i / 100) for i in range(-20, 21)]
    z_scores = []

    for price in stock_prices:
        market_value_of_equity = shares_outstanding * price
        E = market_value_of_equity
        Z = 1.2 * (A / B) + 1.4 * (C / B) + 3.3 * (D / B) + 0.6 * (E / F) + 1.0 * (G / B)
        z_scores.append(Z)
    
    return stock_prices, z_scores

# File paths, stock prices, and fiscal year columns for different stocks
files_and_prices = [
    ('WEB-2018-2023.xlsx', 8.56, '03/23'),
    ('FLT-2018-2023.xlsx', 20.52, '06/23'),
    ('CTD-2018-2023.xlsx', 14.94, '06/23'),
    ('HLO-2018-2023.xlsx', 2.34, '06/23'),
    ("SDR-2020-2023.xlsx", 5.41, '06/23')
]

# Plotting the Z-Scores for different stocks
plt.figure(figsize=(10, 6))

for file_path, stock_price, fiscal_year_col in files_and_prices:
    stock_prices, z_scores = calculate_altman_z_score(file_path, stock_price, fiscal_year_col)
    plt.plot(stock_prices, z_scores, label=f'{file_path.split("/")[-1].split("-")[0]}')
    plt.scatter(stock_price, z_scores[20], marker='o', label=f'{file_path.split("/")[-1].split("-")[0]} Current Price')
    # Add a label for the current stock price and Z-Score. Add a box around the text. And align right and above by 10px. Label SP: , Z-Score:
    plt.text(stock_price + 0.5, z_scores[20] + 1.3, f'Price: {stock_price}\nZ-Score: {z_scores[20]:.2f}', bbox=dict(facecolor='white', alpha=0.5, edgecolor='black'), ha='right', va='top', fontsize=8)

plt.xlabel('Stock Price ($)')
plt.ylabel('Altman Z-Score')
plt.title('Altman Z-Score vs. Stock Price')
#plt.grid()

# Highlight sections based on Altman Z-Score for financial distress (grey), neutral (yellow), and safe (green). Values based on Altman's original paper.
plt.axhspan(0, 1.88, color='red', alpha=0.2)
plt.axhspan(1.88, 3.0, color='grey', alpha=0.2)
plt.axhspan(3.0, 15, color='green', alpha=0.2)

# Restrict size of axis from 0-15
plt.ylim(0, 15)

plt.legend()
plt.show()