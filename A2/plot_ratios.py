import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os

INDIVIDUAL = True
# Company details
companies = {
    'Webjet (WEB)': 'WEB-2018-2023.xlsx',
    'HelloWorld (HLO)': 'HLO-2018-2023.xlsx',
    'Flight Centre (FLT)': 'FLT-2018-2023.xlsx',
    'Corporate Travel (CTD)': 'CTD-2018-2023.xlsx',
    'Experience Co (EXP)': 'EXP-2018-2023.xlsx'
}

negative_net_profit_margin = {}


# Load and clean the data
data_clean = {}
for ticker, file in companies.items():
    data = pd.read_excel(file, sheet_name='Ratio Analysis', engine='openpyxl')
    data_clean[ticker] = data.replace('--', pd.NA)

# Define years for plotting
years_common = ['06/18', '06/19', '06/20', '03/21', '03/22', '03/23']
years_common = ['2018', '2019', '2020', '2021', '2022', '2023']

# Function to filter and clean ratio data for common years
def filter_clean_ratio_data_common(data, ratio_name):
    ratio_data = data[data['Item'] == ratio_name]
    if not ratio_data.empty:
        ratio_data = ratio_data.iloc[0, 3:3 + len(years_common)].apply(pd.to_numeric, errors='coerce').fillna(np.nan)
    else:
        ratio_data = pd.Series([np.nan] * len(years_common), index=years_common)
    return ratio_data

# Define the ratios and titles for the plots
ratios = {
    'Net Profit Margin (%)': 'Net Profit Margin (%) (Profitability Analysis)',
    #'EBIT Margin (%)': 'EBIT Margin (%) (Profitability Analysis)',
    #'EBITA Margin (%)': 'EBITA Margin (%) (Profitability Analysis)',
    'EBITDA Margin (%)': 'EBITDA Margin (%) (Profitability Analysis)',
    'ROE (%)': 'Return on Equity (ROE) (%) (Profitability Analysis)',
    'ROA (%)': 'Return on Assets (ROA) (%) (Profitability Analysis)',
    'ROIC (%)': 'Return on Invested Capital (ROIC) (%) (Profitability Analysis)',
    'NOPLAT Margin (%)': 'NOPLAT Margin (%) (Profitability Analysis)',

    #'Invested Capital Turnover': 'Invested Capital Turnover (Asset Management Analysis)',
    'Inventory Turnover': 'Inventory Turnover (Asset Management Analysis)',
    'Asset Turnover': 'Asset Turnover (Asset Management Analysis)',
    #'LT Asset Turnover': 'Long-term Asset Turnover (Asset Management Analysis)',
    'PPE Turnover': 'PPE Turnover (Asset Management Analysis)',
    'Depreciation/PP&E (%)': 'Depreciation/PP&E (%) (Asset Management Analysis)',
    #'Depreciation/Revenue (%)': 'Depreciation/Revenue (%) (Asset Management Analysis)',
    'Wkg Capital/Revenue (%)': 'Working Capital/Revenue (%) (Asset Management Analysis)',
    'Working Cap Turnover': 'Working Capital Turnover (Asset Management Analysis)',

    'Financial Leverage': 'Financial Leverage (Debt and Safety Analysis)',
    #'Gross Gearing (D/E) (%)': 'Gross Gearing (D/E) (%) (Debt and Safety Analysis)',
    'Net Gearing (%)': 'Net Gearing (%) (Debt and Safety Analysis)',
    #'Net Interest Cover': 'Net Interest Cover (Debt and Safety Analysis)',
    'Current Ratio': 'Current Ratio (Debt and Safety Analysis)',
    #'Quick Ratio': 'Quick Ratio (Debt and Safety Analysis)',
    #'Gross Debt/CF': 'Gross Debt/CF (Debt and Safety Analysis)',
    'Net Debt/CF': 'Net Debt/CF (Debt and Safety Analysis)',

    #'NTA per Share ($)': 'NTA per Share ($)',
    #'BV per Share ($)': 'BV per Share ($)',
    'Cash per Share ($)': 'Cash per Share ($)',
    #'Receivables/Op. Rev. (%)': 'Receivables/Op. Rev. (%)',
    #'Inventory/Trading Rev. (%)': 'Inventory/Trading Rev. (%)',
    #'Creditors/Op. Rev. (%)': 'Creditors/Op. Rev. (%)',

    'Funds from Ops./EBITDA (%)': 'Funds from Ops./EBITDA (%) (Cash Flow Analysis)',
    'Depreciation/Capex (%)': 'Depreciation/Capex (%) (Cash Flow Analysis)',
    'Capex/Operating Rev. (%)': 'Capex/Operating Rev. (%) (Cash Flow Analysis)',
    #'Days Inventory': 'Days Inventory (Asset Management Analysis)',
    #'Days Receivables': 'Days Receivables (Asset Management Analysis)',
    #'Days Payables': 'Days Payables (Asset Management Analysis)',
    'Gross CF per Share ($)': 'Gross CF per Share ($) (Cash Flow Analysis)',

    #'Sales per Share ($)': 'Sales per Share ($)',
    #'Year End Share Price': 'Year End Share Price',
    'Market Cap.': 'Market Cap.',
    'Net Debt': 'Net Debt',
    'Enterprise Value': 'Enterprise Value',
    #'EV/EBITDA': 'EV/EBITDA',
    #'EV/EBIT': 'EV/EBIT',
    #'Market Cap./Rep NPAT': 'Market Cap./Rep NPAT',
    'Market Cap./Trading Rev.': 'Market Cap./Trading Rev.',
    'Price/Book Value': 'Price/Book Value',
    'Price/Gross Cash Flow': 'Price/Gross Cash Flow',
    'PER': 'Price to Earnings Ratio (PER)',
}

if INDIVIDUAL:
    # Categories array
    categories = ['Profitability Analysis', 'Asset Management Analysis', 'Debt and Safety Analysis', 'Cash Flow Analysis', 'Valuation Measures', 'Other']

    # Ensure the Plots directory and subdirectories for each category exist
    base_path = 'Plots'
    os.makedirs(base_path, exist_ok=True)
    for category in categories:
        os.makedirs(os.path.join(base_path, category), exist_ok=True)

    # Generate and save plots
    for ratio, title in ratios.items():
        category_folder = next((category for category in categories if category in title), None)
        if category_folder:
            plot_path = os.path.join(base_path, category_folder, f"{ratio.replace('/', '_').replace(' ', '_').replace('.', '')}.png")
        else:
            plot_path = os.path.join(base_path, 'Other', f"{ratio.replace('/', '_').replace(' ', '_').replace('.', '')}.png")
        fig, ax = plt.subplots(figsize=(10, 5))
        ratio_values = []

        for ticker, clean_data in data_clean.items():
            ratio_data = filter_clean_ratio_data_common(clean_data, ratio)
            ax.plot(years_common, ratio_data.values, marker='o', linestyle='-', label=f'{ticker}')
            ratio_values.append(ratio_data.values)

        # Calculate and plot the average
        average_ratio = np.nanmean(ratio_values, axis=0)
        ax.plot(years_common, average_ratio, marker='*', linestyle='-', linewidth=2, color='k', label='Average')

        ax.set_title(title)
        ax.set_xlabel('Year')
        ax.set_ylabel('Ratio Value')
        ax.legend()
        plt.tight_layout()
        fig.savefig(plot_path)
        plt.close(fig)

else:
    # Create a figure with subplots for all ratios
    fig, axes = plt.subplots(nrows=len(ratios), ncols=1, figsize=(12, 80))
    fig.suptitle('Ratio Analysis Comparison: Webjet vs. Competitors', fontsize=16)

    # Plot each ratio for all companies
    for ax, (ratio, title) in zip(axes, ratios.items()):
        ratio_values = []
        for ticker, clean_data in data_clean.items():
            ratio_data = filter_clean_ratio_data_common(clean_data, ratio)
            ax.plot(years_common, ratio_data.values, marker='o', linestyle='-', label=f'{ticker}')
            ratio_values.append(ratio_data.values)
        # Calculate and plot the average
        average_ratio = np.nanmean(ratio_values, axis=0)
        ax.plot(years_common, average_ratio, marker='*', linestyle='-', linewidth=2, color='k', label='Average')
        
        ax.set_title(title)
        ax.set_xlabel('Year')
        ax.set_ylabel(ratio)
        ax.legend()

    plt.tight_layout(rect=[0, 0.03, 1, 0.95])
    plt.show()

    # Export as PNG
    fig.savefig('ratio_analysis_comparison.png')

