import pandas as pd

# Load Excel sheets
income = pd.read_excel("financials.xlsx", sheet_name="Income Statement")
balance = pd.read_excel("financials.xlsx", sheet_name="Balance Sheet")

# Set index
income.set_index("Particulars", inplace=True)
balance.set_index("Particulars", inplace=True)

# Income Statement values
revenue_2023 = income.loc["Revenue", 2023]
revenue_2024 = income.loc["Revenue", 2024]

profit_2023 = income.loc["Net Profit", 2023]
profit_2024 = income.loc["Net Profit", 2024]

# Growth calculations
revenue_growth = ((revenue_2024 - revenue_2023) / revenue_2023) * 100
profit_growth = ((profit_2024 - profit_2023) / profit_2023) * 100

# Balance Sheet values (2024)
current_assets = balance.loc["Current Assets", 2024]
current_liabilities = balance.loc["Current Liabilities", 2024]
total_liabilities = balance.loc["Total Liabilities", 2024]
shareholders_equity = balance.loc["Shareholders Equity", 2024]

# Ratio calculations
current_ratio = current_assets / current_liabilities
debt_equity_ratio = total_liabilities / shareholders_equity
net_profit_margin = (profit_2024 / revenue_2024) * 100

# Output results
print("FINANCIAL STATEMENT ANALYSIS (FY 2023–2024)\n")

print(f"Revenue Growth: {revenue_growth:.2f}%")
print(f"Net Profit Growth: {profit_growth:.2f}%")
print(f"Current Ratio (2024): {current_ratio:.2f}")
print(f"Debt–Equity Ratio (2024): {debt_equity_ratio:.2f}")
print(f"Net Profit Margin (2024): {net_profit_margin:.2f}%")
