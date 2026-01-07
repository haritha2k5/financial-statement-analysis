import pandas as pd
import matplotlib.pyplot as plt

# -----------------------------
# LOAD DATA
# -----------------------------
income = pd.read_excel("financials.xlsx", sheet_name="Income Statement")
balance = pd.read_excel("financials.xlsx", sheet_name="Balance Sheet")
cashflow = pd.read_excel("financials.xlsx", sheet_name="Cash Flow Statement")

# Set index
income.set_index("Particulars", inplace=True)
balance.set_index("Particulars", inplace=True)
cashflow.set_index("Particulars", inplace=True)

# Clean index labels (handles apostrophes & spaces)
balance.index = balance.index.str.replace("'", "").str.strip()

# -----------------------------
# INCOME STATEMENT VALUES
# -----------------------------
revenue_2023 = income.loc["Revenue", 2023]
revenue_2024 = income.loc["Revenue", 2024]

profit_2023 = income.loc["Net Profit", 2023]
profit_2024 = income.loc["Net Profit", 2024]

expenses_2023 = income.loc["Expenses", 2023]
expenses_2024 = income.loc["Expenses", 2024]

ebitda_2023 = income.loc["EBITDA", 2023]
ebitda_2024 = income.loc["EBITDA", 2024]

# -----------------------------
# BALANCE SHEET VALUES (2024)
# -----------------------------
current_assets = balance.loc["Current Assets", 2024]
current_liabilities = balance.loc["Current Liabilities", 2024]
total_liabilities = balance.loc["Total Liabilities", 2024]
shareholders_equity = balance.loc["Shareholders Equity", 2024]

# -----------------------------
# CASH FLOW VALUES (2024)
# -----------------------------
cfo_2024 = cashflow.loc["Cash flow from Operating Activities", 2024]
cfi_2024 = cashflow.loc["Cash flow from Investing Activities", 2024]
cff_2024 = cashflow.loc["Cash flow from Financing Activities", 2024]

# -----------------------------
# CALCULATIONS
# -----------------------------
revenue_growth = ((revenue_2024 - revenue_2023) / revenue_2023) * 100
profit_growth = ((profit_2024 - profit_2023) / profit_2023) * 100

current_ratio = current_assets / current_liabilities
debt_equity_ratio = total_liabilities / shareholders_equity
net_profit_margin = (profit_2024 / revenue_2024) * 100

# -----------------------------
# OUTPUT
# -----------------------------
print("\nFINANCIAL STATEMENT ANALYSIS (FY 2023–2024)\n")
print(f"Revenue Growth: {revenue_growth:.2f}%")
print(f"Net Profit Growth: {profit_growth:.2f}%")
print(f"Current Ratio (2024): {current_ratio:.2f}")
print(f"Debt–Equity Ratio (2024): {debt_equity_ratio:.2f}")
print(f"Net Profit Margin (2024): {net_profit_margin:.2f}%")

# -----------------------------
# PLOTS
# -----------------------------

# 1. Revenue comparison
plt.figure()
plt.bar(["2023", "2024"], [revenue_2023, revenue_2024])
plt.title("Revenue Comparison (FY 2023 vs FY 2024)")
plt.ylabel("Revenue")
plt.show()

# 2. Net Profit comparison
plt.figure()
plt.bar(["2023", "2024"], [profit_2023, profit_2024])
plt.title("Net Profit Comparison (FY 2023 vs FY 2024)")
plt.ylabel("Net Profit")
plt.show()

# 3. Revenue vs Expenses
plt.figure()
plt.bar(["Revenue 2023", "Revenue 2024"], [revenue_2023, revenue_2024])
plt.bar(["Expenses 2023", "Expenses 2024"], [expenses_2023, expenses_2024])
plt.title("Revenue vs Expenses Comparison")
plt.show()

# 4. EBITDA vs Net Profit
plt.figure()
plt.bar(["EBITDA 2023", "EBITDA 2024"], [ebitda_2023, ebitda_2024])
plt.bar(["Net Profit 2023", "Net Profit 2024"], [profit_2023, profit_2024])
plt.title("EBITDA vs Net Profit Comparison")
plt.show()

# 5. Key Financial Ratios (2024)
plt.figure()
plt.bar(
    ["Current Ratio", "Debt–Equity", "Net Profit Margin"],
    [current_ratio, debt_equity_ratio, net_profit_margin]
)
plt.title("Key Financial Ratios (FY 2024)")
plt.show()

# 6. Cash Flow Breakdown (2024)
plt.figure()
plt.bar(
    ["Operating", "Investing", "Financing"],
    [cfo_2024, cfi_2024, cff_2024]
)
plt.title("Cash Flow Breakdown (FY 2024)")
plt.show()
