import pandas as pd

# Sample data for testing
income_data = [
    ['Financial Row', 'Amount (Jun 2025)', 'Comparative Amount (May 2025)', 'Variance', '% Variance'],
    ['Revenue - license', '$1,192,600.39', '$1,115,498.08', '$77,102.31', '6.91%'],
    ['Revenue - support', '$148,793.28', '$144,111.08', '$4,682.20', '3.25%'],
    ['Total Revenue', '$1,341,393.67', '$1,259,609.16', '$81,784.51', '6.49%']
]

balance_data = [
    ['Financial Row', 'Amount (As of Jun 2025)', 'Comparison Amount (As of May 2025)', 'Variance', '% Variance'],
    ['Cash and cash equivalents', '$23,485,077.68', '$24,718,315.10', '($1,233,237.42)', '-4.99%'],
    ['Accounts receivable', '$3,134,835.66', '$3,348,408.54', '($213,572.88)', '-6.38%'],
    ['Total Assets', '$28,892,240.93', '$30,310,166.94', '($1,417,926.01)', '-4.68%']
]

# Create DataFrames
df_income = pd.DataFrame(income_data[1:], columns=income_data[0])
df_balance = pd.DataFrame(balance_data[1:], columns=balance_data[0])

# Create Excel file
with pd.ExcelWriter('financial_package.xlsx', engine='openpyxl') as writer:
    df_income.to_excel(writer, sheet_name='Income Statement', index=False)
    df_balance.to_excel(writer, sheet_name='Balance Sheet', index=False)

print("Excel file created: financial_package.xlsx")
