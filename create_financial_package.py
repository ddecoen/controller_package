#!/usr/bin/env python3
"""
Script to create CSV files from financial data and combine them into Excel.
"""

import pandas as pd
import io

# Income Statement Data
income_statement_data = '''"Coder Technologies, Inc",,,,
"Coder Technologies, Inc.",,,,
Month-over-Month Income Statement,,,,
Jun 2025,,,,
,,,,
Options: Activity Only,,,,
Financial Row,Amount (Jun 2025),Comparative Amount (May 2025),Variance,% Variance
Ordinary Income/Expense,,,,
Income,,,,
40000 - Revenue,,,,
40001 - Revenue - license,"$1,192,600.39 ","$1,115,498.08 ","$77,102.31 ",6.91%
40005 - Revenue - support,"$148,793.28 ","$144,111.08 ","$4,682.20 ",3.25%
Total - 40000 - Revenue,"$1,341,393.67 ","$1,259,609.16 ","$81,784.51 ",6.49%
Total - Income,"$1,341,393.67 ","$1,259,609.16 ","$81,784.51 ",6.49%
Cost Of Sales,,,,
50000 - Cost of revenue,,,,
50010 - Support,"$117,711.57 ","$114,002.68 ","$3,708.89 ",3.25%
Total - 50000 - Cost of revenue,"$117,711.57 ","$114,002.68 ","$3,708.89 ",3.25%
Total - Cost Of Sales,"$117,711.57 ","$114,002.68 ","$3,708.89 ",3.25%
Gross Profit,"$1,223,682.10 ","$1,145,606.48 ","$78,075.62 ",6.82%
Expense,,,,
60000 - Operating expenses,,,,
61000 - Compensation expenses,,,,
61011 - Salaries,"$1,380,987.81 ","$1,326,380.04 ","$54,607.77 ",4.12%
61012 - Employee bonus,"$166,176.21 ","$173,942.57 ","($7,766.36)",-4.46%
61013 - Contractor pay,"$46,227.22 ","$84,545.92 ","($38,318.70)",-45.32%
61014 - Severance,"$43,237.34 ","$53,557.43 ","($10,320.09)",-19.27%
61105 - Employee commissions,"$193,290.81 ","$217,229.96 ","($23,939.15)",-11.02%
61500 - Personnel costs,,,,
61521 - Payroll taxes,"$122,204.15 ","$118,362.62 ","$3,841.53 ",3.25%
61522 - Employee benefits,"$91,592.24 ","$110,108.37 ","($18,516.13)",-16.82%
61523 - Employer match,"$40,902.86 ","$37,359.97 ","$3,542.89 ",9.48%
61530 - Reimbursements,"$3,811.50 ","$3,689.64 ",$121.86 ,3.30%
61531 - Learning & development,"$1,885.63 ",$125.87 ,"$1,759.76 ","1,398.08%"
61532 - Stipends,"$1,035.90 ","$3,510.71 ","($2,474.81)",-70.49%
61533 - Fringe platform,$0.00 ,"$4,950.00 ","($4,950.00)",-100.00%
61534 - HeyTaco,"$9,353.28 ","$14,268.66 ","($4,915.38)",-34.45%
61535 - Gifts/other,$718.97 ,$405.36 ,$313.61 ,77.37%
61536 - Payroll fees,"$16,616.85 ","$11,676.83 ","$4,940.02 ",42.31%
Total - 61500 - Personnel costs,"$288,121.38 ","$304,458.03 ","($16,336.65)",-5.37%
Total - 61000 - Compensation expenses,"$2,118,040.77 ","$2,160,113.95 ","($42,073.18)",-1.95%
62000 - Travel & entertainment,,,,
62041 - Air transportation,"$23,247.51 ","$36,284.37 ","($13,036.86)",-35.93%
62042 - Ground transportation,"$7,965.42 ","$8,414.30 ",($448.88),-5.33%
62043 - Lodging,"$18,525.24 ","$16,821.18 ","$1,704.06 ",10.13%
62044 - Meals & entertainment,"$11,620.20 ","$20,966.85 ","($9,346.65)",-44.58%
62045 - Other travel related expenses,$138.48 ,$415.59 ,($277.11),-66.68%
Total - 62000 - Travel & entertainment,"$61,496.85 ","$82,902.29 ","($21,405.44)",-25.82%
63000 - Tools & equipment,,,,
63051 - Hardware equipment,"$1,073.96 ","$2,141.47 ","($1,067.51)",-49.85%
"63052 - Software, platform, and other tools","$165,843.07 ","$197,568.29 ","($31,725.22)",-16.06%
Total - 63000 - Tools & equipment,"$166,917.03 ","$199,709.76 ","($32,792.73)",-16.42%
64000 - Professional fees,,,,
64061 - Professional services,"$283,903.28 ","$149,828.50 ","$134,074.78 ",89.49%
64062 - Legal fees,"$9,000.00 ","$13,396.67 ","($4,396.67)",-32.82%
Total - 64000 - Professional fees,"$292,903.28 ","$163,225.17 ","$129,678.11 ",79.45%
65000 - Administrative,,,,
65090 - Bank charges & payment processing,$510.94 ,$236.11 ,$274.83 ,116.40%
"65091 - Shipping, freight, & delivery","$4,613.80 ",$399.72 ,"$4,214.08 ","1,054.26%"
65110 - Business insurance,"$3,085.26 ","$3,389.38 ",($304.12),-8.97%
65120 - Administrative,$211.16 ,$673.15 ,($461.99),-68.63%
65130 - Recruiting,"$12,288.48 ","$1,200.00 ","$11,088.48 ",924.04%
65910 - Depreciation,"$7,424.93 ","$7,037.26 ",$387.67 ,5.51%
65950 - Business license & taxes,"$2,682.31 ",$0.00 ,"$2,682.31 ",0.00%
Total - 65000 - Administrative,"$30,816.88 ","$12,935.62 ","$17,881.26 ",138.23%
66000 - Research & development,,,,
66070 - Cloud costs,"$36,347.04 ","$37,823.46 ","($1,476.42)",-3.90%
Total - 66000 - Research & development,"$36,347.04 ","$37,823.46 ","($1,476.42)",-3.90%
67000 - Sales and marketing,,,,
67210 - Events & tradeshows,"$128,184.26 ","$61,928.64 ","$66,255.62 ",106.99%
67220 - Content marketing & campaigns,"$12,157.63 ","$50,181.12 ","($38,023.49)",-75.77%
67240 - Field marketing,"$26,489.79 ",$0.00 ,"$26,489.79 ",0.00%
67250 - Advertising & paid media,"$163,302.29 ","$54,393.27 ","$108,909.02 ",200.23%
67260 - Marketing supplies,"$122,020.79 ","$46,592.35 ","$75,428.44 ",161.89%
67290 - Marketing - other,$0.00 ,$636.60 ,($636.60),-100.00%
Total - 67000 - Sales and marketing,"$452,154.76 ","$213,731.98 ","$238,422.78 ",111.55%
68000 - Office expense,,,,
68071 - Rent,"$16,476.46 ","$35,089.13 ","($18,612.67)",-53.04%
68072 - CAM Operating Costs,"$41,484.55 ","$5,900.00 ","$35,584.55 ",603.13%
68073 - Office equipment/workstations,"$2,608.26 ",$712.78 ,"$1,895.48 ",265.93%
68074 - Office consumables/supplies,"$2,903.01 ","$1,295.25 ","$1,607.76 ",124.13%
68076 - Utilities/services,"$1,828.26 ","$2,293.13 ",($464.87),-20.27%
Total - 68000 - Office expense,"$65,300.54 ","$45,290.29 ","$20,010.25 ",44.18%
69000 - Other Operating Expenses,,,,
69000 - Other Operating Expenses,$680.68 ,$0.00 ,$680.68 ,0.00%
69102 - State Franchise or Income Tax,$0.00 ,"$4,200.00 ","($4,200.00)",-100.00%
Total - 69000 - Other Operating Expenses,$680.68 ,"$4,200.00 ","($3,519.32)",-83.79%
Total - 60000 - Operating expenses,"$3,224,657.83 ","$2,919,932.52 ","$304,725.31 ",10.44%
Total - Expense,"$3,224,657.83 ","$2,919,932.52 ","$304,725.31 ",10.44%
Net Ordinary Income,"($2,000,975.73)","($1,774,326.04)","($226,649.69)",12.77%
Other Income and Expenses,,,,
Other Income,,,,
71000 - Other income,,,,
71002 - Interest income,$411.55 ,$425.13 ,($13.58),-3.19%
71004 - Dividend income,"$75,474.39 ","$84,817.16 ","($9,342.77)",-11.02%
Total - 71000 - Other income,"$75,885.94 ","$85,242.29 ","($9,356.35)",-10.98%
Total - Other Income,"$75,885.94 ","$85,242.29 ","($9,356.35)",-10.98%
Other Expense,,,,
71050 - Realized Gain/Loss on FX,$188.84 ,$455.16 ,($266.32),-58.51%
Total - Other Expense,$188.84 ,$455.16 ,($266.32),-58.51%
Net Other Income,"$75,697.10 ","$84,787.13 ","($9,090.03)",-10.72%
Net Income,"($1,925,278.63)","($1,689,538.91)","($235,739.72)",13.95%'''

# Balance Sheet Data (truncated for brevity - you can add the full data)
balance_sheet_data = '''"Coder Technologies, Inc",,,,
"Coder Technologies, Inc.",,,,
Month-over-Month Balance Sheet,,,,
End of Jun 2025,,,,
,,,,
Options: Activity Only,,,,
Financial Row,Amount (As of Jun 2025),Comparison Amount (As of May 2025),Variance,% Variance
ASSETS,,,,
Current Assets,,,,
Bank,,,,
11000 - Cash and cash equivalents,,,,
11001 - JPM operating - 3369,"$860,497.76 ","$1,169,271.12 ","($308,773.36)",-26.41%
11002 - JPM money market - 6601,"$626,404.90 ","$626,343.35 ",$61.55 ,0.01%
11005 - JPM investment - 4586,"$21,998,175.02 ","$22,922,700.63 ","($924,525.61)",-4.03%
Total - 11000 - Cash and cash equivalents,"$23,485,077.68 ","$24,718,315.10 ","($1,233,237.42)",-4.99%
Total Bank,"$23,485,077.68 ","$24,718,315.10 ","($1,233,237.42)",-4.99%
Accounts Receivable,,,,
12000 - Receivables,,,,
12001 - Accounts receivable - trade,"$3,134,835.66 ","$3,348,408.54 ","($213,572.88)",-6.38%
Total - 12000 - Receivables,"$3,134,835.66 ","$3,348,408.54 ","($213,572.88)",-6.38%
Total Accounts Receivable,"$3,134,835.66 ","$3,348,408.54 ","($213,572.88)",-6.38%
Other Current Asset,,,,
13000 - Prepaid expenses,,,,
13010 - Prepaid software & subscriptions,"$464,890.83 ","$449,374.42 ","$15,516.41 ",3.45%
13011 - Prepaid marketing expenses,"$644,927.89 ","$762,252.21 ","($117,324.32)",-15.39%
13012 - Prepaid offsite/conference costs,"$189,374.88 ","$138,441.87 ","$50,933.01 ",36.79%
13013 - Prepaid insurance,"$26,980.22 ","$30,065.48 ","($3,085.26)",-10.26%
13100 - Prepaid Other,"$70,430.02 ","$55,930.02 ","$14,500.00 ",25.93%
Total - 13000 - Prepaid expenses,"$1,396,603.84 ","$1,436,064.00 ","($39,460.16)",-2.75%
Total ASSETS,"$28,892,240.93 ","$30,310,166.94 ","($1,417,926.01)",-4.68%'''

def create_csv_files():
    """Create CSV files from the data strings."""
    
    # Write income statement CSV
    with open('income_statement.csv', 'w', encoding='utf-8') as f:
        f.write(income_statement_data)
    
    # Write balance sheet CSV
    with open('balance_sheet.csv', 'w', encoding='utf-8') as f:
        f.write(balance_sheet_data)
    
    print("CSV files created successfully!")

def combine_to_excel():
    """Combine CSV files into Excel workbook."""
    
    # Create Excel writer
    with pd.ExcelWriter('coder_financial_package.xlsx', engine='openpyxl') as writer:
        
        # Read and write income statement
        try:
            df_income = pd.read_csv('income_statement.csv')
            df_income.to_excel(writer, sheet_name='Income Statement', index=False)
            print("Added Income Statement sheet")
        except Exception as e:
            print(f"Error processing income statement: {e}")
        
        # Read and write balance sheet
        try:
            df_balance = pd.read_csv('balance_sheet.csv')
            df_balance.to_excel(writer, sheet_name='Balance Sheet', index=False)
            print("Added Balance Sheet sheet")
        except Exception as e:
            print(f"Error processing balance sheet: {e}")
    
    print("\nExcel file 'coder_financial_package.xlsx' created successfully!")

def main():
    print("Creating financial package...")
    
    # Create CSV files
    create_csv_files()
    
    # Combine into Excel
    combine_to_excel()
    
    print("\nFinancial package creation complete!")
    print("Files created:")
    print("  - income_statement.csv")
    print("  - balance_sheet.csv")
    print("  - coder_financial_package.xlsx")

if __name__ == "__main__":
    main()
