##for ECB dataset same functions from VISA, UHG PayPal company as their dataset is in Excel File Format

import pandas as pd
import time

start_time = time.time()

def excel(file_path, sheet_name):
    return pd.read_excel(file_path, sheet_name=sheet_name, index_col=0)
        
def calculate_ratios(df_income, df_balance):
    NI = df_income.loc['Net income'].values
    Rev = df_income.loc['Net revenues'].values
    NPR = NI / Rev

    TA = df_balance.loc['Total assets'].values 
    TL = df_balance.loc['Total liabilities'].values
    DA = TL / TA

    ratios= {
        'Debit-Asset Ratio': DA,
        'Net profit ratio': NPR,
    }

    return ratios

output_data = []  

file_paths = ["D:/PY_FemGamer/Thesis/02_DataBase/ECB_XLSX/ECB.xlsx",]

for file_path in file_paths:
    df_income = excel(file_path, sheet_name="Income Statement")
    df_balance = excel(file_path, sheet_name="Balance Sheet")
                                
    print("File :", file_path)
    calculator = calculate_ratios(df_income, df_balance)
    print()
    output_data.append(calculator)

output_df = pd.DataFrame(output_data)
output_file = "D:/PY_FemGamer/Thesis/03_Output/Ratios_File3.xlsx"
output_df.to_excel(output_file, sheet_name='Ratios', index=False)

print("Output saved to:", output_file)

end_time = time.time()
Program_Runtime = (end_time - start_time) / 60
print("Runtime: ", round(Program_Runtime, 2), "minutes")
