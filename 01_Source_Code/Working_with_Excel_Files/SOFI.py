###separte script was made for SOFI and also its dataset was copied nad paste into a separate excel file due to date and time error

import yfinance as yf
import pandas as pd
import time

start_time= time.time()
data_file= "D:/PY_FemGamer/Thesis/02_DataBase/SOFI.xlsx"



df_income= pd.read_excel(data_file, sheet_name="Income Statement", index_col=0)
df_balance= pd.read_excel(data_file, sheet_name="Balance Sheet", index_col=0)
df_cashflow= pd.read_excel(data_file, sheet_name="Cash Flow Statement", index_col=0)


NI= df_income.loc['Net Income'].values 
Rev= df_income.loc['Total Revenue'].values 
NPR= (NI/Rev)
    #print("Net profit ratio: ", NPR)

   
AC= df_balance.loc['Receivables'].values 
ST= df_balance.loc['Other Short Term Investments'].values
CCE= df_balance.loc['Cash And Cash Equivalents'].values
PA= df_balance.loc['Prepaid Assets'].values
OR= df_balance.loc['Other Receivables'].values
CA= AC+ST+CCE+PA
    #print(CA)
CL= df_balance.loc['Payables And Accrued Expenses'].values
# AccExp= df_balance.loc['Current Accrued Expenses'].values
# CL= AP+ AccExp
    #print(CL)
CR= CA/CL
    #print("Current Ratio: ", CR) 

SE= df_balance.loc['Stockholders Equity'].values
TL= df_balance.loc['Total Liabilities Net Minority Interest'].values
DER= TL/SE
    #print("Debt-Equity Ratio: ", DER)

TA= df_balance.loc['Total Assets'].values
TL= df_balance.loc['Total Liabilities Net Minority Interest'].values
DAR= TL/TA
    #print("Debt-Asset Ratio: ", DAR)
        
                
OCF= df_cashflow.loc['Operating Cash Flow'].values 
CFR= OCF/CL
    #print("Cash Flow Ratio:", CFR)
   
                
SE= df_balance.loc['Stockholders Equity'].values 
TA= df_balance.loc['Total Assets'].values
LR= TA/SE
    #print("Leverage Ratio:", LR)

list_ratios={
        'Net Porfit Ratio: ': NPR,
        'Current Ratio: ': CR,
        'Debt-Equity Ratio: ': DER,
        'Cash Flow Ratio:': CFR,
        'Leverage Ratio:': LR,
        'Debt-Asset Ratio: ': DAR,
    }

data= list_ratios
df= pd.DataFrame(data)
    #print(df)
    
    ##separate output sheets for each company for clarity##

with pd.ExcelWriter('D:/PY_FemGamer/Thesis/03_Output/Ratios_SOFI.xlsx') as writer:
    df.to_excel(writer, sheet_name='Ratios', index=False)


print('outputs are saved in excel') 

end_time= time.time()
Program_Runtime= (end_time - start_time)/60
print("Runtime: ", round(Program_Runtime, 2), "minutes")

    