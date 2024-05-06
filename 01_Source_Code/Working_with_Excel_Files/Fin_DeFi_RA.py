####for ratio analysis seperate file for the dataset from file Fin_Defi_Dataset.py


import yfinance as yf
import pandas as pd
import time
from matplotlib.pyplot import*

start_time= time.time()

class Ratio_Calculations:  
    # here excel files of Inc_Statement, Bal_Sheet, CF_statement are called separately
    IS = "D:/PY_FemGamer/Thesis/02_DataBase/Income_Statement.xlsx"
    BS = "D:/PY_FemGamer/Thesis/02_DataBase/Balance_Sheet.xlsx"
    CF = "D:/PY_FemGamer/Thesis/02_DataBase/Cash_Flow.xlsx"

    df_income= pd.read_excel(IS, sheet_name='COIN', index_col=0) ##activating the worksheet of the company this time
    df_balance= pd.read_excel(BS, sheet_name='COIN', index_col=0)
    df_cashflow= pd.read_excel(CF, sheet_name='COIN', index_col=0)

    ##finding the values and doing ratio calculations
    NI= df_income.loc['Net Income'].values 
    Rev= df_income.loc['Total Revenue'].values 
    NPR= (NI/Rev)
    #print("Net profit ratio: ", NPR)

    TCL= df_balance.loc['Current Liabilities'].values 
    TCA= df_balance.loc['Current Assets'].values
    CR= TCA / TCL
    #print("Current Ratio: ", CR) 

    SE= df_balance.loc['Stockholders Equity'].values
    TL= df_balance.loc['Total Liabilities Net Minority Interest'].values
    DER= TL/SE
    #print("Debt-Equity Ratio: ", DER)

    TA= df_balance.loc['Total Assets'].values
    TL= df_balance.loc['Total Liabilities Net Minority Interest'].values
    DAR= TL/TA
    #print("Debt-Asset Ratio: ", DAR)
        
                
    TCL= df_balance.loc['Current Liabilities'].values 
    OCF= df_cashflow.loc['Operating Cash Flow'].values 
    CFR= OCF/TCL
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

    with pd.ExcelWriter('D:/PY_FemGamer/Thesis/03_Output/Ratios_COIN.xlsx') as writer:
        df.to_excel(writer, sheet_name='Ratios', index=False)


    print('outputs are saved in excel') 

    end_time= time.time()
    Program_Runtime= (end_time - start_time)/60
    print("Runtime: ", round(Program_Runtime, 2), "minutes")

    