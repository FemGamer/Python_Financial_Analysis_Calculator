import pandas as pd
import time 

###for BlackRock

class File_1:
    start_time= time.time()
    def excel(file_path, sheet_name):
        return pd.read_excel(file_path, sheet_name=sheet_name, index_col=0)
        
    def calculate_ratios(df_income, df_balance, df_cashflow):
        
        NI= df_income.loc['Net income'].values
        Rev= df_income.loc['Total revenue'].values
        NPR= (NI/Rev)
        #print("Net profit ratio: ", NPR)

        SE= df_balance.loc['Total stockholders equity'].values 
        TL= df_balance.loc['Total liabilities'].values
        DER= TL/SE
        #print("Debt-Equity Ratio: ", DER)

        NI= df_cashflow.loc['Net income'].values
        Divi= df_cashflow.loc['Dividends paid'].values
        PR= Divi/NI
        #print("Payout Ratio:", PR)

        TA= df_balance.loc['Total assets'].values 
        TL= df_balance.loc['Total liabilities'].values
        
        DA = TL/TA
        #print('Debit-Asset Ratio: ',DA)

        #current assets
        CCE= df_balance.loc['Cash and cash equivalents'].values
        AR= df_balance.loc['Accounts receivable'].values
        Invst= df_balance.loc['Investments'].values

        CA= CCE+AR+Invst
        #print("Current Assets:", CA)

        ##current liabilities
        SB= df_balance.loc['Borrowings'].values
        AP= df_balance.loc['Accounts payable'].values
        AE= df_balance.loc['Accrued compensation and benefits'].values

        CL= SB+AP+AE
        #print("Current liabilities: ", CL)

        CuR= CA/CL
        #print('Current Ratio: ', CuR)

        WorkCap = CA-CL
        #print('Working Capital: ', WorkCap)


        SE= df_balance.loc['Total stockholders equity'].values 
        TA= df_balance.loc['Total assets'].values 
        LR= SE/TA
        #print("Leverage Ratio:", LR)     

        return{
            'Debit-Asset Ratio': DA,
            'Payout Ratio': PR,
            'Debt-Equity Ratio': DER,
            'Net profit ratio': NPR,
            "Leverage Ratio:": LR,
            'Working Capital: ': WorkCap,
            'Current Ratio: ': CuR,
        }

    file_paths1=[
            
            "D:/PY_FemGamer/DataBase/BR_XLSX/2014.xlsx",
            "D:/PY_FemGamer/DataBase/BR_XLSX/2015.xlsx",
            "D:/PY_FemGamer/DataBase/BR_XLSX/2016.xlsx",
            "D:/PY_FemGamer/DataBase/BR_XLSX/2017.xlsx",
            "D:/PY_FemGamer/DataBase/BR_XLSX/2018.xlsx",
            "D:/PY_FemGamer/DataBase/BR_XLSX/2019.xlsx",
            "D:/PY_FemGamer/DataBase/BR_XLSX/2020.xlsx",
            "D:/PY_FemGamer/DataBase/BR_XLSX/2021.xlsx",
            "D:/PY_FemGamer/DataBase/BR_XLSX/2022.xlsx",
            "D:/PY_FemGamer/DataBase/BR_XLSX/2023.xlsx",
            
          
            ]
        
        
    output_data=[]
        
    for file_path in file_paths1:
        df_income= excel(file_path, sheet_name="Income Statement")
        df_balance= excel(file_path, sheet_name="Balance Sheet")
        df_cashflow= excel(file_path, sheet_name="Cash Flow Statement")

        print("File :", file_path)
        calculator=calculate_ratios(df_income, df_balance, df_cashflow)
        print()
        output_data.append(calculator)

    output_df = pd.DataFrame(output_data)
    output_file = "D:/PY_FemGamer/Thesis/Output/Ratios_BR.xlsx"
    output_df.to_excel(output_file, index=False)

    print("Output saved to:", output_file)

    end_time= time.time()
    Program_Runtime= (end_time - start_time)/60
    print("Runtime: ", round(Program_Runtime, 2), "minutes") 

#######################################################################

