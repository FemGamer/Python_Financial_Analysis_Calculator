##this file is for ratio analysis for JP Morgan Chase and Goldman Sachs

import pandas as pd
import time 

###for JPM & GS 

class File_1:
    start_time= time.time()
    def excel(file_path, sheet_name):
        return pd.read_excel(file_path, sheet_name=sheet_name, index_col=0) ##funtion to read excel file into the program
        
    def calculate_ratios(df_income, df_balance, df_cashflow):###ratios are formulated here according to their formulas
        
        NI= df_income.loc['Net income'].values
        Rev= df_income.loc['Net revenues'].values
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

        #current assets which were assumed this case
        CCE= df_balance.loc['Cash and cash equivalents'].values
        DB= df_balance.loc['Deposits with banks'].values
        FDS= df_balance.loc['Federal funds sold'].values
        TrA= df_balance.loc['Trading assets'].values
        NL= df_balance.loc['Loans Net'].values
        AR= df_balance.loc['Accounts receivable'].values

        CA= CCE+DB+FDS+TrA+NL+AR
        #print("Current Assets:", CA)

        ##current liabilities which were assumed this case
        Dl= df_balance.loc['Deposits'].values
        FDP= df_balance.loc['Federal funds purchased'].values
        SB= df_balance.loc['Short-term borrowings'].values
        TrL= df_balance.loc['Trading liabilities'].values
        AP= df_balance.loc['Accounts payable'].values

        CL= Dl+FDP+SB+TrL+AP
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

##calling of files from the storage, each files will be run separately for better outcome.
    file_paths1=[
            
            # "D:/PY_FemGamer/DataBase/GS_XLSX/01_31-12-2013.xlsx",
            # "D:/PY_FemGamer/DataBase/GS_XLSX/02_31-12-2014.xlsx",
            # "D:/PY_FemGamer/DataBase/GS_XLSX/03_31-12-2015.xlsx",
            # "D:/PY_FemGamer/DataBase/GS_XLSX/04_31-12-2016.xlsx",
            # "D:/PY_FemGamer/DataBase/GS_XLSX/05_31-12-2017.xlsx",
            # "D:/PY_FemGamer/DataBase/GS_XLSX/06_31-12-2018.xlsx",
            # "D:/PY_FemGamer/DataBase/GS_XLSX/07_31-12-2019.xlsx",
            # "D:/PY_FemGamer/DataBase/GS_XLSX/08_31-12-2020.xlsx",
            # "D:/PY_FemGamer/DataBase/GS_XLSX/09_31-12-2021.xlsx",
            # "D:/PY_FemGamer/DataBase/GS_XLSX/10_31-12-2022.xlsx",
            
            "D:/PY_FemGamer/DataBase/JPM_XLSX/01_31-12-2013.xlsx",
            "D:/PY_FemGamer/DataBase/JPM_XLSX/02_31-12-2014.xlsx",
            "D:/PY_FemGamer/DataBase/JPM_XLSX/03_31-12-2015.xlsx",
            "D:/PY_FemGamer/DataBase/JPM_XLSX/04_31-12-2016.xlsx",
            "D:/PY_FemGamer/DataBase/JPM_XLSX/05_31-12-2017.xlsx",
            "D:/PY_FemGamer/DataBase/JPM_XLSX/06_31-12-2018.xlsx",
            "D:/PY_FemGamer/DataBase/JPM_XLSX/07_31-12-2019.xlsx",
            "D:/PY_FemGamer/DataBase/JPM_XLSX/08_31-12-2020.xlsx",
            "D:/PY_FemGamer/DataBase/JPM_XLSX/09_31-12-2021.xlsx",
            "D:/PY_FemGamer/DataBase/JPM_XLSX/10_31-12-2022.xlsx",

            ]
        
        
    output_data=[]
        
    for file_path in file_paths1:
        df_income= excel(file_path, sheet_name="Income Statement") ### here activating the income statement worksheet from the file for getting paticular data
        df_balance= excel(file_path, sheet_name="Balance Sheet")### ## here activating the balance sheet worksheet from the file for getting paticular data
        df_cashflow= excel(file_path, sheet_name="Cash Flow Statement")### ## here activating the cash flow statement worksheet from the file for getting paticular data

        print("File :", file_path)
        calculator=calculate_ratios(df_income, df_balance, df_cashflow)
        print()
        output_data.append(calculator)

    output_df = pd.DataFrame(output_data)### the funtion here performs all calculations in one-go
    ####saving the output of calculations in the excel file####

    #output_file = "D:/PY_FemGamer/Thesis/Ratios_GS.xlsx"
    output_file = "D:/PY_FemGamer/Thesis/Ratios_JPM.xlsx"
    output_df.to_excel(output_file, index=False)

    print("Output saved to:", output_file)

    end_time= time.time()
    Program_Runtime= (end_time - start_time)/60
    print("Runtime: ", round(Program_Runtime, 2), "minutes") 

#######################################################################
##########END###########

