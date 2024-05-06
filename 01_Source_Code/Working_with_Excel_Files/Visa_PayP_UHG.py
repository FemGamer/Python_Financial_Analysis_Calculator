#for VISA, UHG PayPal company as their dataset is in Excel File Format


import pandas as pd
import time 

class File_2:
    start_time= time.time() 
    def excel(file_path, sheet_name):
        return pd.read_excel(file_path, sheet_name=sheet_name, index_col=0) ##funtion to read excel file into the program
    
    def calculate_ratios(df_income, df_balance, df_cashflow): ###ratios are formulated here according to their formulas
        
        NI= df_income.loc['Net income'].values
        Rev= df_income.loc['Net revenues'].values
        NPR= (NI/Rev)
        #print("Net profit ratio: ", NPR)

        SE= df_balance.loc['Total stockholders equity'].values 
        TL= df_balance.loc['Total liabilities'].values
        DER= TL/SE
        #print("Debt-Equity Ratio: ", DER)

        # NI= df_cashflow.loc['Net income'].values
        # Divi= df_cashflow.loc['Dividends paid'].values
        # PR= Divi/NI
        # #print("Payout Ratio:", PR)

        TA= df_balance.loc['Total assets'].values 
        TL= df_balance.loc['Total liabilities'].values
        DA = TL/TA
        #print('Debit-Asset Ratio: ',DA)


        TCA= df_balance.loc['Total current assets'].values
        TCL= df_balance.loc['Total current liabilities'].values

        CR= TCA/TCL
        #print('Current Ratio: ', CR)
        
        OCF= df_cashflow.loc['Net cash provided by operating activities'].values
        print("Operating Cash Flow: ", OCF)
        #print("Cash Flow Ratio:", CFRatio)
        #CFR= OCF/TCL

        SE= df_balance.loc['Total stockholders equity'].values 
        TA= df_balance.loc['Total assets'].values 
        LR= SE/TA
        #print("Leverage Ratio:", LR)     

        return{
            'Debit-Asset Ratio': DA,
            'Debt-Equity Ratio': DER,
            'Net profit ratio': NPR,
            'Leverage Ratio:': LR,
            #'Cash Flow Ratio:' : CFR,
            'Current Ratio: ': CR,
            'Operating Cash Flow: ': OCF,
            'Total current liabilities': TCL,
            
            
        }

##calling of files from the storage, each files will be run separately for better outcome.
    file_paths2=[
        
            # "D:/PY_FemGamer/Thesis/02_DataBase/VISA_XLSX/02_30-09-2014.xlsx",
            # "D:/PY_FemGamer/Thesis/02_DataBase/VISA_XLSX/03_30-09-2015.xlsx",
            # "D:/PY_FemGamer/Thesis/02_DataBase/VISA_XLSX/04_30-09-2016.xlsx",
            # "D:/PY_FemGamer/Thesis/02_DataBase/VISA_XLSX/05_30-09-2017.xlsx",
            # "D:/PY_FemGamer/Thesis/02_DataBase/VISA_XLSX/06_30-09-2018.xlsx",
            # "D:/PY_FemGamer/Thesis/02_DataBase/VISA_XLSX/07_30-09-2019.xlsx",
            # "D:/PY_FemGamer/Thesis/02_DataBase/VISA_XLSX/08_30-09-2020.xlsx",
            # "D:/PY_FemGamer/Thesis/02_DataBase/VISA_XLSX/09_30-09-2021.xlsx",
            # "D:/PY_FemGamer/Thesis/02_DataBase/VISA_XLSX/10_30-09-2022.xlsx",
            # "D:/PY_FemGamer/Thesis/02_DataBase/VISA_XLSX/11_30-09-2023.xlsx",
            

            "D:/PY_FemGamer/Thesis/02_DataBase/PayPal_XLSX/01_2015.xlsx",
            "D:/PY_FemGamer/Thesis/02_DataBase/PayPal_XLSX/02_2016.xlsx",
            "D:/PY_FemGamer/Thesis/02_DataBase/PayPal_XLSX/03_2017.xlsx",
            "D:/PY_FemGamer/Thesis/02_DataBase/PayPal_XLSX/04_2018.xlsx",
            "D:/PY_FemGamer/Thesis/02_DataBase/PayPal_XLSX/05_2019.xlsx",
            "D:/PY_FemGamer/Thesis/02_DataBase/PayPal_XLSX/06_2020.xlsx",
            "D:/PY_FemGamer/Thesis/02_DataBase/PayPal_XLSX/07_2021.xlsx",
            "D:/PY_FemGamer/Thesis/02_DataBase/PayPal_XLSX/08_2022.xlsx",
            "D:/PY_FemGamer/Thesis/02_DataBase/PayPal_XLSX/09_2023.xlsx",

            # "D:/PY_FemGamer/02_DataBase/UHG_XLSX/2014.xlsx",
            # "D:/PY_FemGamer/02_DataBase/UHG_XLSX/2015.xlsx",
            # "D:/PY_FemGamer/02_DataBase/UHG_XLSX/2016.xlsx",
            # "D:/PY_FemGamer/02_DataBase/UHG_XLSX/2017.xlsx",
            # "D:/PY_FemGamer/02_DataBase/UHG_XLSX/2018.xlsx",
            # "D:/PY_FemGamer/02_DataBase/UHG_XLSX/2019.xlsx",
            # "D:/PY_FemGamer/02_DataBase/UHG_XLSX/2020.xlsx",
            # "D:/PY_FemGamer/02_DataBase/UHG_XLSX/2021.xlsx",
            # "D:/PY_FemGamer/02_DataBase/UHG_XLSX/2022.xlsx",
            # "D:/PY_FemGamer/02DataBase/UHG_XLSX/2023.xlsx",
    
    ]

    output_data2=[]
    for file_path in file_paths2:
        df_income= excel(file_path, sheet_name="Income Statement")## here activating the income statement worksheet from the file for getting paticular data
        df_balance= excel(file_path, sheet_name="Balance Sheet")## ## here activating the balance sheet worksheet from the file for getting paticular data
        df_cashflow= excel(file_path, sheet_name="Cash Flow Statement") ## ## here activating the cash flow statement worksheet from the file for getting paticular data

        print("File :", file_path)
        calculator=calculate_ratios(df_income, df_balance, df_cashflow)
        print()
        output_data2.append(calculator)### the funtion here performs all calculations in one-go

    output_df = pd.DataFrame(output_data2)
    ####saving the output of calculations in the excel file####
    #output_file = "D:/PY_FemGamer/Thesis/03_Output/Ratios_VISA.xlsx"
    output_file = "D:/PY_FemGamer/Thesis/03_Output/Ratios_PayPal.xlsx"
    #output_file = "D:/PY_FemGamer/Thesis/03_Output/Ratios_UHG.xlsx"
    output_df.to_excel(output_file, index=False)

    print("Output saved to:", output_file)

    end_time= time.time()
    Program_Runtime= (end_time - start_time)/60
    print("Runtime: ", round(Program_Runtime, 2), "minutes")

####END####