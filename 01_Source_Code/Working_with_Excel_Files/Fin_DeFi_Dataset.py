##  of fintech and defi companies extracted through yahoo finance
### 4 years of financial - cross refrenced with company websites
## 2nd part of it will be with excel and pdf as few companies have more 


import yfinance as yf
import pandas as pd
import time
from matplotlib.pyplot import*

start_time= time.time()
class ComProfiles:
    #calling the tickers with this function all at once                
    tickers=yf.Tickers("POLICYBZR.NS PAYTM.NS POND-USD INST-USD CRV-EUR MKR-EUR BAL-EUR MATIC-EUR COMP5692-EUR YFI-EUR UNI7083-EUR SNX-EUR N26NB.NX SOFI HOOD AAVE-EUR COIN ")
    ##calling the information of company profiles with '.info'
    Company_Profiles={
        'POLICYBZR.NS': tickers.tickers['POLICYBZR.NS'].info,
        'PAYTM.NS': tickers.tickers['PAYTM.NS'].info,
        'POND-USD': tickers.tickers['POND-USD'].info,
        'INST-USD': tickers.tickers['INST-USD'].info,
        'CRV-EUR': tickers.tickers['CRV-EUR'].info,
        'MKR-EUR': tickers.tickers['MKR-EUR'].info,
        'BAL-EUR': tickers.tickers['BAL-EUR'].info,
        'MATIC-EUR': tickers.tickers['MATIC-EUR'].info,
        'COMP5692-EUR': tickers.tickers['COMP5692-EUR'].info,
        'YFI-EUR': tickers.tickers['YFI-EUR'].info,
        'UNI7083-EUR': tickers.tickers['UNI7083-EUR'].info,
        'SNX-EUR': tickers.tickers['SNX-EUR'].info,
        'N26NB.NX': tickers.tickers['N26NB.NX'].info,
        'SOFI': tickers.tickers['SOFI'].info,
        'HOOD': tickers.tickers['HOOD'].info,
        'AAVE-EUR': tickers.tickers['AAVE-EUR'].info,
        'COIN': tickers.tickers['COIN'].info,
    }

    comp_prof= Company_Profiles
    df = pd.DataFrame(comp_prof.values()) #after data extraction transfering the output into dataframe for having it saved into excel file

    with pd.ExcelWriter("D:/PY_FemGamer/Thesis/02_DataBase/Company_Profiles.xlsx") as writer:
        df.to_excel(writer, sheet_name='Profiles', index=False)

    print("Data has been saved to excel file ")

class Income_Statement:
    tickers=yf.Tickers("POLICYBZR.NS PAYTM.NS POND-USD INST-USD CRV-EUR MKR-EUR BAL-EUR MATIC-EUR COMP5692-EUR YFI-EUR UNI7083-EUR SNX-EUR N26NB.NX SOFI HOOD AAVE-EUR COIN")
    ##calling the information of company profiles with '.income_stmt'
    Income_Statement={
        
        'SOFI': tickers.tickers['SOFI'].income_stmt,
        'HOOD': tickers.tickers['HOOD'].income_stmt,
        'AAVE-EUR': tickers.tickers['AAVE-EUR'].income_stmt,
        'COIN': tickers.tickers['COIN'].income_stmt,
        'POLICYBZR.NS': tickers.tickers['POLICYBZR.NS'].income_stmt,
        'PAYTM.NS': tickers.tickers['PAYTM.NS'].income_stmt,
        'POND-USD': tickers.tickers['POND-USD'].income_stmt,
        'INST-USD': tickers.tickers['INST-USD'].income_stmt,
        'CRV-EUR': tickers.tickers['CRV-EUR'].income_stmt,
        'MKR-EUR': tickers.tickers['MKR-EUR'].income_stmt,
        'BAL-EUR': tickers.tickers['BAL-EUR'].income_stmt,
        'MATIC-EUR': tickers.tickers['MATIC-EUR'].income_stmt,
        'COMP5692-EUR': tickers.tickers['COMP5692-EUR'].income_stmt,
        'YFI-EUR': tickers.tickers['YFI-EUR'].income_stmt,
        'UNI7083-EUR': tickers.tickers['UNI7083-EUR'].income_stmt,
        'SNX-EUR': tickers.tickers['SNX-EUR'].income_stmt,
        'N26NB.NX': tickers.tickers['N26NB.NX'].income_stmt,
        
    }

    with pd.ExcelWriter('D:/PY_FemGamer/Thesis/02_DataBase/Income_Statement.xlsx') as writer:
        for tickers, income_stmt in Income_Statement.items():
            #after data extraction transfering the output into dataframe for having it saved into excel file
            df0 = income_stmt.reset_index()  
            df0.to_excel(writer, sheet_name=tickers, index=False)
    
    print("Data has been saved to Excel file")

class Balance_Sheet:
    tickers=yf.Tickers("POLICYBZR.NS PAYTM.NS POND-USD INST-USD CRV-EUR MKR-EUR BAL-EUR MATIC-EUR COMP5692-EUR YFI-EUR UNI7083-EUR SNX-EUR N26NB.NX SOFI HOOD AAVE-EUR COIN ")
    ##calling the information of company profiles with '.balance_sheet'
    Balance_Sheet={
        
        'SOFI': tickers.tickers['SOFI'].balance_sheet,
        'HOOD': tickers.tickers['HOOD'].balance_sheet,
        'AAVE-EUR': tickers.tickers['AAVE-EUR'].balance_sheet,
        'COIN': tickers.tickers['COIN'].balance_sheet,
        'POLICYBZR.NS': tickers.tickers['POLICYBZR.NS'].balance_sheet,
        'PAYTM.NS': tickers.tickers['PAYTM.NS'].balance_sheet,
        'POND-USD': tickers.tickers['POND-USD'].balance_sheet,
        'INST-USD': tickers.tickers['INST-USD'].balance_sheet,
        'CRV-EUR': tickers.tickers['CRV-EUR'].balance_sheet,
        'MKR-EUR': tickers.tickers['MKR-EUR'].balance_sheet,
        'BAL-EUR': tickers.tickers['BAL-EUR'].balance_sheet,
        'MATIC-EUR': tickers.tickers['MATIC-EUR'].balance_sheet,
        'COMP5692-EUR': tickers.tickers['COMP5692-EUR'].balance_sheet,
        'YFI-EUR': tickers.tickers['YFI-EUR'].balance_sheet,
        'UNI7083-EUR': tickers.tickers['UNI7083-EUR'].balance_sheet,
        'SNX-EUR': tickers.tickers['SNX-EUR'].balance_sheet,
        'N26NB.NX': tickers.tickers['N26NB.NX'].balance_sheet,
        
    }

    with pd.ExcelWriter('D:/PY_FemGamer/Thesis/02_DataBase/Balance_Sheet.xlsx') as writer:
        for ticker, balance_sheet in Balance_Sheet.items():
            df1 = balance_sheet.reset_index()
            df1.to_excel(writer, sheet_name=ticker, index=False)
    
    print("Data has been saved to Excel file")

class Cash_Flows:
    tickers=yf.Tickers("POLICYBZR.NS PAYTM.NS POND-USD INST-USD CRV-EUR MKR-EUR BAL-EUR MATIC-EUR COMP5692-EUR YFI-EUR UNI7083-EUR SNX-EUR N26NB.NX SOFI HOOD AAVE-EUR COIN")
    ##calling the information of company profiles with '.cash_flow'
    CF_Statement={
        
        'SOFI': tickers.tickers['SOFI'].cash_flow,
        'HOOD': tickers.tickers['HOOD'].cash_flow,
        'AAVE-EUR': tickers.tickers['AAVE-EUR'].cash_flow,
        'COIN': tickers.tickers['COIN'].cash_flow,
        'POLICYBZR.NS': tickers.tickers['POLICYBZR.NS'].cash_flow,
        'PAYTM.NS': tickers.tickers['PAYTM.NS'].cash_flow,
        'POND-USD': tickers.tickers['POND-USD'].cash_flow,
        'INST-USD': tickers.tickers['INST-USD'].cash_flow,
        'CRV-EUR': tickers.tickers['CRV-EUR'].cash_flow,
        'MKR-EUR': tickers.tickers['MKR-EUR'].cash_flow,
        'BAL-EUR': tickers.tickers['BAL-EUR'].cash_flow,
        'MATIC-EUR': tickers.tickers['MATIC-EUR'].cash_flow,
        'COMP5692-EUR': tickers.tickers['COMP5692-EUR'].cash_flow,
        'YFI-EUR': tickers.tickers['YFI-EUR'].cash_flow,
        'UNI7083-EUR': tickers.tickers['UNI7083-EUR'].cash_flow,
        'SNX-EUR': tickers.tickers['SNX-EUR'].cash_flow,
        'N26NB.NX': tickers.tickers['N26NB.NX'].cash_flow,
        
    }

    with pd.ExcelWriter('D:/PY_FemGamer/Thesis/02_DataBase/Cash_Flow.xlsx') as writer:
        for ticker, cash_flow in CF_Statement.items():
            df2 = cash_flow.reset_index() #after data extraction transfering the output into dataframe for having it saved into excel file
            df2.to_excel(writer, sheet_name=ticker, index=False)
    
    print("Data has been saved to Excel file")
    

    end_time= time.time()
    Program_Runtime= (end_time - start_time)/60
    print("Runtime: ", round(Program_Runtime, 2), "minutes")
###functions work but one error is the naming convention for proper results that needs to be fixed

###finally  extraction is done and complete for the  from yahoo finance#
