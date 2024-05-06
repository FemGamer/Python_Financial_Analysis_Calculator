###combine stock anaylsis of 34 companies for statistical findings and regression analysis
##instead of 5 years 10 years of stock histroy will be exttracted from yahoo finance
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt



class Stock_Data:
    #calling the tickers with this function all at once               
    tickers=yf.Tickers ("POLICYBZR.NS PAYTM.NS POND-USD INST-USD CRV-EUR MKR-EUR BAL-EUR MATIC-EUR COMP5692-EUR YFI-EUR UNI7083-EUR SNX-EUR N26NB.NX SOFI HOOD AAVE-EUR COIN BCS CBK.DE KOTAKBANK.NS ZURN.SW BLK 0005.HK BNP.PA HDB SBIN.NS 3908.HK 2318.HK BRK-B DBK.DE GS JPM ALV.DE CS.PA")
    ##calling the historical data with .history function
    Stock_History={
        'POLICYBZR.NS': tickers.tickers['POLICYBZR.NS'].history(period="10y"),
        'PAYTM.NS': tickers.tickers['PAYTM.NS'].history(period="10y"),
        'POND-USD': tickers.tickers['POND-USD'].history(period="10y"),
        'INST-USD': tickers.tickers['INST-USD'].history(period="10y"),
        'CRV-EUR': tickers.tickers['CRV-EUR'].history(period="10y"),
        'MKR-EUR': tickers.tickers['MKR-EUR'].history(period="10y"),
        'BAL-EUR': tickers.tickers['BAL-EUR'].history(period="10y"),
        'MATIC-EUR': tickers.tickers['MATIC-EUR'].history(period="10y"),
        'COMP5692-EUR': tickers.tickers['COMP5692-EUR'].history(period="10y"),
        'YFI-EUR': tickers.tickers['YFI-EUR'].history(period="10y"),
        'UNI7083-EUR': tickers.tickers['UNI7083-EUR'].history(period="10y"),
        'SNX-EUR': tickers.tickers['SNX-EUR'].history(period="10y"),
        'N26NB.NX': tickers.tickers['N26NB.NX'].history(period="10y"),
        'SOFI': tickers.tickers['SOFI'].history(period="10y"),
        'HOOD': tickers.tickers['HOOD'].history(period="10y"),
        'AAVE-EUR': tickers.tickers['AAVE-EUR'].history(period="10y"),
        'COIN': tickers.tickers['COIN'].history(period="10y"),
        '0005.HK': tickers.tickers['0005.HK'].history(period="10y"),
        'BNP.PA': tickers.tickers['BNP.PA'].history(period="10y"),
        'HDB': tickers.tickers['HDB'].history(period="10y"),
        'SBIN.NS': tickers.tickers['SBIN.NS'].history(period="10y"),
        '3908.HK': tickers.tickers['3908.HK'].history(period="10y"),
        '2318.HK': tickers.tickers['2318.HK'].history(period="10y"),
        'BRK-B': tickers.tickers['BRK-B'].history(period="10y"),
        'DBK.DE': tickers.tickers['DBK.DE'].history(period="10y"),
        'GS': tickers.tickers['GS'].history(period="10y"),
        'JPM': tickers.tickers['JPM'].history(period="10y"),
        'ALV.DE': tickers.tickers['ALV.DE'].history(period="10y"),
        'CS.PA': tickers.tickers['CS.PA'].history(period="10y"),
        'BCS': tickers.tickers['BCS'].history(period="10y"),
        'CBK.DE': tickers.tickers['CBK.DE'].history(period="10y"),
        'KOTAKBANK.NS': tickers.tickers['KOTAKBANK.NS'].history(period="10y"),
        'ZURN.SW': tickers.tickers['ZURN.SW'].history(period="10y"),
        'BLK': tickers.tickers['BLK'].history(period="10y"),
    }

    stock_data= Stock_History
    # print(stock_data)

    df= pd.concat([df.assign(Ticker=ticker) for ticker, df in stock_data.items()])
    summary_stats= df.describe()##after getting , step of descriptive statistics was done
    #print(summary_stats)
    plot_df= summary_stats.plot.bar()
    save_plot= plot_df.get_figure()

    with pd.ExcelWriter('D:/PY_FemGamer/Thesis/04_Results/Stock_Combined_Analysis.xlsx') as writer:
        df.to_excel(writer, sheet_name='Historical Data', index=False)
        summary_stats.to_excel(writer, sheet_name='Summary Statistics', index=["Count", "Mean", "Std.", "Min", "25%", "50%","75%", "Max"])
        

        plt.figure(figsize=(10, 6))
        plot_df= summary_stats.plot.bar(ax=plt.gca())
        plt.title('Summary Statistics')
        plt.xlabel('Metrics')
        plt.ylabel('Values')
        plt.xticks(rotation=45)
        plt.tight_layout()
        save_plot = plt.gcf()
        save_plot.savefig("D:/PY_FemGamer/Thesis/04_Results/graph_1_1.png")

    plt.show()
    print("Run successfull ")
