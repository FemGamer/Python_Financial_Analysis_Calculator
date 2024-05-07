###final pdf working file as i found proper pdf files with financial data
##datas in this script are mostly annual reports of centralised financial companies
## have issues with some values in net profit after tax are not getting extracted but now that for future
import os.path
import pdfplumber
import pandas as pd
import re
import time
import os
from multiprocessing import Pool

start_time = time.time()
page_keywords_path = "D:/PY_FemGamer/keywords.txt"
page_keywords = [] ### list with page keywords for filtering

### reads keyword file line by line and add it to list
def read_keywords(page_keywords_path):
    if not os.path.exists(page_keywords_path): ### checks if file exists
        print("No keyword-file exists!")
        print("")
        return False
    
    file = open(page_keywords_path, 'r')
    lines = file.readlines()
    if lines:
        print("Keywords:")
        for line in lines:
            line = line.strip()
            if line.startswith("#"): ### ignore comments in keyword file
                continue
            print(" - " + line)
            page_keywords.append(line)
    else:
        print("No keywords!")  
    print("")

### search the defined keywords in the page content and RETURNS THE FIRST MATCH!
def is_page_relevant(text):
    for keyword in page_keywords:
        if bool(re.search(keyword, text)):
            return keyword
        
        

def extract_data(file_path):
    financial_data = {}
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()

            if page_keywords:
                found_keyword = is_page_relevant(text)
                if found_keyword:
                    print("MATCH: keyword '" + found_keyword + "' found on page " + str(page.page_number))
                else:
                    continue
                print("")
            
            net_profit= re.search(r'Net consolidated income after tax\s*([\d,.-]+)', text, re.IGNORECASE)
            if net_profit:
                value = net_profit.group(1).replace(",", "")
                try:
                    financial_data["Net profit"] = float(value)
                except ValueError:
                    print("Error converting '{}' to float. Skipping...".format(value))

            revenue_patterns = [
                (r'Revenues from insurance activities\s*([\d,.-]+)', 'Revenues from insurance activities'),
                (r'Net revenues from banking activities\s*([\d,.-]+)', 'Net revenues from banking activities'),
                (r'Revenues from other activities\s*([\d,.-]+)', 'Revenues from other activities')
            ]

            operating_income = 0
            for pattern, category in revenue_patterns:
                revenue_match = re.search(pattern, text)
                if revenue_match:
                    value = revenue_match.group(1).replace(",", "").replace(".", "").replace("-", "")
                    try:
                        financial_data[category] = float(value)
                        operating_income += float(value)
                    except ValueError:
                        print("Error converting '{}' to float. Skipping...".format(value))

            financial_data["Total income"] = operating_income


            Liabilities_Value = re.search(r'TOTAL SHAREHOLDERS(?:\’)? EQUITY AND LIABILITIES\s*([\d,]+)', text, re.IGNORECASE)
            if Liabilities_Value:
                value = Liabilities_Value.group(1).replace(",", "")
                try:
                    financial_data["Total shareholders' equity and liabilities"] = float(value)
                except ValueError:
                    print("Error converting '{}' to float. Skipping...".format(value))

            Assets_Value = re.search(r'Total assets\s*([\d,]+)', text, re.IGNORECASE)
            if Assets_Value:
                value = Assets_Value.group(1).replace(",", "")
                try:
                    financial_data["Total assets"] = float(value)
                except ValueError:
                    print("Error converting '{}' to float. Skipping...".format(value))
            
            shareholder_equity= re.search(r'Total shareholders(?:\’)? equity\s*([\d,]+)', text, re.IGNORECASE)
            if shareholder_equity:
                value= shareholder_equity.group(1).replace(",", "")
                try:
                    financial_data["Equity"] = float(value)
                except ValueError:
                    print("Error converting '{}' to float. Skipping...".format(value))

            dividends_paid_match = re.search(r'Dividends payout\s*\(([-,\d.]+)\)', text, re.IGNORECASE)
            if dividends_paid_match:
                value = dividends_paid_match.group(1) or dividends_paid_match.group(2)
                value = value.replace(",", "")
                try:
                    financial_data["Dividends paid"] = float(value)
                except ValueError:
                    print("Error converting '{}' to float. Skipping...".format(value))



    return financial_data


def ratio_calculations(financial_data):
    Net_Profit = financial_data.get("Net profit")
    Operating_income = financial_data.get("Total income")
    Liabilities = financial_data.get("Total shareholders' equity and liabilities")
    Assets = financial_data.get("Total assets")
    Equity= financial_data.get("Equity")
    Dividends_Paid= financial_data.get("Dividends Paid")

    if Net_Profit and Operating_income:
        Net_Profit = float(Net_Profit)
        Operating_income = float(Operating_income)
        NPRatio = (Net_Profit / Operating_income) 
    else:
        NPRatio = None

    if Liabilities and Assets:
        Liabilities = float(Liabilities)
        Assets = float(Assets)
        DARatio = (Liabilities / Assets) 
    else:
        DARatio = None

    if Equity and Liabilities:
        Equity= float(Equity)
        Liabilities= float(Liabilities)
        DERatio= (Liabilities/Equity)
    else: DERatio= None

    if Equity and Liabilities:
        Equity= float(Equity)
        Liabilities= float(Liabilities)
        LERatio= (Equity/Liabilities)
    else: LERatio= None

    if Dividends_Paid and Net_Profit:
        Dividends_Paid = float(Dividends_Paid)
        Net_Profit = float(Net_Profit)
        PRatio = (Dividends_Paid/Net_Profit)
    else: PRatio = None

    return NPRatio, DARatio, DERatio, LERatio, PRatio

def extract_and_calculate(file_path):
    financial_data = extract_data(file_path)
    NPRatio, DARatio, DERatio, LERatio, PRatio = ratio_calculations(financial_data)
    return financial_data, NPRatio, DARatio, DERatio, LERatio, PRatio

def run(file_paths, batch_size=10):
    read_keywords(page_keywords_path)
    batch_count = 0
    excel_file_path = "D:/PY_FemGamer/Thesis/03_Output/Ratios_AXA.xlsx"  # Change this path as needed

    # Create a dummy DataFrame to ensure at least one sheet is visible
    dummy_df = pd.DataFrame({'Dummy': []})

    with pd.ExcelWriter(excel_file_path) as writer:  # Use pd.ExcelWriter to write to multiple sheets
        dummy_df.to_excel(writer, sheet_name='Dummy', index=False)

        for i in range(0, len(file_paths), batch_size):
            batch_count += 1
            batch_paths = file_paths[i:i + batch_size]
            batch_data = []

            with Pool(processes=10) as pool:
                for file_path in batch_paths:
                    batch_data.append(pool.apply_async(extract_and_calculate, args=(file_path,)))
                pool.close()
                pool.join()

            all_financial_data = []  # List to store all financial data for this batch
            for result in batch_data:
                result_data = result.get()
                if result_data:
                    financial_data, NPRatio, DARatio, DERatio, LERatio, PRatio = result_data
                    all_financial_data.append(financial_data)
                    # Append NPRatio and DARatio to financial data
                    financial_data["NPRatio"] = NPRatio
                    financial_data["DARatio"] = DARatio
                    financial_data["DERatio"] = DERatio
                    financial_data["LERatio"] = LERatio
                    financial_data["PRatio"]  = PRatio

            if all_financial_data:
                df = pd.DataFrame(all_financial_data)
                df.to_excel(writer, sheet_name=f'Sheet_{batch_count}', index=False)  # Write to separate sheet
                print(f"Batch {batch_count} exported to Sheet_{batch_count}")
            else:
                print(f"No financial data extracted for batch {batch_count}")

    print("Financial data exported to:", excel_file_path)


    end_time= time.time()
    Program_Runtime= (end_time - start_time)/60
    print("Runtime: ", round(Program_Runtime, 2), "minutes")    



if __name__ == "__main__":
    file_paths = [
                
    
        "D:/PY_FemGamer/Thesis/02_DataBase/AXA_PDF/2013.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/AXA_PDF/2014.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/AXA_PDF/2015.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/AXA_PDF/2016.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/AXA_PDF/2017.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/AXA_PDF/2018.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/AXA_PDF/2019.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/AXA_PDF/2020.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/AXA_PDF/2021.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/AXA_PDF/2022.pdf",

        ]  # Modify this list with your desired file paths
    run(file_paths)


##structure is finally created after lots of trails and errors and understanding!#