###few errors are reminaing but in all it works
###final pdf working file as i found proper pdf files with financial data
##datas in this script are mostly annual reports of centralised financial companies
import os.path
import pdfplumber
import pandas as pd
import re
import time
import os
from multiprocessing import Pool

start_time = time.time()
#file_path = "D:/PY_FemGamer/2018_06.pdf"
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
        
        
###in general i think this code could be more simplified, i am finding out but i have so little experience in dealing pdf file with python     

def extract_data(file_path):
    financial_data = {}
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()

            # Check if the page contains relevant keywords
            if page_keywords:
                found_keyword = is_page_relevant(text)
                if found_keyword:
                    print("MATCH: keyword '" + found_keyword + "' found on page " + str(page.page_number))
                else:
                    continue
                print("")

            # Extract net profit
            net_profit_match= re.search(r'Net Profit for the year\s*([-,\d.]+)', text, re.IGNORECASE)
            if net_profit_match:
                print("page: " + str(page.page_number) + "pattern: " + net_profit_match.group(1) + "     value: " + net_profit_match.group(1))
                value = net_profit_match.group(1).replace(",", "").replace(".", "")
                try:
                    financial_data["Net profit"] = float(value)
                except ValueError:
                    print("Error converting '{}' to float. Skipping...".format(value))

            # Extract operating income
            operating_income_match = re.search(r'(?:Interest|Total)\s*([\d,]+)', text, re.IGNORECASE)
            if operating_income_match:
                print("page: " + str(page.page_number) + "pattern: " + operating_income_match.group(1) + "     value: " + operating_income_match.group(1))
                value = operating_income_match.group(1).replace(",", "")
                try:
                    financial_data["Operating income"] = float(value)
                except ValueError:
                    print("Error converting '{}' to float. Skipping...".format(value))

            # Extract total liabilities
            total_liabilities_match = re.search(r'(?:Total liabilities|Total Liabilities|TOTAL)\s*([\d,]+)', text, re.IGNORECASE)
            if total_liabilities_match:
                print("page: " + str(page.page_number) + "pattern: " + total_liabilities_match.group(1) + "     value: " + total_liabilities_match.group(1))
                value = total_liabilities_match.group(1).replace(",", "").replace(".", "").replace("-", "")
                try:
                    financial_data["Total liabilities"] = float(value)
                except ValueError:
                    print("Error converting '{}' to float. Skipping...".format(value))

            # Extract total assets
            total_assets_match = re.search(r'(?:Total assets|TOTAL ASSETS|TOTAL|Total Assets)\s*([\d,.-]+)', text)
            if total_assets_match:
                print("page: " + str(page.page_number) + "pattern: " + total_assets_match.group(1) + "     value: " + total_assets_match.group(1))
                value = total_assets_match.group(1).replace(",", "").replace(".", "").replace("-", "")
                try:
                    financial_data["Total assets"] = float(value)
                except ValueError:
                    print("Error converting '{}' to float. Skipping...".format(value)),
    
            dividends_paid_match = re.search(r'Dividend paid\s*\(([-,\d.]+)\)', text, re.IGNORECASE)
            if not dividends_paid_match:
                dividends_paid_match = re.search(r'Dividend paid including tax thereon\s*\(([-,\d.]+)\)', text, re.IGNORECASE)
            if dividends_paid_match:
                value = dividends_paid_match.group(1) or dividends_paid_match.group(2)
                value = value.replace(",", "")
                try:
                    financial_data["Dividends paid"] = float(value)
                except ValueError:
                    print("Error converting '{}' to float. Skipping...".format(value))
            
            shareholder_equity= re.search(r'Capital\s*([\d,]+)', text, re.IGNORECASE)
            if shareholder_equity:
                value= shareholder_equity.group(1).replace(",", "")
                try:
                    financial_data["Equity"] = float(value)
                except ValueError:
                    print("Error converting '{}' to float. Skipping...".format(value))
            
           

    return financial_data


def ratio_calculations(financial_data):
    Net_Profit = financial_data.get("Net profit")
    Operating_income = financial_data.get("Operating income")
    Liabilities = financial_data.get("Total liabilities")
    Assets = financial_data.get("Total assets")

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

    return NPRatio, DARatio

def extract_and_calculate(file_path):
    financial_data = extract_data(file_path)
    NPRatio, DARatio = ratio_calculations(financial_data)
    return financial_data, NPRatio, DARatio

def run(file_paths, batch_size=10):
    read_keywords(page_keywords_path)
    batch_count = 0
    excel_file_path = "D:/PY_FemGamer/Thesis/03_Output/Ratios_SBI.xlsx"  # Change this path as needed

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
                    financial_data, NPRatio, DARatio = result_data
                    all_financial_data.append(financial_data)
                    # Append NPRatio and DARatio to financial data
                    financial_data["NPRatio"] = NPRatio
                    financial_data["DARatio"] = DARatio

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

        "D:/PY_FemGamer/Thesis/02_DataBase/SBI_PDF/31-03-2014.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/SBI_PDF/31-03-2015.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/SBI_PDF/31-03-2016.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/SBI_PDF/31-03-2017.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/SBI_PDF/31-03-2018.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/SBI_PDF/31-03-2019.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/SBI_PDF/31-03-2019.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/SBI_PDF/31-03-2020.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/SBI_PDF/31-03-2021.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/SBI_PDF/31-03-2022.pdf",
        


        ]  # Modify this list with your desired file paths
    run(file_paths)


##structure is finally created after lots of trails and errors and understanding!#