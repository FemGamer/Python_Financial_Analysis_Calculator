#for zurich, allz
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
page_keywords_path = "D:/PY_FemGamer/keywords.txt" ##the hard-coding file containing the description of values to be called
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
        
        
###data extraction section
def extract_data(file_path):
    financial_data = {} ##puting all extracted data in a list
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()

            if page_keywords: ### if any keywords are defined
                found_keyword = is_page_relevant(text)
                if found_keyword: ### if any keyword was found
                    print("MATCH: keyword '" + found_keyword + "' found on page " + str(page.page_number)) ### print info to console
                else:
                    continue ### or jump direct to the next page
                print("")
                
            ###finding net profit from the data files
            Net_Profit_Value = re.search(r'Net income\s*(?:\(loss\))?\s*\(?([-,\d.]+)\)?', text, re.IGNORECASE)
            if not Net_Profit_Value:
                Net_Profit_Value = re.search(r'Net income\s*\(loss\)\s*\(([-,\d.]+)\)', text, re.IGNORECASE)
            if not Net_Profit_Value:
                Net_Profit_Value= re.search(r'Profit after tax\s*([-,\d.]+)|Profit for the year attributable to shareholder\s*([-,\d.]+)', text, re.IGNORECASE)
            if not Net_Profit_Value:
                Net_Profit_Value= re.search(r'PROFIT FOR THE YEAR\s*([-,\d.]+)', text, re.IGNORECASE)
            if not Net_Profit_Value:
                Net_Profit_Value= re.search(r'Profit for the year\s*([-,\d.]+)', text, re.IGNORECASE)
            if not Net_Profit_Value:
                Net_Profit_Value= re.search(r'(Net consolidated income after tax\s*[.:]*\s*\$?\s*([\d,]+))', text, re.IGNORECASE)
            if not Net_Profit_Value:
                Net_Profit_Value= re.search(r'Net income after taxes\s*([-,\d.]+)', text, re.IGNORECASE)
            if Net_Profit_Value:
                value = Net_Profit_Value.group(1).replace(",", "")
                try:
                    financial_data["Net profit"] = float(value)
                except ValueError:
                    print("Error converting '{}' to float. Skipping...".format(value))
              
            ###finding value if Liabilities from the data files
            Liabilities_Value = re.search(r'Total\s*liabilities\s*([\d,]+)', text, re.IGNORECASE)
            if Liabilities_Value:
                value = Liabilities_Value.group(1) or Liabilities_Value.group(2)
                value = value.replace(",", "")
                try:
                    financial_data["Total liabilities"] = float(value)
                except ValueError:
                    print("Error converting '{}' to float. Skipping...".format(value))

            ###finding shareholders equity value from the data files
            shareholders_equity= re.search(r'Shareholders(?:\’)? equity\s*([\d,]+)', text, re.IGNORECASE)#for zurich, allz 
            if shareholders_equity:
                value = shareholders_equity.group(1).replace(",", "")
            try:
                financial_data["Total Equity"] = float(value)
            except ValueError:
                    print("Error converting '{}' to float. Skipping...".format(value))

            ###finding dividends paif from the data files        
            dividends_paid= re.search(r'Dividends paid\s*\(?([\d,]+)\)?\s*\(?([\d,]+)\)?', text, re.IGNORECASE)
            if not dividends_paid:
                dividends_paid= re.search(r'Dividends paid to shareholders\s\(?([\d,]+)\)?\s*\(?([\d,]+)\)?', text, re.IGNORECASE)
            if dividends_paid:
                value1 = dividends_paid.group(1).replace(",", "")
                value2 = dividends_paid.group(2).replace(",", "")
                # Assuming the second value is the correct one if both are present
                value = value2 if value2 else value1
                try:
                    financial_data["Dividends paid"] = float(value)
                except ValueError:
                    print("Error converting '{}' to float. Skipping...".format(value))

            ###finding operating income from the data files
            Operating_Income_Value = re.search(r'(?:Operating income|OPERATING INCOME|Net operating income|Total revenue|Total income|Total revenues)\s*([\d,.-]+)', text)
            if Operating_Income_Value:
                value = Operating_Income_Value.group(1).replace(",", "").replace(".", "").replace("-", "")
                try:
                    financial_data["Operating income"] = float(value)
                except ValueError:
                    print("Error converting '{}' to float. Skipping...".format(value))

            ###finding total assets value from the data files
            Assets_Value = re.search(r'Total assets\s*(?:\(fair value\))?\s*([\d,]+)', text, re.IGNORECASE)
            if Assets_Value:
                value = Assets_Value.group(1).replace(",", "")
                try:
                    financial_data["Total assets"] = float(value)
                except ValueError:
                    print("Error converting '{}' to float. Skipping...".format(value))

    return financial_data

##financial ratios calculations step
def ratio_calculations(financial_data):
    Net_Profit = financial_data.get("Net profit")
    Operating_income = financial_data.get("Operating income")
    Liabilities = financial_data.get("Total liabilities")
    Assets = financial_data.get("Total assets")
    Equity= financial_data.get("Shareholder equity")
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
        Equity = float(Equity)
        Liabilities = float(Liabilities)
        DERatio = (Liabilities / Equity)
    else: DERatio = None

    if Equity and Liabilities:
        Equity = float(Equity)
        Liabilities = float(Liabilities)
        LERatio = (Equity / Liabilities)
    else: LERatio = None

    if Dividends_Paid and Net_Profit:
        Dividends_Paid = float(Dividends_Paid)
        Net_Profit = float(Net_Profit)
        PRatio = (Dividends_Paid / Net_Profit)
    else: PRatio = None
    

    return NPRatio, DARatio, DERatio, LERatio, PRatio,

##extracting the financial data and calculated ratios
def extract_and_calculate(file_path):
    financial_data = extract_data(file_path)
    NPRatio, DARatio, DERatio, LERatio, PRatio = ratio_calculations(financial_data)
    return financial_data, NPRatio, DARatio, DERatio, LERatio, PRatio,

def run(file_paths, batch_size=10):
    read_keywords(page_keywords_path)
    batch_count = 0
    excel_file_path = "D:/PY_FemGamer/Thesis/03_Output/Ratios_Zurich_Allianze.xlsx"  # Change this path as needed

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
                df.to_excel(writer, sheet_name=f'Sheet_{batch_count}', index=False)  # Write to separate sheets
                print(f"Batch {batch_count} exported to Sheet_{batch_count}")
            else:
                print(f"No financial data extracted for batch {batch_count}")

    print("Financial data exported to:", excel_file_path)



    end_time= time.time()
    Program_Runtime= (end_time - start_time)/60
    print("Runtime: ", round(Program_Runtime, 2), "minutes")



if __name__ == "__main__":
    file_paths = [

        "D:/PY_FemGamer/Thesis/02_DataBase/Allz_PDF/2013_01.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/Allz_PDF/2014_02.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/Allz_PDF/2015_03.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/Allz_PDF/2016_04.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/Allz_PDF/2017_05.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/Allz_PDF/2018_06.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/Allz_PDF/2019_07.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/Allz_PDF/2020_08.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/Allz_PDF/2021_09.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/Allz_PDF/2022_10.pdf",


        "D:/PY_FemGamer/Thesis/02_DataBase/Zurich_PDF/2014.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/Zurich_PDF/2015.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/Zurich_PDF/2016.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/Zurich_PDF/2017.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/Zurich_PDF/2018.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/Zurich_PDF/2019.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/Zurich_PDF/2020.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/Zurich_PDF/2021.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/Zurich_PDF/2022.pdf",
        "D:/PY_FemGamer/Thesis/02_DataBase/Zurich_PDF/2023.pdf",

        
        
        ]  # Modify this list with your desired file paths
    run(file_paths)


##structure is finally created after lots of trails and errors and understanding!#


