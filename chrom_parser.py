import pandas as pd
import numpy as np
import myfunction
import math
import sys


df = pd.read_csv("D:\Margo\python\Chromatography\substances.csv")
files_list = myfunction.get_names("D:\Margo\python\Chromatography\data")

i = 0


for item in files_list:

    # save ID 
    index = item.find("xls")
    xls_id = item[index-5:index-1]
    i = i + 1
    df.at[i,"ID"] = xls_id
    
    # open Excel file and save data
    xls = pd.ExcelFile(item)
    sheet = xls.parse("Start")

    
    for j in range(len(sheet.index)):
        
        # Get data from chromatography worksheet
        substance = sheet.at[j, "Name"]
        value = sheet.at[j, "Fil"]
        
        # if [substance] is empty, omit it and continue
        if(pd.isnull(substance)):
            print(item, "contains empty row")
            continue

        # Check if substance is present in dataframe
        # if not, new column named [substance] added

        try:
            check = df.at[i,substance]
        except KeyError:
            df[substance] = np.nan
            print("new column added:", substance)



        # assign 0 for NaN values (in order to be able to perform next operations)
        if(math.isnan(df.at[i,substance])):
            df.at[i,substance] = 0
        # save new value and the old one (if value for [substance] has already been saved)            
        df.at[i,substance] = df.at[i,substance] + value 
    try:
        print(item, "COMPLETED")       
    except UnicodeEncodeError:
        print("Can't print filename")
        continue


df.to_csv("chrom_data_raw.csv")        
df = myfunction.clean_chrom_data(df) # remove items with ID containing letters and other symbols; 
df.to_csv("chrom_data_clean.csv")



