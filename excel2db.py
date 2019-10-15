import pandas as pd
import sqlite3
import re

cnx = sqlite3.connect('./sqlite.db')

xls = pd.ExcelFile('sample_data.xlsx')

# Create table for each sheet in excel
for sheetname in xls.sheet_names:
    df = pd.read_excel(xls,sheet_name=sheetname)

    for header in df.columns:
        print(df[header].dtype)
        if ((df[header].dtype == 'O' or df[header].dtype == 'int64') and df[header][0] != 0):
            # if data[0] is none, find until get one has data
            i = 0
            if (pd.isna(df[header][0])):
                print("Try to find data in blank column: " + header)
                for i in range(len(df.index)):
                    if (not pd.isna(df[header][i])):
                        print("Data found at: " + str(i))
                        break
            
            print("Matching date type for column: "+ header)
            # dd/mm/yyyy or dd/mm/yy <year range 200X to 202X
            x = re.search("^(0[1-9]|1[0-9]|2[0-9]|3[0-1])(.|-|\/)(0[1-9]|1[0-2])(.|-|\/)(20[0-2][0-9]|[0-2][0-9])$",str(df[header][i]))
            # mm/dd/yyyy or mm/dd/yy
            y = re.search("^(0[1-9]|1[0-2])(.|-|\/)(0[1-9]|1[0-9]|2[0-9]|3[0-1])(.|-|\/)(20[0-2][0-9]|[0-2][0-9])$",str(df[header][i]))
            # yyyymmdd or yyyy-mm-dd or yyyy.mm.dd
            z = re.search("^(20[0-2][0-9](.|-|/|)(0[1-9]|1[0-2])(.|-|/|)(0[1-9]|1[0-9]|2[0-9]|3[0-1]))$",str(df[header][i]))
 
            if (x != None or y != None or z != None ):
                print("Date field found: " + header)
                if (z != None):
                    # Process type Z for e.g. 20190930, convert to 2019-09-30 or otherwise pandas can't recognize as date type
                    if (len(str(df[header][i])) <= 8):
                        df[header] = df[header].astype(str).apply(lambda x: "{}{}{}{}{}".format(x[0:4],'-',x[4:6],'-',x[6:8]))
                # Convert column type to Datetime
                df[header] = pd.to_datetime(df[header])

    df.to_sql(name=sheetname, con=cnx)

