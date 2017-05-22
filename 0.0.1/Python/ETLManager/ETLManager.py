import pandas as pd
import xlrd
from collections import OrderedDict
import simplejson as json
# Read from an excel file
sourceSDFile = "C:\\Users\\Brandon\Documents\\GitHub\\School-District-Surveys-Project\\Resources\\SchoolsCountyStateList.xls"
#pd.ExcelFile class returns a DataFrame
#xslx = pd.ExcelFile(sourceSDFile)
sheetName = "sdlist_1516"
#open the workbook
print("Starting: Reading File")
wb = xlrd.open_workbook(sourceSDFile)
print("Completed: Reading File")
#Primary Sheet
primarySheet = wb.sheet_by_name(sheetName)

# List to hold dictionaries
list = []
print("Starting: Processing data into a local list")
for rownum in range(2, primarySheet.nrows):
    item = OrderedDict()
    row_values = primarySheet.row_values(rownum)
    item["StateAbbr"] = row_values[0]
    item["StateFIPS"] = row_values[1]
    item["DistrictIDNumber"] = row_values[2]
    item["SchoolDistrictName"] = row_values[3]
    item["CountyNames"] = row_values[4]
    item["CountyFIPS"] = row_values[5]
    list.append(item)
print("Complete: Processing data into a local list")

print("Starting: Serializing list to JSON")
# Serialize the lsit of dicts to JSON
j = json.dumps(list)
print("Complete: Serializing list to JSON")

# Write to file 
print("Starting: Write JSON data to disk")
with open("C:\\Users\\Brandon\Documents\\GitHub\\School-District-Surveys-Project\\Resources\\data.json",'w') as f:
     f.write(j)

print("Complete: Write JSON data to disk")
    
#dataFrame = pd.read_excel(xslx, sheetName)
#print(dataFrame.head())
#StateAbbr(String),StateFIPS(INT),DistrictIDNumber(String),SchoolDistrictName(String)	
#CountyNames(String), CountyFIPS (String)

