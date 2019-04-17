import pandas as pd
import xlwt
import xlrd
from xlutils.copy import copy as xl_copy
import csv

df = pd.read_csv('16Apr8am_17Apr8pm.csv')
timepd = df['Time']
objidpd = df['Object Id']
decisionpd = df['Decision Data']

def output(filename, sheet):
    book = xlwt.Workbook()
    sh = book.add_sheet(sheet)

    book.save(filename)

def save_column_in_excel(fileName,sheetNumber,columnNumber,columnList):
    rb = xlrd.open_workbook(fileName, formatting_info=True)
    wb = xl_copy(rb)
    Sheet1 = wb.get_sheet(sheetNumber)
    link_comments=0
    for each in columnList:
        if isinstance(each, str):
            while len(each) > 30000:
                Sheet1.write(link_comments, columnNumber, each[:30000])
                each = each[30000:]
                columnNumber += 1
            else:
                Sheet1.write(link_comments, columnNumber, each)
                link_comments += 1
        elif isinstance(each, list):
            # if its list of strings:
            taglist = ", ".join([str(tag) for tag in each])
            columnNumberedit = columnNumber
            if len(taglist) > 30000:
                while len(taglist) > 30000:
                    Sheet1.write(link_comments, columnNumberedit, taglist[:30000])
                    columnNumberedit += 1
                    taglist = taglist[30000:]
                Sheet1.write(link_comments, columnNumberedit, taglist[:30000])
                link_comments += 1
            else:
                Sheet1.write(link_comments, columnNumber, taglist)
                link_comments += 1

        elif isinstance(each, int):
            Sheet1.write(link_comments, columnNumber, each)
            link_comments += 1
    wb.save(fileName)

time = []
objid = []
decision = []

for each in timepd:
    time.append(each)

for each in objidpd:
    objid.append(each)

for each in decisionpd:
    decision.append(each)

decisionnew = []
for l in decision:
    decisionnew.append([x.strip() for x in l[10:-3].split('","')])
    
print (len(time),len(objid), len(decision))

# {tagnamestr: {"count": num, "id": []}}
result = {}

count = 0
for tree in decisionnew:
    for tag in tree:
        if tag not in result:
            result[tag] = {"count": 0, "id": []}
        else:
            result[tag]["count"] += 1
            result[tag]["id"].append(objid[count])
            result[tag]["id"] = list(set(result[tag]["id"]))
    count += 1

fileName = "16Apr8am_17Apr8pm_processed"
sheetName = "test"

output(fileName + ".xls",sheetName)
# open existing workbook
rb = xlrd.open_workbook(fileName + ".xls", formatting_info=True)
# make a copy of it
wb = xl_copy(rb)

sheet1col1TagNames = ['tagname'] + [tagname for tagname in result]
sheet1_col2_num3 = ['count'] + [value['count'] for key, value in result.items()]
sheet1_col3_jobs = ['objids'] + [value['id'] for key, value in result.items()]

save_column_in_excel(fileName + ".xls", 0, 0, sheet1col1TagNames)
save_column_in_excel(fileName + ".xls", 0, 1, sheet1_col2_num3)
save_column_in_excel(fileName + ".xls", 0, 2, sheet1_col3_jobs)

# rows = zip(sheet1col1TagNames, sheet1_col2_num3, sheet1_col3_jobs)

# with open(fileName+'.csv', "w") as f:
#     writer = csv.writer(f)
#     for row in sheet1col1TagNames:
#         writer.writerow(row)