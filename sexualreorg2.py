import pandas as pd
import xlwt
import xlrd
from xlutils.copy import copy as xl_copy
from datetime import datetime

df = pd.read_csv('scuba2.csv')
timepd = df['Time']
jobidpd = df["Job Id"]
objidpd = df['objid']
decisionpd = df['decision']
useridpd = df['Userid']
actoridpd = df['Actor Id']

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
            Sheet1.write(link_comments, columnNumber, each)
            link_comments += 1
        elif isinstance(each, list):
            Sheet1.write(link_comments, columnNumber, str(each))
            link_comments += 1
        elif isinstance(each, int):
            Sheet1.write(link_comments, columnNumber, each)
            link_comments += 1
        else:
            Sheet1.write(link_comments, columnNumber, str(each))
            link_comments += 1 
    wb.save(fileName)

time = []
jobid = []
objid = []
decision = []
userid = []
actorid = []

for i in range(0,len(timepd)):
    time.append(timepd[i])
    jobid.append(jobidpd[i])
    objid.append(objidpd[i])
    userid.append(useridpd[i])
    actorid.append(actoridpd[i])

for each in decisionpd:
    decision.append(each)

decisionnew = []
for l in decision:
    decisionnew.append([x.strip() for x in l[10:-3].split('","')])
    

for n in range(0, len(time)):
    time[n] = int(time[n])

date_strings = [datetime.utcfromtimestamp(d).strftime('%Y-%m-%d %H:%M:%S') for d in time]


# {tagnamestr: {"count": num, "id": []}}
# result = {}

# count = 0
# for tree in decisionnew:
#     for tag in tree:
#         if tag not in result:
#             result[tag] = {"count": 0, "id": []}
#         else:
#             result[tag]["count"] += 1
#             result[tag]["id"].append(objid[count])
#             result[tag]["id"] = list(set(result[tag]["id"]))
#     count += 1

fileName = "brendon2"
sheetName = "15April2019"

output(fileName + ".xls",sheetName)
# open existing workbook
rb = xlrd.open_workbook(fileName + ".xls", formatting_info=True)
# make a copy of it
wb = xl_copy(rb)
wb.save(fileName+".xls")

# sheet1col1TagNames = ['tagname'] + [tagname for tagname in result]
# sheet1_col2_num3 = ['count'] + [value['count'] for key, value in result.items()]
# sheet1_col3_jobs = ['objids'] + [value['id'] for key, value in result.items()]

save_column_in_excel(fileName + ".xls", 0, 0, ["Time"]+time)
save_column_in_excel(fileName + ".xls", 0, 1, ["TimePrettyfy"]+date_strings)
save_column_in_excel(fileName + ".xls", 0, 2, ["decision tree"]+decisionnew)
save_column_in_excel(fileName + ".xls", 0, 3, ["jobid"]+jobid)
save_column_in_excel(fileName + ".xls", 0, 4, ["objid"]+objid)
save_column_in_excel(fileName + ".xls", 0, 5, ["userid"]+userid)
save_column_in_excel(fileName + ".xls", 0, 6, ["actorid"]+actorid)

