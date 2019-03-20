import json
from pprint import pprint
from collections import Counter
import xlwt
import xlrd
from xlutils.copy import copy as xl_copy

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
            Sheet1.write(link_comments, columnNumber, ",".join(each))
            link_comments += 1
        elif isinstance(each, int):
            Sheet1.write(link_comments, columnNumber, each)
            link_comments += 1
    wb.save(fileName)

with open('alan0.json') as f:
    data = json.load(f)

finaldict = {}
# finaldict structure: {tagname1 : { 3: {'num': num, 'jobs': [] }, 2: {'num': num, 'jobs': [] }, 1: {'num': num, 'jobs': [] }, tagname2: {}... }

with open('extra.txt') as f:
    extra = f.readlines()
extra = [x.strip() for x in extra]

processed = 0

for line in data:
    if line['fbid'] not in extra:
        print (line)
        tagdict = line['tag_map']
        taglist = []
        numtagseachjob = 0
        for key, value in tagdict.items():
            numtagseachjob += value
            key = key[1:-1].split(",") # converting from string to list
            for tag in key:
                count = value
                while count > 0:
                    taglist.append(tag)
                    count -= 1
        if numtagseachjob >= 3: # only want jobs with 3 or more tags
            processed += 1
            tagcount = Counter(taglist)
            for tagname in tagcount:
                if tagcount[tagname] >= 3:
                    if tagname in finaldict:
                        finaldict[tagname][3]['num'] += 1
                        finaldict[tagname][3]['jobs'].append(line['fbid'])
                    else:
                        finaldict[tagname] = { 3: {'num': 1, 'jobs': [line['fbid']] }, 2: {'num': 0, 'jobs': [] }, 1: {'num': 0, 'jobs': [] } }
                if tagcount[tagname] == 2:
                    if tagname in finaldict:
                        finaldict[tagname][2]['num'] += 1
                        finaldict[tagname][2]['jobs'].append(line['fbid'])
                    else:
                        finaldict[tagname] = { 3: {'num': 0, 'jobs': [] }, 2: {'num': 1, 'jobs': [line['fbid']] }, 1: {'num': 0, 'jobs': [] } }
                if tagcount[tagname] == 1:
                    if tagname in finaldict:
                        finaldict[tagname][1]['num'] += 1
                        finaldict[tagname][1]['jobs'].append(line['fbid'])
                    else:
                        finaldict[tagname] = { 3: {'num': 0, 'jobs': [] }, 2: {'num': 0, 'jobs': [] }, 1: {'num': 1, 'jobs': [line['fbid']] } }

# print (finaldict)

# removing duplicates in the jobslist in finaldict
for tagname in finaldict:
    for num in finaldict[tagname]:
        finaldict[tagname][num]['jobs'] = list(set(finaldict[tagname][num]['jobs']))
        # print (finaldict[tagname][num]['num'])
        # print (finaldict[tagname][num]['jobs'])

sheet1col1TagNames = [' '] + [tagname for tagname in finaldict]
sheet1_col2_num3 = ['3ormore'] + [value[3]['num'] for key, value in finaldict.items()]
sheet1_col3_jobs = ['jobslist'] + [value[3]['jobs'] for key, value in finaldict.items()]

sheet2col1TagNames = [' '] + [tagname for tagname in finaldict]
sheet2_col2_num3 = ['2'] + [value[2]['num'] for key, value in finaldict.items()]
sheet2_col3_jobs = ['jobslist'] + [value[2]['jobs'] for key, value in finaldict.items()]

sheet3col1TagNames = [' '] + [tagname for tagname in finaldict]
sheet3_col2_num3 = ['1'] + [value[1]['num'] for key, value in finaldict.items()]
sheet3_col3_jobs = ['jobslist'] + [value[1]['jobs'] for key, value in finaldict.items()]

fileName = "tagcounts"
sheetName = "3ormore"

output(fileName + ".xls",sheetName)
# open existing workbook
rb = xlrd.open_workbook(fileName + ".xls", formatting_info=True)
# make a copy of it
wb = xl_copy(rb)
# add sheet to workbook with existing sheets
Sheet2 = wb.add_sheet('2')
wb.save(fileName+".xls")
Sheet3 = wb.add_sheet('1')
wb.save(fileName+".xls")

save_column_in_excel(fileName + ".xls", 0, 0, sheet1col1TagNames)
save_column_in_excel(fileName + ".xls", 0, 1, sheet1_col2_num3)
save_column_in_excel(fileName + ".xls", 0, 2, sheet1_col3_jobs)

save_column_in_excel(fileName + ".xls", 1, 0, sheet2col1TagNames)
save_column_in_excel(fileName + ".xls", 1, 1, sheet2_col2_num3)
save_column_in_excel(fileName + ".xls", 1, 2, sheet2_col3_jobs)

save_column_in_excel(fileName + ".xls", 2, 0, sheet3col1TagNames)
save_column_in_excel(fileName + ".xls", 2, 1, sheet3_col2_num3)
save_column_in_excel(fileName + ".xls", 2, 2, sheet3_col3_jobs)

print ("Number of Jobs: ", len(data))
print ("processed", processed)

