import json
from pprint import pprint
from collections import Counter
import xlwt
import xlrd
from xlutils.copy import copy as xl_copy

with open('alan0.json') as f:
    data = json.load(f)

# finaldict structure: {tagname1 : { 3: num, 2: num, 1: num}, tagname2 : { 3: num, 2: num, 1: num}...}
finaldict = {}

for line in data:
    tagdict = line['tag_map']
    taglist = []
    for key, value in tagdict.items():
        # print (key, value)
        key = key[1:-1].split(",") # converting from string to list
        # print (key)
        for tag in key:
            count = value
            while count > 0:
                taglist.append(tag)
                count -= 1
    tagcount = Counter(taglist)
    for tagname in tagcount:
        if tagcount[tagname] >= 3:
            if tagname in finaldict:
                finaldict[tagname][3] += 1
            else:
                finaldict[tagname] = { 3: 1, 2: 0, 1: 0}
        if tagcount[tagname] == 2:
            if tagname in finaldict:
                finaldict[tagname][2] += 1
            else:
                finaldict[tagname] = { 3: 0, 2: 1, 1: 0}
        if tagcount[tagname] == 1:
            if tagname in finaldict:
                finaldict[tagname][1] += 1
            else:
                finaldict[tagname] = { 3: 0, 2: 0, 1: 1}

print (finaldict)

col1TagNames = [' '] + [tagname for tagname in finaldict]
col2_num3 = ['3'] + [value[3] for key, value in finaldict.items()]
col3_num2 = ['2'] + [value[2] for key, value in finaldict.items()]
col4_num1 = ['1'] + [value[1] for key, value in finaldict.items()]

def output(filename, sheet):
    book = xlwt.Workbook()
    sh = book.add_sheet(sheet)

    book.save(filename)

fileName = "tagcounts"
sheetName = "tagsAndCounts"

output(fileName + ".xls",sheetName)

def save_column_in_excel(fileName,sheetNumber,columnNumber,columnList):
    rb = xlrd.open_workbook(fileName, formatting_info=True)
    wb = xl_copy(rb)
    Sheet1 = wb.get_sheet(sheetNumber)
    link_comments=0
    for each in columnList:
        Sheet1.write(link_comments, columnNumber, each)
        link_comments += 1
    wb.save(fileName)

save_column_in_excel(fileName+".xls", 0, 0, col1TagNames)
save_column_in_excel(fileName+".xls", 0, 1, col2_num3)
save_column_in_excel(fileName+".xls", 0, 2, col3_num2)
save_column_in_excel(fileName+".xls", 0, 3, col4_num1)

print ("Number of Jobs: ", len(data))