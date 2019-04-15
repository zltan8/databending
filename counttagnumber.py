import xlrd
import re
from collections import Counter

def read_column_in_excel(workbookName,sheetName,columnNumber):
    columnList=[]
    rb = xlrd.open_workbook(workbookName)
    sheet = rb.sheet_by_name(sheetName)
    row_count = len(sheet.col_values(0))
    row_read=0
    while row_read < row_count:
        each=sheet.cell(row_read,columnNumber).value
        columnList.append(each)
        row_read+=1
    #return list of elements in column
    return columnList

def splitnonalpha(s):
   pos = 1
   while pos < len(s) and s[pos].isalpha():
      pos+=1
   return (s[:pos], s[pos:])

tagnamearr = []
# {tagname: {jobs:[],count:0}, ...}
resultdict = {}

if __name__ == "__main__":
    jobsarr = read_column_in_excel('book2.xlsx', 'Sheet1', 0)
    jobsarr = [int(i) for i in jobsarr]

    count = 0
    for line in read_column_in_excel('book2.xlsx', 'Sheet1', 1):
        line = line.replace('"',"")
        line = line.replace("[","")
        line = line.replace("]","")
        line = line.replace("{","")
        line = line.replace("}","")
        linearr = line.split(", ")
        for tag in linearr:
            if tag in resultdict:
                resultdict[tag]["count"] += 1
                resultdict[tag]["jobs"].append(jobsarr[count])
            else:
                resultdict[tag] = {"jobs":[jobsarr[count]], "count":1}
        count += 1

    for tag in resultdict:
        resultdict[tag]["jobs"] = list(set(resultdict[tag]["jobs"]))

    print (resultdict)

    #     if isinstance(line, str):
    #         if 'suitable' in line.lower():
    #             result.append('suitable')
    #         # s = re.search('test_rep(.*),', line.lower())
    #         # if s:
    #         #     print (s.group(1).split(',')[0])
    #         elif 'test_rep' in line.lower():
    #             after = line.lower().split('test_rep')[1]
    #             result.append(splitnonalpha(after)[0])
    #         else:
    #             result.append("none met")
    #     else:
    #         result.append("not string")

    #     count += 1

    # print ("number of lines", count)

    with open('alanoutput.txt', 'w') as f:
        for item in resultdict:
            f.write("%s\t" % item)
            f.write("%s\t" % resultdict[item]["count"])
            f.write("%s\n" % resultdict[item]["jobs"])
