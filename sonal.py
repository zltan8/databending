import xlrd
import re

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

result = []
count = 0
nsfa1 = []
nsfa2 = []
nsfa3 = []
result = []

if __name__ == "__main__":
    for line in read_column_in_excel('sonal.xlsx', 'Sheet2', 0):
        nsfa1.append(line.strip())

    for line in read_column_in_excel('sonal.xlsx', 'Sheet2', 1):
        if len(line.strip()) > 1:
            nsfa2.append(line.strip())
    for line in read_column_in_excel('sonal.xlsx', 'Sheet2', 2):
        if len(line.strip()) > 1:
            nsfa3.append(line.strip())

    for line in read_column_in_excel('sonal.xlsx', 'Animal', 10):
        line = line[2:-2].split('","')
        newline = []
        for t in line:
            newline.append(t.strip())
        if len([i for i in newline if i in nsfa1]) > 0:
            result.append('NSFA')
        else:
            if len([i for i in newline if i in nsfa2]+[i for i in newline if i in nsfa3])>1:
                result.append('NSFA')
            else:
                result.append('SFA')
    result = result[1:]
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

    with open('sonaloutput.txt', 'w') as f:
        for item in result:
            f.write("%s\n" % item)