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

with open('extra.txt') as f:
    extra = f.readlines()
extra = [x.strip() for x in extra]

# tagmap = {
#     "tagname": {
#         "target1": 0,
#         "target2": 0,
#         "size": 100
#         }
# }

tagmap = {}

processed = 0

for line in data:
    if line['fbid'] not in extra:
        tagdict = line['tag_map']
        taglist = []
        numtagseachjob = 0
        for key, value in tagdict.items():
            numtagseachjob += value
            key = key[1:-1].split(",") # converting from string to list
            for tag in key:
                count = value
                while count > 0:
                    taglist.append(tag.strip()) # removing front and back whitespace
                    count -= 1
        taglist_nodup = list(set(taglist))
        if len(taglist_nodup) > 1: # only want jobs with more than 1 distinct tag
            processed += 1
            tagcount = Counter(taglist)
            for tagname in tagcount:
                taglistremovetagname = [n for n in taglist if n != tagname]
                if tagname in tagmap:
                    tagmap[tagname]["size"] += tagcount[tagname]
                    for n in taglistremovetagname:
                        if n in tagmap[tagname]:
                            tagmap[tagname][n] += 1
                        else:
                            tagmap[tagname][n] = 1
                else:
                    tagmap[tagname] = {"size": tagcount[tagname]}

tagmaptop = {}
for tagname in tagmap:
    if tagmap[tagname]["size"] > 60:
        tagmaptop[tagname] = tagmap[tagname]

# print ("Number of Jobs: ", len(data))
# print ("processed", processed)

# converting to json

finaljson = {"nodes": [],"links": []}

# generate list from tagmap
alltags = [tag for tag in tagmaptop]

sourceC = 0
for tn in alltags:
    finaljson["nodes"].append({"group": 1, "name": tn, "size": tagmaptop[tn]["size"]})
    for tar in tagmaptop[tn]:
        if tar in alltags:
            finaljson["links"].append({"source": sourceC, "target": alltags.index(tar), "value": tagmaptop[tn][tar]})
    sourceC += 1

print (finaljson)



# with open('network.json', 'w') as outfile:
#     json.dump(finaljson, outfile)

# print (tagmap["test_obs_sexual_sidebreast_underbreast"])

# # print ("Number of Jobs: ", len(data))
# # print ("processed", processed)

# # converting to json

# finaljson = {"nodes": [],"links": []}

# # generate list from tagmap
# alltags = [tag for tag in tagmap]

# sourceC = 0
# for tn in alltags:
#     finaljson["nodes"].append({"group": 1, "name": tn, "size": tagmap[tn]["size"]})
#     for tar in tagmap[tn]:
#         if tar != "size":
#             finaljson["links"].append({"source": sourceC, "target": alltags.index(tar), "value": tagmap[tn][tar]})
#     sourceC += 1

# print (finaljson)



with open('network.json', 'w') as outfile:
    json.dump(finaljson, outfile)
