import xlrd
import re
from collections import Counter

en = []
# 
en_species = {}

with open('endangered.txt', 'r') as f:
    for line in f:
        animal = re.sub(r'\([^)]*\)', '', line[4:-6].split('</td>')[0]).strip()
        en.append(animal)
        en_species

en = list(set(en))

acount = {}

for a in en:
    a = a.split(" ")[-1].lower()
    if a in acount:
        acount[a] += 1
    else:
        acount[a] = 1

print (acount)

for a in en:
    a = a.split(" ")[-1].lower()
    if a in acount:
        acount[a] += 1
    else:
        acount[a] = 1

with open('enutput.txt', 'w') as f:
    for item in acount:
        f.write("%s\t" % item)
        f.write("%s\n" % acount[item])