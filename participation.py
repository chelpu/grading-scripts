import gspread
import sys
import os
import operator

def getSheet(name):
    try:
        sht = gc.open(name)
        return sht
    except gspread.SpreadsheetNotFound as e:
        print e   
        return e

def totalReviews(sht):
    wkst = sht.get_worksheet(0)
    students = wkst.col_values(2)

    for i in range(1, len(students)):
        if not (students[i] == None):
            uniqname = students[i].lower().strip()
            print uniqname
            if uniqname in part_dict:
                part_dict[uniqname] = part_dict[uniqname] + 1
            else:
                part_dict[uniqname] = 1

def totalWkshtReviews(wksht):
    students = wksht.col_values(2)

    for i in range(1, len(students)):
        if not (students[i] == None):
            uniqname = students[i].lower().strip()
            print uniqname
            if uniqname in part_dict:
                part_dict[uniqname] = part_dict[uniqname] + 1
            else:
                part_dict[uniqname] = 1


name = 'Company Presentation Peer Evaluation (Responses)'

gc = gspread.Client(auth=(os.environ['EMAIL'], os.environ['PASSWD']))
gc.login()

# result sheet
participation_sht = gc.open('Feedback Contributions')
participation_wksht = participation_sht.get_worksheet(0)
part_dict = {}

# company presos
sht = getSheet(name)
totalReviews(sht)

name = 'MVP 1 Preso Peer Reviews'

# MVP 1 presos
sht = getSheet(name)
totalReviews(sht)

# redo presos
name = 'Peer Feedback 2/18'
sht = getSheet(name)
totalReviews(sht)

name = 'Project Presentations 2.23.15'
sht = getSheet(name)
totalReviews(sht)

name = 'Project Presentation Peer Evaluation (Responses)'
sht = getSheet(name)
wkshts = sht.worksheets()
for i in range(0, len(wkshts)):
    wksht = wkshts[i]
    totalWkshtReviews(wksht)

print "-------------results-------------\n"
sorted_dict = sorted(part_dict.items(), key=operator.itemgetter(1))
for tup in sorted_dict:
    print tup[0] , "\t", tup[1]