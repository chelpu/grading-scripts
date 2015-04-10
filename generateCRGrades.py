import gspread
import sys
import os

name = 'Company Presentation Peer Evaluation (Responses)'
person = raw_input("Whose feedback do you want? ")
score = raw_input("What was their score? ")

gc = gspread.Client(auth=(os.environ['EMAIL'], os.environ['PASSWD']))
gc.login()
try:
    sht = gc.open(name)
except gspread.SpreadsheetNotFound as e:
    print e

feedback_sht = gc.open('Company Report Feedback - ' + person)
feedback_wksht = feedback_sht.get_worksheet(0)
wkst = sht.get_worksheet(0)
first_row = wkst.row_values(1)

cells = wkst.findall(person)
num = 2
peer_score = 0
for i in range(0, len(cells)):
    if cells[i].col == 3:
        print cells[i].row
        feedback_row = wkst.row_values(cells[i].row)
        feedback_wksht.update_acell('A'+str(num), feedback_row[3])
        peer_score = peer_score + int(feedback_row[5])
        if not (feedback_row[4] == None):
            feedback_wksht.update_acell('B'+str(num), feedback_row[4])
        num = num + 1

peer_score = peer_score / float(num - 2)
feedback_wksht.update_acell('A'+str(num+2), 'Instructor Score: ' + str(score))
feedback_wksht.update_acell('B'+str(num+2), 'Peer Score: ' + str(peer_score))
