import xlrd
from jira import JIRA
import getpass
import datetime
import sys
import os.path
import json
from os import path

jira_return = ""
jira_user = ""
while (jira_user == "" and jira_user != "q"):
    jira_user = input("Enter UserName: ")
if (jira_user == "q"):
    sys.exit()
jira_password = ""
if( jira_user == ""):
    jira_user = "ImportUser"
else:
    jira_password = getpass.getpass('Password:')

jira_server = "https://atlassian/jira"
jira_server = {'server': jira_server}
try:
    jira = JIRA(options=jira_server, basic_auth=(jira_user, jira_password))
except JIRAError as e:
    if e.status_code == 401:
        print("Login to JIRA failed. Check your username and password")
        sys.exit()

if len(sys.argv) > 1:
    ifile = sys.argv[1]
else:
    ifile = input("Please Drop File on Window. Then press Enter \n")
    #ifile = input("Please Drop File on EXE. Press Enter To Exit \n")
if (path.exists(ifile) == False):
    ifile = ifile.replace('"','')
    if (path.exists(ifile) == False):
        input("File does not exist, Press enter to exit")
        sys.exit()
workbook = xlrd.open_workbook(ifile)
try:
    worksheet = workbook.sheet_by_name('Q1025')
except:
    print("Sheet with name Q1025 does not exist. Exiting")
    input("")
    sys.exit()

key = worksheet.cell(2, 1).value
try:
    issue = jira.issue(key)
except JIRAError as e:
    print("Could Not find issue with key " + key + ". Please Check Excel File.")
    input("")
    sys.exit()
print("Updating Issue Type " + issue.fields.issuetype.name + " with key " + key)  

#do not search last row because it isn't needed.
num_of_rows = worksheet.nrows - 1
num_of_cols = worksheet.ncols
data = {}
part_number = ""
supplier = ""
row = 1
status = "Open"
sdescript  = ""
for row in range(1,num_of_rows):
    #Data is in only every other row so skip not needed rows
    #This block will need to be changed if the table is ever updated
    if (row % 2) == 0:
        continue
    for col in range(0,num_of_cols):
        celltype = worksheet.cell_type(row, col)
        #skip empty cells
        if ((celltype == 0) or (celltype == 6)):
            continue
        row_below = row+1
        celltype = worksheet.cell_type(row_below, col)
        value = 'null'
        value = worksheet.cell(row_below, col).value       
        #XL_CELL_DATE
        if (celltype == 3):
            #"%m/%d/%Y, %H:%M:%S"
            #value = datetime.datetime(*xlrd.xldate_as_tuple(value, workbook.datemode)).strftime('%Y-%m-%d') + 'T00:00:00+0000'
            #value = datetime.datetime(*xlrd.xldate_as_tuple(value, workbook.datemode)).strftime('%Y-%m-%dT%H:%M:%S+0000')
            value = datetime.datetime(*xlrd.xldate_as_tuple(value, workbook.datemode)).strftime('%Y-%m-%d')
            if (len(value) < 3):
                value = ""
        #EMPTY CELL
        elif ((celltype == 0) and (celltype == 6)):
            value = 'null'
        #XL_CELL_TEXT
        elif (celltype == 1):
            value = str(value)  
        else:
            value = str(value)
        text_value = str(worksheet.cell(row, col).value)
        cf = ""
        if (text_value == "Status"):
            status = value
            continue
        elif (text_value == "CAR #"):
            continue
        elif (text_value == "Part Number"):
            cf = 'customfield_10705'
            part_number = value
        elif (text_value == "Revision"):
            cf = 'customfield_11003'
        elif (text_value == "Part Description"):
            cf = 'customfield_11062'
        elif (text_value == "Supplier"):
            cf = 'customfield_11202'
            supplier = value
        elif (text_value == "Purchase Order Number"):
            cf = 'customfield_11215'
        elif (text_value == "Serial Numbers"):
            cf = 'customfield_10800'
        elif (text_value == "Corrective Action Due Date"):
            cf = 'customfield_12920'
        elif (text_value == "Ref Part #"):
            cf = 'customfield_11602'
        elif (text_value == "Quantity Rejected"):
            cf = 'customfield_13220'
            #Quantity Rejected is a number field and requires a number value
            value = float(value)
        elif (text_value == "RMA/NCMR#"):
            cf = 'customfield_13003'
        elif (text_value == "Description of Problem"):
            cf = 'description'
            sdescript = value
            #continue
        elif (text_value == "Containment Date"):
            cf = 'customfield_13202'
        elif (text_value == "Containment Action By"):
            cf = 'customfield_13319'
        elif (text_value == "Containment Action"):
            cf = 'customfield_12909'
        elif (text_value == "Root Cause Date"):
            cf = 'customfield_13207'
        elif (text_value == "Root Cause By"):
            cf = 'customfield_13320'
        elif (text_value == "Root Cause Type"):
            cf = 'customfield_12914'
            if (len(value) > 1):
                value = {'value':  value}
            else:
                #value = {"id": "-1"}
                value = {'value':  'Brainstorming'}
        elif (text_value == "Root Cause Analysis"):
            cf = 'customfield_12911'
        elif (text_value == "Corrective Action Date"):
            cf = 'customfield_13204'
        elif (text_value == "Corrective Action By"):
            cf = 'customfield_13321'
        elif (text_value == "Define and Implement Corrective Actions"):
            cf = 'customfield_12916'
        elif (text_value == "Verification Due Date"):
            cf = 'customfield_13336'
        elif (text_value == "Verification Date"):
            cf = 'customfield_13317'
        elif (text_value == "Verification By"):
            cf = 'customfield_13324'
        elif (text_value == "Verification of Corrective Action"):
            cf = 'customfield_13318'
        elif (text_value == "Preventative Action Date"):
            cf = 'customfield_13206'
        elif (text_value == "Preventative Action By"):
            cf = 'customfield_13322'
        elif (text_value == "Define and Implement Preventative Action"):
            cf = 'customfield_12917'
        elif (text_value == "Effectiveness Required"):
            cf = 'customfield_12918'
            if (len(value) > 1):
                value = {'value':  value}
            else:
                value = {"id": "-1"}
        elif (text_value == "Effectiveness Due Date"):
            cf = 'customfield_13335'
        elif (text_value == "Effectiveness Date"):
            cf = 'customfield_13210'
        elif (text_value == "Effectiveness by"):
            cf = 'customfield_13323'
        elif (text_value == "Effectiveness Evaluation"):
            cf = 'customfield_12919'
        #right now this is handled in JIRA, it is listed here if we want to keep track of it later
        elif (text_value == "CA Approved Date"):
            cf = ''
        elif (text_value == "CA Approval By"):
            cf = ''
        elif (text_value == "Final Approved Date"):
            cf = ''
        elif (text_value == "Final Approval By"):
            cf = ''
        elif (text_value == "Please submit form to Skurka Aerospace Inc."):
            continue
        else:
            if  (celltype != 0 and celltype != 6):
                #will show cells that couldn't be matched (someone updated the table without updating this script)
                print("Unmatched field with TEXT = " + text_value)
                print("Please contact Keith as the script must be updated.")
        if (cf != "" and value != ""):
            data[cf] = value
#print(data)
summ = "Supplier : " + supplier + " : " + part_number
data["summary"] = summ
print("Uploading Data, Please Wait")

issue.update(fields=data)

#print ("Number of Errors Occured = " + sret)
#print("Excel Status = " + status)
issuestatus = str(issue.fields.status)
#print("Current Issue Status = " + issuestatus)
if issuestatus == "Open":
    if (status == "Closed"):
        jira.transition_issue(issue, transition='11')
    elif (status == "Voided"):
        jira.transition_issue(issue, transition='31')
elif issuestatus == "Approvals":
    if (status == "Closed"):
        jira.transition_issue(issue, transition='11')
    elif (status == "Voided"):
        jira.transition_issue(issue, transition='31')
elif issuestatus == "Closed":
    if(status == "Open"):
        jira.transition_issue(issue, transition='21')

print("Upload Completed")  
