
from jira import JIRA
import sys
import json
#import ibm_db
import pyodbc
import logging
import logging.handlers
import os
import os.path
from pathlib import Path

logfile = os.path.expanduser('~\Documents\Logs\sqlimport.log')
logpath = os.path.expanduser('~\Documents\Logs')
if not os.path.isdir(os.path.expanduser('~\Documents')):
    os.mkdir('~\Documents')
if not os.path.isdir(os.path.expanduser(logpath)):
    os.mkdir(logpath)
if not os.path.isfile(logfile):
    Path(logfile).touch()




def main():
    jira_return = ""

    jira_password = ""
    jira_user = "taskuser"

    jira_server = "http://atlassian/jira"
    jira_server = {'server': jira_server}
    try:
        jira = JIRA(options=jira_server, basic_auth=(jira_user, jira_password))
    except JIRAError as e:
        if e.status_code == 401:
            print("Login to JIRA failed. Check your username and password")
            sys.exit(0)



    server = 'tcp:ATLASSIAN' 
    database = 'master' 
    username = 'jiramapics' 
    password = '' 
    cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
    cursor = cnxn.cursor()


    issuestoupdate = jira.search_issues("project = MRB and issuetype = 'NCMR' AND 'Purchase Order Number' is not empty", maxResults=5000)
    for issue in issuestoupdate:
    #issue = issue = jira.issue("MRB-5713")
        if (issue):
            print("Getting Data for " + issue.fields.issuetype.name + " with key " + issue.key)
            issuefields = issue.fields()
            sPo = getattr(issue.fields, "customfield_11215")
            sMo = getattr(issue.fields, "customfield_11053")
            sSource = getattr(issue.fields, "customfield_11051")
            print(sSource)
            bMo = False
            bPo = False
            bFill = True
            if (sSource == "Mfg Order" and sPo != None):
                bFill = False
            data = {}
            sqlquery = ""
            if (sPo != None):
                sPo=sPo.upper()
                sPo=sPo.replace(" ","")
                if (len(sPo) < 5):
                    sqlquery = ""
                else:
                    bPo = True;
                    sqlquery = "SELECT [ORDNO] FROM [MAPICS].[POWER9].[AMFLIBF].[POITEM] WHERE ORDNO = '"+ sPo + "';"
                    cursor.execute(sqlquery) 
                    row = cursor.fetchone()
                    if (row == None):
                        sqlquery = "SELECT [ITNBR],[ITDSC],[PACKC],[ENGNO],[BUYNO],[VNDNR],[QTYOR],[DKQTY],[HOUSE],[ACTPL],[EXTPR],[EXTPL],[WHSLC] FROM [MAPICS].[POWER9].[AMFLIBF].[POHISTI] WHERE ORDNO = '"+ sPo + "' AND MDATE > '1150000';"
                    else:
                        sqlquery = "SELECT [ITNBR],[ITDSC],[PACKC],[ENGNO],[BUYNO],[VNDNR],[QTYOR],[DKQTY],[HOUSE],[ACTPL],[EXTPR],[EXTPL],[WHSLC] FROM [MAPICS].[POWER9].[AMFLIBF].[POITEM] WHERE ORDNO = '"+ sPo + "';"
            if (sMo != None and sqlquery == ""):
                sMo=sMo.upper()
                sMo=sMo.replace(" ","")
                if (len(sMo) < 5):
                    sqlquery = ""
                else:
                    bMo = True;
                    sqlquery = "SELECT [ORDNO] FROM [MAPICS].[POWER9].[AMFLIBF].[MOMAST] WHERE ORDNO = '"+ sMo + "';"
                    cursor.execute(sqlquery) 
                    row = cursor.fetchone()
                    if (row == None):
                        sqlquery = "SELECT [FITEM],[FDESC],[ITRV],[ENGNO],[QTYRC],[FSKLC],[FITWH],[OPCUR],[WCCUR],[ORQTY],[QTDEV],[QCCUR],[QTSCP],[QTSPL] FROM [MAPICS].[POWER9].[AMFLIBF].[MOHMST] WHERE ORDNO = '"+ sMo + "' AND CRDT > '1150000';" 
                    else:
                        sqlquery = "SELECT [FITEM],[FDESC],[ITRV],[ENGNO],[QTYRC],[FSKLC],[FITWH],[OPCUR],[WCCUR],[ORQTY],[QTDEV],[QCCUR],[QTSCP],[QTSPL] FROM [MAPICS].[POWER9].[AMFLIBF].[MOMAST] WHERE ORDNO = '"+ sMo + "';"
            if (sqlquery != ""):
                sreturn = "";
                sBuyerNo = "";
                sVendorNo = "";
                cursor.execute(sqlquery) 
                row = cursor.fetchone()
                print(row)
                while row: 
                    if (bPo == True):
                        sItemNo = row[0]
                        sItemDesc = row[1]
                        sRev = row[2]
                        sEngNo = row[3]
                        sBuyerNo = row[4]
                        sVendorNo = row[5]
                    elif (bMo == True):
                        sItemNo = row[0]
                        sItemDesc = row[1]
                        sRev = row[2]
                        sEngNo = row[3]
                    if (bPo == True or bMo == True):
                        sreturn = sreturn + "Part Number:" + sItemNo + "\r\n"
                        sreturn = sreturn + "Part Description:" + sItemDesc + "\r\n"
                        sreturn = sreturn + "Part Revision:" + sRev + "\r\n"
                        sreturn = sreturn + "Engineering Number:" + sEngNo + "\r\n"
                        if (sItemNo.startswith("*") or sItemNo.startswith("OSP")):
                            bFill = False
                        if (bFill and 1==2):
                            data["customfield_10705"] = sItemNo
                            data["customfield_11062"] = sItemDesc
                            data["customfield_11003"] = sRev
                        data["customfield_14001"] = sEngNo
                    row = cursor.fetchone()
                if (sBuyerNo != ""):
                    sqlquery = "SELECT [BUYNM],[BUYPH],[EADR] FROM [MAPICS].[POWER9].[AMFLIBF].[BUYERF] WHERE BUYNO = "+ sBuyerNo + ";"
                    cursor.execute(sqlquery) 
                    row = cursor.fetchone()
                    print(row)
                    while row:
                        sBuyerName = row[0]
                        sBuyerPhone = row[1]
                        sBuyerEmail = row[2]
                        sreturn = sreturn + "Buyer Name:" + sBuyerName + "\r\n"
                        sreturn = sreturn + "Buyer Phone:" + sBuyerPhone + "\r\n"
                        sreturn = sreturn + "Buyer Email:" + sBuyerEmail + "\r\n"
                        row = cursor.fetchone()
                if (sVendorNo != None):
                    sqlquery = "SELECT [VNAME],[VADD1],[VADD2],[VCITY],[VSTAC],[VZIPC],[VETEL],[FAXTN] FROM [MAPICS].[POWER9].[AMFLIBF].[VENNAM] WHERE VNDNR = '"+ sVendorNo + "';"
                    cursor.execute(sqlquery) 
                    row = cursor.fetchone()
                    print(row)
                    while row:
                        sVendorName = row[0]
                        if (row[2] != ''):
                            sVendorAddress = row[0] + "\r\n" + row[1] + "\r\n" + row[2] + "\r\n" + row[3] + ", " + row[4] + " " + row[5]
                        else:
                            sVendorAddress = row[0] + "\r\n" + row[1] + "\r\n" + row[3] + ", " + row[4] + " " + row[5]
                        sVendorPhone = row[6]
                        sVendorFax = row[7]
                        sreturn = sreturn + "Vendor Name:" + sVendorName + "\r\n"
                        sreturn = sreturn + "Vendor Address:" + sVendorAddress + "\r\n"
                        sreturn = sreturn + "Vendor Phone:" + sVendorPhone + "\r\n"
                        sreturn = sreturn + "Vendor Fax:" + sVendorFax + "\r\n"
                        data["customfield_11818"] = sVendorAddress + "\r\n" + sVendorPhone
                        data["customfield_11202"] = sVendorName
                        row = cursor.fetchone()
                if (sreturn != ""):
                    print("Updating Issue " + issue.key)
                    #data["customfield_14000"] = sreturn
                    #Clear Background Flag
                    #data["customfield_14501"] = None
                    jira_return = issue.update(fields=data)
                else:
                    sqlquery = "SELECT TOP 1 [BUYNO] FROM [MAPICS].[POWER9].[AMFLIBF].[BUYERF]"
                    cursor.execute(sqlquery) 
                    row = cursor.fetchone()
                    if (row != None):
                        #NO DATA FOUND???? Clear Background Flag
                        print("No Data Found For Issue " + issue.key)
                        #data["customfield_14501"] = None
                        jira_return = issue.update(fields=data)
                    

    #Start to Add Rework Barcode
    issuestoupdate = jira.search_issues("project = MRB AND issuetype = Defect AND Disposition = Rework AND 'Rework Barcode' IS EMPTY AND 'Rework MO' IS NOT EMPTY AND 'Rework Operation' IS NOT EMPTY AND updated  >= -24h", maxResults=5000)
    for issue in issuestoupdate:
    #issue = issue = jira.issue("MRB-5032")
    #if (True):
        if (issue):
            print("Updating Rework Barcode for " + str(issue))
            data = {}
            sqlquery = ""
            sRwkMo = getattr(issue.fields, "customfield_11207")
            sRwkOp = getattr(issue.fields, "customfield_14600")
            sMo = getattr(issue.fields, "customfield_11053")
            if (sRwkMo != None):
                sRwkMo=sRwkMo.upper()
                sRwkMo=sRwkMo.replace(" ","")
                if (len(sRwkMo) < 5):
                    sqlquery = ""
                else:
                    sqlquery = "SELECT [TURNA],[TURNN],[TURNC] FROM [MAPICS].[POWER9].[AMFLIBF].[MOROUT] WHERE ORDNO = '"+ sRwkMo + "' AND OPSEQ = '"+ sRwkOp + "';" 
            if (sqlquery == "" and sMo != None):
                sMo=sMo.upper()
                sMo=sMo.replace(" ","")
                if (len(sMo) < 5):
                    sqlquery = ""
                else:
                    sqlquery = "SELECT [TURNA],[TURNN],[TURNC] FROM [MAPICS].[POWER9].[AMFLIBF].[MOROUT] WHERE ORDNO = '"+ sMo + "' AND OPSEQ = '"+ sRwkOp + "';" 
            if (sqlquery != ""):
                cursor.execute(sqlquery) 
                row = cursor.fetchone()
                sCode = ""
                while row:
                    print(row)
                    sCode = str(row[0])+str(row[1])+str(row[2])
                    row = cursor.fetchone()
                if (sCode != ""):
                    sBarcode = "<a href=\"http://atlassian/jira/browse/"+issue.key+"\"><img src=\"http://atlassian/barcode.php?code="+sCode+"\" alt=\"NCMR Barcode\" border=\"0\"></a>"
                    data["customfield_14601"] = sCode
                    data["customfield_14700"] = sBarcode
                    jira_return = issue.update(fields=data)

#Start Scrap Cost AutoFill
    issuestoupdate = jira.search_issues("project = MRB AND issuetype = Defect AND Disposition = Scrap AND 'Scrap Cost' IS EMPTY AND updated >= -24h")
    #issue = jira.issue("MRB-5035")
    #if (True):
    for issue in issuestoupdate:
        if (issue):
            print("Updating Scrap Cost for " + str(issue))
            data = {}
            sqlquery = ""
            parentissueid = (issue.fields.parent, issue.key)
            parentissue = jira.issue(parentissueid)
            sMo = getattr(parentissue.fields, "customfield_11053")
            sPo = getattr(parentissue.fields, "customfield_11215")
            dQuantity = getattr(issue.fields, "customfield_13220")
            #print(dQuantity)
            if (dQuantity != None):
                dQuantity = int(dQuantity)
            bMo = False
            bPo = False
            if (sPo != None):
                sPo=sPo.upper()
                sPo=sPo.replace(" ","")
                if (len(sPo) < 5):
                    sqlquery = ""
                else:
                    bPo = True;
                    sqlquery = "SELECT [ORDNO] FROM [MAPICS].[POWER9].[AMFLIBF].[POITEM] WHERE ORDNO = '"+ sPo + "';"
                    cursor.execute(sqlquery) 
                    row = cursor.fetchone()
                    if (row == None):
                        sqlquery = "SELECT [EXTPR],[ACTQY] FROM [MAPICS].[POWER9].[AMFLIBF].[POHISTI] WHERE ORDNO = '"+ sPo + "';"
                    else:
                        sqlquery = "SELECT [EXTPR],[ACTQY] FROM [MAPICS].[POWER9].[AMFLIBF].[POITEM] WHERE ORDNO = '"+ sPo + "';"
            if (sMo != None and sqlquery == ""):
                sMo=sMo.upper()
                sMo=sMo.replace(" ","")
                if (len(sMo) < 5):
                    sqlquery = ""
                else:
                    bMo = True;
                    sqlquery = "SELECT [ORDNO] FROM [MAPICS].[POWER9].[AMFLIBF].[MOMAST] WHERE ORDNO = '"+ sMo + "';"
                    cursor.execute(sqlquery) 
                    row = cursor.fetchone()
                    if (row == None):
                        sqlquery = "SELECT [CSTPC] FROM [MAPICS].[POWER9].[AMFLIBF].[MOHMST] WHERE ORDNO = '"+ sMo + "' AND CRDT > '1150000';" 
                    else:
                        sqlquery = "SELECT [CSTPC],[SETCO],[LABCO],[OVHCO],[ISSCO],[RECCO],[ORQTY],[QTYRC] FROM [MAPICS].[POWER9].[AMFLIBF].[MOMAST] WHERE ORDNO = '"+ sMo + "';"
            if (sqlquery != ""):
                cursor.execute(sqlquery) 
                row = cursor.fetchone()
                while row:
                    print(row)
                    dCost = None
                    if (bPo):
                        dTotalCost = row[0]
                        dTotalQuantity = row[1]
                        if (dTotalCost != None and dTotalQuantity != None):
                            dCost = dTotalCost
                            if (dQuantity != 0 and dTotalQuantity != 0):
                                dCost = dCost*dQuantity/dTotalQuantity;
                    elif (bMo):          
                        dUnitCost = row[0]
                        if (len(row) > 1):
                            dSetCost = row[1]
                            dLabCost = row[2]
                            dOverheadCost = row[3]
                            dIssueCost = row[4]
                            dReceiptCost = row[5]
                            dOrderQuantity = row[6]
                            dRecieptQuantity = row[7]
                            dAccountTotal = dIssueCost+dSetCost+dLabCost+dOverheadCost
                            if (dAccountTotal > dReceiptCost and dAccountTotal != 0):
                                dCost = dAccountTotal - dReceiptCost
                                if (dQuantity != 0 and dRecieptQuantity != 0):
                                    dTotalQuantity = dOrderQuantity - dRecieptQuantity
                                    if (dTotalQuantity != 0):
                                        dCost = dCost*dQuantity/dTotalQuantity
                        if dCost == None:
                            dCost = dUnitCost
                            if (dUnitCost != None and dUnitCost != 0):
                                dCost = dUnitCost*dQuantity
                    if (dCost != None):
                        data["customfield_13217"] = float(dCost)
                        jira_return = issue.update(fields=data)
                    row = cursor.fetchone()
    cursor.close()
    cnxn.close()

handler = logging.handlers.WatchedFileHandler(os.environ.get("LOGFILE", logfile))
formatter = logging.Formatter(logging.BASIC_FORMAT)
handler.setFormatter(formatter)
root = logging.getLogger()
root.setLevel(os.environ.get("LOGLEVEL", "INFO"))
logging.basicConfig(format='%(asctime)s %(message)s')
root.addHandler(handler)
 
try:
    main()
except Exception:
    logging.exception("Exception in main()")
    logging.exception(Exception)
    sys.exit(1)
except:
    logging.exception("Exception in main()")
    sys.exit(1)
sys.exit()
