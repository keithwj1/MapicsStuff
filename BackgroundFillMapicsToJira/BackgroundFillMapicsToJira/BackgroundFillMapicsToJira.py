from jira import JIRA,JIRAError
import sys
import json
import pyodbc
import logging
import logging.handlers
import os
import os.path
from pathlib import Path


#from requests.utils import DEFAULT_CA_BUNDLE_PATH; print(DEFAULT_CA_BUNDLE_PATH)


sys.path.append(os.path.expanduser('~\source\repos\Mapics'))

from MapicsClasses import *

logfile = os.path.expanduser('~\Documents\Logs\sqlimport.log')
logpath = os.path.expanduser('~\Documents\Logs')
docpath = os.path.expanduser('~\Documents')
sqlautoitprogram = os.path.expanduser('~\Documents\AutoIt Script\sqlpropertiesopen.exe')
if not os.path.isdir(os.path.expanduser('~\Documents')):
    os.mkdir('~\Documents')
if not os.path.isdir(os.path.expanduser(logpath)):
    os.mkdir(logpath)
if not os.path.isfile(logfile):
    Path(logfile).touch()
       
class cMapicsData:
    def __init__(self):
        self.PurchaseOrderList = cPurchaseOrderList();
        self.ManufacturingOrderList = cManufacturingOrderList();
        self.RouterOppList = cRouterList();
        self.SupplierList = cSupplierList();
        self.ItemLocations = cItemLocationList();
        self.PartList = cPartList();
        self.BuyerList = cBuyerList();
        self.bFilled = False;
    def FillMapicsData(self):
        #data gets filled now when things are searched for
        if self.bFilled == False:
            self.TestMapicsConnection();
            #self.SupplierList.FillMapicsData();        
            #self.PurchaseOrderList.FillMapicsData()
            #self.ManufacturingOrderList.FillMapicsData()
            #self.RouterOppList.FillMapicsData();

            #We no longer fill EVERY part in the system, Only those that have had PO's Placed
            #self.PartList.FillList();
            #self.ItemLocations.FillMapicsData(self.PurchaseOrderList)
            #self.FillNewParts()
            #must be last-ish
            #self.AddMissingMapicsData()
            #self.bFilled = True
    #DEPRECIATED DO NOT USE
    def AddMissingMapicsData(self):       
        print("Begin Assigning Item Locations")
        for item in self.ItemLocations.aItemLocation:
            part = self.PartList.FindPart(item.sNumber);
            if part != None:
                part.ItemLocationList.aItemLocation.append(item);
        print("Begin Assigning Supplier Parts")
        for po in self.PurchaseOrderList.aPurchaseOrder:
            supplier = self.SupplierList.FindSupplier(po.sSupplierNumber)
            if (supplier != None):
                mappart = self.PartList.FindPart(po.sPartNumber)
                part = supplier.PartList.FindPart(po.sPartNumber)
                if part == None:
                    supplier.PartList.aPart.append(mappart)
                po.PartList.aPart.append(mappart)
                supplier.PurchaseOrderList.aPurchaseOrder.append(po);
    def FillNewParts(self):
        #Get New Parts List
        self.PurchaseOrderList.FillMapicsData()
        print("Begin Compiling New Parts List")
        for po in self.PurchaseOrderList.aPOOpen: 
            bIsNew = True
            for oldpo in self.PurchaseOrderList.aPOHistory:
                if (po.sPartNumber == oldpo.sPartNumber and po.sSupplierNumber == oldpo.sSupplierNumber):
                    bIsNew = False
                    break;
            if bIsNew:
                Supplier = self.SupplierList.FindSupplier(po.sSupplierNumber);
                SupplierPart = Supplier.PartList.FindPart(po.sPartNumber);
                SupplierPart.bNewPart = True;
    def TestMapicsConnection(self):
        server = 'tcp:ATLASSIAN' 
        database = 'master' 
        username = 'jiramapics' 
        password = '' 
        cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
        cursor = cnxn.cursor()
        sqlquery = "SELECT TOP 1 [BUYNO] FROM [MAPICS].[POWER9].[AMFLIBF].[BUYERF]"
        try:
            cursor.execute(sqlquery) 
        except:
            if cursor != None:
                cursor.close()
            if cnxn != None:
                cnxn.close()
            sqlautoitprogram = os.path.expanduser('~\Documents\AutoIt Script\sqlpropertiesopen.exe')
            os.system(sqlautoitprogram)
            logging.exception("Running SQL Login Program")
            sys.exit(0)
        if cursor != None:
            cursor.close()
        if cnxn != None:
            cnxn.close()
class cJiraData:
    def __init__(self):
        self.NCMRList = cNCMRList();
        self.SupplierList = cSupplierList();
    
    def FillJiraData(self):
        jira_return = ""

        jira_password = ""
        jira_user = "taskuser"

        jira_server = "https://atlassian/jira"
        jira_server = {'server': jira_server}
        try:
            jira = JIRA(options=jira_server, basic_auth=(jira_user, jira_password))
        except JIRAError as e:
            if e.status_code == 401:
                print("Login to JIRA failed. Check your username and password")
                sys.exit(0)    
    
        issuestocheck = jira.search_issues("project = MRB and issuetype = 'NCMR' and created >= -54w AND 'Purchase Order Number' is not empty AND 'Purchase Order Number' !~ 'N/A' AND Status != 'Void'", maxResults=5000)
        for issue in issuestocheck:
        #issue = jira.issue("MRB-5713")
        #if 1==1:
            if (issue):
                curKey = issue.key
                lastKey = ""
                sCreator = getattr(issue.fields, "reporter")
                ncmrcreated = str(issue.fields.created)[0:10]
                print("Getting Data for " + issue.fields.issuetype.name + " with key " + issue.key + " Created By " + str(sCreator))
                issuefields = issue.fields()
                sPo = str(getattr(issue.fields, "customfield_11215"))
                #sPo = sPo.replace(" ","")
                sPo = sPo.upper()
                sCAIssued = str(getattr(issue.fields, "customfield_11004"))
                sNCMRKeyStrings = curKey.split("-")
                nNCMRNumber = int(sNCMRKeyStrings[1])
                sPartNumber = str(getattr(issue.fields, "customfield_10705"))
                sPartDescription = str(getattr(issue.fields, "customfield_11062"))
                sSupplier = str(getattr(issue.fields, "customfield_11202"))
                sSupplierNumber = str(getattr(issue.fields, "customfield_15100"))
                try:
                    nQuantityRej = int(getattr(issue.fields, "customfield_13231"))
                    nQuantityRec = int(getattr(issue.fields, "customfield_13232"))
                except:
                    nQuantityRej = 0
                    nQuantityRec = 0
                curNCMR = self.NCMRList.AddNCMR(nNCMRNumber,sPartNumber,sPartDescription,sSupplier,sSupplierNumber,sCAIssued,sPo,ncmrcreated,sCreator,nQuantityRej,nQuantityRec)
                #get defect info
                for subtasks in issue.fields.subtasks:
                   subtaskID = subtasks.key
                   subtaskissue = jira.issue(subtasks.key)
                   #TryToGetFields
                   try:
                       sShouldBe = str(getattr(subtaskissue.fields, "customfield_11034"))
                       sActual = str(getattr(subtaskissue.fields, "customfield_11035"))
                   except:
                        sShouldBe = ""
                        sActual = ""
                   try:
                       nQuantityRej = int(getattr(issue.fields, "customfield_13220"))
                       nQuantityRec = int(getattr(issue.fields, "customfield_13219"))
                   except:
                       nQuantityRej = 0
                       nQuantityRec = 0
                   curNCMR.DefectList.AddDefect(sShouldBe,sActual,nQuantityRej,nQuantityRec)
                
                

class cCombineMapcisJira:
    def __init__(self):
        self.MapicsDataList = cMapicsData();
        self.JiraData = cJiraData();
    def FillMapicsData(self):
        self.MapicsDataList.FillMapicsData()
    def FillJiraData(self):
        self.JiraData.FillJiraData();
    def FillJiraMapicsData(self):      
        self.FillMapicsData()
        self.FillJiraData();
    def CombineData(self):
        print("Begin Combining Data")
        for ncmr in self.JiraData.NCMRList.aNCMR:
            allpo = self.MapicsDataList.PurchaseOrderList.FindAllPO(ncmr.sPurchaseOrder,True)
            if allpo == None:
                logging.warning("Incorrect PO for %s", "Key " + ncmr.sFullNumber + " Created By " + str(ncmr.sCreator) + " PO Number = " + ncmr.sPurchaseOrder)
            else:
                ncmr.sSupplierNumber = allpo[0].sSupplierNumber
                supplier = self.MapicsDataList.SupplierList.FindSupplier(ncmr.sSupplierNumber)
                if supplier != None:
                    supplier.NCMRList.aNCMR.append(ncmr)
                    ncmr.sSupplier = supplier.sName
                    for po in allpo:
                        part = supplier.PartList.FindPart(po.sPartNumber);
                        part.NCMRList.aNCMR.append(ncmr)
                        supplier.bHasNCMR = True;
                    if supplier.bHasNCMR:
                        jirasupplier = self.JiraData.SupplierList.FindSupplier(supplier.sNumber)
                        if jirasupplier == None:
                            self.JiraData.SupplierList.aSupplier.append(supplier)
                        else:
                            jirasupplier.bHasNCMR = True;

class cSystemFillJira:
    def __init__(self):
        self.MapicsData = cMapicsData();
    def GetSeparator(self,sString = ""):
        sSeparator = ""
        if "," in sString:
            sSeparator = ","
        elif ";" in sString:
            sSeparator = ";"
        elif "\\" in sString:
            sSeparator = "\\"
        elif "/" in sString:
            sSeparator = "/"
        elif "&" in sString:
            sSeparator = "&"
        elif " " in sString:
            sSeparator = " " 
        if sSeparator != "":
            return sSeparator
        else:
            return None
    def SeparateString(self,sString = ""):
        sString = sString.upper()
        sSeparator = self.GetSeparator(sString)
        if sSeparator != None:
            if sSeparator != " ":
                sString = sString.replace(" ","")
            return sString.split(sSeparator)
        else:
            return None
    def FillMapicsData(self):
        self.MapicsData.FillMapicsData()
    def main(self):
        #self.MapicsData.TestMapicsConnection();
        #self.JiraLogin();
        #self.FillMapicsData();
        self.AddToJira();
    def JiraLogin(self):
        jira_return = ""

        jira_password = ""
        jira_user = "taskuser"

        jira_server = "http://atlassian/jira"
        jira_server = {'server': jira_server}
        try:
            self.jira = JIRA(options=jira_server, basic_auth=(jira_user, jira_password))
        except JIRAError as e:
            if e.status_code == 401:
                print("Login to JIRA failed. Check your username and password")
                sys.exit(0)

    def AddToJira(self):
        self.JiraLogin()
        self.AddMapicsData()
        self.AddScrapCost()
        self.AddReworkBarcode()
    def AddReworkBarcode(self):
        #Start to Add Rework Barcode
        issuestoupdate = self.jira.search_issues("project = MRB AND issuetype = Defect AND Disposition = Rework AND 'Rework Barcode' IS EMPTY AND 'Rework MO' IS NOT EMPTY AND 'Rework MO' !~ 'N/A' AND 'Rework Operation' IS NOT EMPTY AND updated  >= -24h", maxResults=5000)
        for issue in issuestoupdate:
        #issue = issue = self.jira.issue("MRB-5032")
        #if (True):
            if (issue):
                #self.MapicsData.RouterOppList.FillMapicsData();
                print("Updating Rework Barcode for " + str(issue))
                data = {}
                sqlquery = ""
                sRwkMo = getattr(issue.fields, "customfield_11207")
                if (sRwkMo != None):
                    sRwkMo=sRwkMo.upper()
                sRwkOp = getattr(issue.fields, "customfield_14600")
                sMo = getattr(issue.fields, "customfield_11053")
                if (sRwkMo != None):
                    sRwkMo=sRwkMo.upper()
                    sRwkMo=sRwkMo.replace(" ","")
                    if (len(sRwkMo) > 5):
                        rwkrouter = self.MapicsData.RouterOppList.FindRouter(sRwkMo,sRwkOp)
                        if rwkrouter != None:
                            sCode = rwkrouter.sOppCode
                            if sCode != "":
                                sBarcode = "<a href=\"https://atlassian/jira/browse/"+issue.key+"\"><img src=\"https://atlassian/barcode.php?code="+sCode+"\" alt=\"NCMR Barcode\" border=\"0\"></a>"
                                data["customfield_14601"] = sCode
                                data["customfield_14700"] = sBarcode
                                jira_return = issue.update(fields=data)
                        
    def AddScrapCost(self):
        #Start Scrap Cost AutoFill
        issuestoupdate = self.jira.search_issues("project = MRB AND issuetype = Defect AND Disposition = Scrap AND 'Scrap Cost' IS EMPTY AND resolution = Unresolved")
        #issue = self.jira.issue("MRB-5035")
        #if (True):
        for issue in issuestoupdate:
            if (issue):     
                #self.MapicsData.PurchaseOrderList.FillMapicsData()
                #self.MapicsData.ManufacturingOrderList.FillMapicsData()
                print("Updating Scrap Cost for " + str(issue))
                data = {}
                sqlquery = ""
                parentissueid = (issue.fields.parent, issue.key)
                parentissue = self.jira.issue(parentissueid)
                sMo = getattr(parentissue.fields, "customfield_11053")
                sTopMo = getattr(parentissue.fields, "customfield_14500")
                sPo = getattr(parentissue.fields, "customfield_11215")
                if sMo == None:
                    sMo = sTopMo
                aPo = None
                aMo = None
                dQuantity = getattr(issue.fields, "customfield_13220")
                #print(dQuantity)
                if (dQuantity != None):
                    dQuantity = int(dQuantity)
                bMo = False
                bPo = False
                if (sPo != None):
                    sPo=sPo.upper()
                    #sPo=sPo.replace(" ","")
                    if (len(sPo) > 4):
                        aPo = self.SeparateString(sPo)
                        bPo = True;
                if (sMo != None):
                    sMo=sMo.upper()
                    #sMo=sMo.replace(" ","")
                    if (len(sMo) > 4):
                        aMo = self.SeparateString(sMo)
                        bMo = True;
                if (bPo or bMo):
                    dCost = None
                    if (bPo):
                        if aPo != None:
                            if len(aPo[0]) > 5:
                                sPo = aPo[0]
                        po = self.MapicsData.PurchaseOrderList.FindAllPO(sPo,True,True,True,self.MapicsData)
                        if po != None:
                            if po[0].sPartNumber.startswith("OSP") and bMo:
                                pass;
                            else:
                                dCost = po[0].GetScrapCost(dQuantity);
                    if bMo and dCost == None:        
                        if aMo != None:
                            if len(aMo[0]) > 5:
                                sMo = aMo[0]
                        mo = self.MapicsData.ManufacturingOrderList.FindAllMO(sMo,True,True,True)
                        if mo != None:
                            dCost = mo[0].GetScrapCost(dQuantity);
                    if (dCost != None):
                        data["customfield_13217"] = float(dCost)
                        jira_return = issue.update(fields=data)
    def AddMapicsData(self):
        issuestoupdate = self.jira.search_issues("project in (MRB,SRMRA,QCH,CA) and issuetype in ('NCMR',Task,'Corrective Action') and (('MAPICS Data Text Field' is empty AND updated >= -24h) OR 'SQL Background Update' ~ Yes)", maxResults=5000)
        #issuestoupdate = self.jira.search_issues("project in (MRB) and issuetype in ('NCMR')", maxResults=50000)
        for issue in issuestoupdate:
        #issue = issue = self.jira.issue("MRB-5031")
        #if 1==1:
            if (issue):
                data = {}
                issuestatus = str(issue.fields.status)
                bClear = False;
                sPo = getattr(issue.fields, "customfield_11215")
                sMo = getattr(issue.fields, "customfield_11053")
                sTopMo = getattr(issue.fields, "customfield_14500")
                if issue.fields.issuetype.name == 'Corrective Action':
                    sCAType = getattr(issue.fields, "customfield_12900")
                    if sCAType != "Supplier":
                        bClear = True
                        if sMo == None and sPo == None:
                            data["customfield_14000"] = " "
                if issuestatus == "Void" or bClear:
                    data["customfield_14501"] = None
                    jira_return = issue.update(fields=data)
                #self.MapicsData.SupplierList.FillMapicsData();        
                #self.MapicsData.PurchaseOrderList.FillMapicsData()
                #self.MapicsData.ManufacturingOrderList.FillMapicsData()
                #self.MapicsData.BuyerList.FillMapicsData();
                creator = getattr(issue.fields, "reporter")
                print("Getting Data for " + issue.fields.issuetype.name + " with key " + issue.key + " Created By " + str(creator))
                issuefields = issue.fields()
                if sTopMo != None:
                    sTopMo=sTopMo.replace(" ","")
                    sTopMo=sTopMo.upper()
                    if sMo != None:
                        sSeparator = self.GetSeparator(sMo)
                        if sSeparator == None:
                            sSeparator = ";"
                        sMo = sMo + sSeparator + sTopMo
                    else:
                        sMo = sTopMo
                aPo = None
                aMo = None
                sSource = getattr(issue.fields, "customfield_11051")
                print(sSource)
                bMo = False
                bPo = False
                bFill = True
                sFullPartList = ""
                sreturn = ""
                if (sPo != None):
                    #sPo=sPo.upper()
                    #sPo=sPo.replace(" ","")
                    if (len(sPo) > 5):
                        bPo = True;
                        aPo = self.SeparateString(sPo)
                if (sMo != None):
                    #sMo=sMo.upper()
                    #sMo=sMo.replace(" ","")
                    if (len(sMo) > 5):
                        bMo = True;
                        aMo = self.SeparateString(sMo)
                if (sSource == "Mfg Order" and sPo != None):
                    bFill = False
                if (bPo or bMo):
                    sreturn = "";
                    sBuyerNo = None;
                    sSupplierNo = None;
                    aOrder = None                  
                    if bMo:
                        if aMo == None:
                            aOrder = self.MapicsData.ManufacturingOrderList.FindAllMO(sMo,True,True,True);
                        else:
                            if aOrder == None:
                                aOrder = []
                            for Mo in aMo:
                                if (len(Mo) > 3):
                                    aCur = self.MapicsData.ManufacturingOrderList.FindAllMO(Mo,True,True,True);
                                    if aCur != None:
                                        aOrder=aOrder + aCur;
                                else:
                                    print(Mo)
                    if bPo:
                        if aPo == None:
                            if aOrder == None:
                                aOrder = self.MapicsData.PurchaseOrderList.FindAllPO(sPo,True,True,True,self.MapicsData);
                                if aOrder != None:
                                    sBuyerNo = aOrder[0].sBuyerNumber
                                    sSupplierNo = aOrder[0].sSupplierNumber
                            else:
                                aCur = self.MapicsData.PurchaseOrderList.FindAllPO(sPo,True,True,True,self.MapicsData); 
                                if aCur != None:
                                    sBuyerNo = aCur[0].sBuyerNumber
                                    sSupplierNo = aCur[0].sSupplierNumber
                                    aOrder=aOrder + aCur;
                        else:
                            if aOrder == None:
                                aOrder = []
                            for Po in aPo:
                                if (len(Po) > 3):
                                    aCur = self.MapicsData.PurchaseOrderList.FindAllPO(Po,True,True,True,self.MapicsData); 
                                    if aCur != None:
                                        sBuyerNo = aCur[0].sBuyerNumber
                                        sSupplierNo = aCur[0].sSupplierNumber
                                        aOrder=aOrder + aCur;
                    if aOrder != None:
                        data["customfield_15000"] = None
                        for oOrder in aOrder:
                            print("Using Order Number " + oOrder.sNumber)
                            sPartNumber = oOrder.sPartNumber
                            sPartDescription = oOrder.sPartDescription
                            sRevision = oOrder.sRevision
                            sEngNumber = oOrder.sEngNumber
                            if bPo:
                                sBuyerNo = oOrder.sBuyerNumber
                                sSupplierNo = oOrder.sSupplierNumber
                            sPartList = ""
                            if (sPartNumber != None):
                                if (sPartNumber != ""):
                                    sPartList = sPartList + "PartNumber!!" + sPartNumber + "=="
                            if (sPartDescription != None):
                                if (sPartDescription != ""):
                                    sPartList = sPartList + "PartDescription!!" + sPartDescription + "=="
                            if (sRevision != None):
                                if sRevision != "":
                                    sPartList = sPartList + "PartRevision!!" + sRevision + "=="
                            if (sEngNumber != None):
                                if (sEngNumber != ""):
                                    sPartList = sPartList + "EngNumber!!" + sEngNumber + "=="
                            if (sSupplierNo != None):
                                if (sSupplierNo != ""):
                                    sPartList = sPartList + "SupplierNumber!!" + sSupplierNo + "=="
                            if (sBuyerNo != None):
                                if (sBuyerNo != ""):
                                    sPartList = sPartList + "BuyerNumber!!" + sBuyerNo + "=="
                            if sPartList != "":
                                sPartList = sPartList[:-2]
                                sPartList = sPartList + ";;"
                                sFullPartList = sFullPartList + sPartList
                            sreturn = sreturn + "Part Number:" + sPartNumber + "\r\n"
                            sreturn = sreturn + "Part Description:" + sPartDescription + "\r\n"
                            sreturn = sreturn + "Part Revision:" + sRevision + "\r\n"
                            sreturn = sreturn + "Engineering Number:" + sEngNumber + "\r\n"
                    #15000 fill supplier part number field                  
                            data["customfield_15100"] = sSupplierNo
                        
                            if (sPartNumber.startswith("*") or sPartNumber.startswith("OSP")):
                                bFill = False
                            if (data["customfield_15000"] == None):
                                data["customfield_15000"] = sPartNumber
                            elif str(data["customfield_15000"]).startswith("*"):
                                data["customfield_15000"] = sPartNumber
                            if (bFill):
                                sJiraPartNumber = getattr(issue.fields, "customfield_10705")
                                if sJiraPartNumber == None:
                                    sJiraPartNumber = sPartNumber
                                if (sPartNumber == sJiraPartNumber):
                                    data["customfield_10705"] = sPartNumber
                                    data["customfield_11062"] = sPartDescription
                                    data["customfield_11003"] = sRevision
                                    data["customfield_14001"] = sEngNumber
                        #Clear SQL Background Flag
                        data["customfield_15000"] = None
                        if sBuyerNo != None:
                            buyer = self.MapicsData.BuyerList.FindBuyer(sBuyerNo);
                            if buyer != None:
                                sreturn = sreturn + "Buyer Name:" + buyer.sName + "\r\n"
                                sreturn = sreturn + "Buyer Phone:" + buyer.sPhone + "\r\n"
                                sreturn = sreturn + "Buyer Email:" + buyer.sEmail + "\r\n"
                                data["customfield_15503"] = buyer.sEmail
                        if sSupplierNo != None:
                            Supplier = self.MapicsData.SupplierList.FindSupplier(sSupplierNo)
                            if Supplier != None:
                                sreturn = sreturn + "Vendor Name:" + Supplier.sName + "\r\n"
                                sreturn = sreturn + "Vendor Address:" + Supplier.sAddress + "\r\n"
                                sreturn = sreturn + "Vendor Phone:" + Supplier.sPhone + "\r\n"
                                sreturn = sreturn + "Vendor Fax:" + Supplier.sFax + "\r\n"
                                sreturn = sreturn + "Vendor Email:" + Supplier.sEmail + "\r\n"
                                data["customfield_11818"] = Supplier.sAddress + "\r\n" + Supplier.sPhone
                                data["customfield_11202"] = Supplier.sName
                                data["customfield_16100"] = Supplier.sEmail
                                data["customfield_14401"] = Supplier.sEmailList

                if (sreturn != ""):
                    if sFullPartList != "":
                        sFullPartList = sFullPartList[:-2]
                        data["customfield_15800"] = sFullPartList
                    print("Updating Issue " + issue.key)
                    data["customfield_14000"] = sreturn
                    #Clear Background Flag
                    data["customfield_14501"] = None
                    jira_return = issue.update(fields=data)
def main():
    FillJiraSystem = cSystemFillJira();
    FillJiraSystem.main();

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
