from jira import JIRA
import sys
import json
import pyodbc
import logging
import logging.handlers
import os
import os.path
from pathlib import Path
import datetime
from datetime import date

class cRouter:
    def __init__(self, sNumber = "",sOpp = "",sOppCode = ""):
        self.sNumber = sNumber
        self.sOpp = sOpp
        self.sOppCode = sOppCode

class cRouterList:
    def __init__(self):
        self.aRouter = []
        self.bFilled = False;
    def AddRouter(self,sNumber = "",sOpp = "",sOppCode = ""):
        newRouter = cRouter(sNumber,sOpp,sOppCode)
        self.aRouter.append(newRouter);
        return newRouter
    def FindRouter(self, sNumber,sOpp):
        self.FillMapicsData();
        for router in self.aRouter:
            if (router != None):
                if sNumber == router.sNumber and sOpp == router.sOpp:
                    return router
        return None
    def FillMapicsData(self):
        if self.bFilled == False:
            server = 'tcp:ATLASSIAN' 
            database = 'master' 
            username = 'jiramapics' 
            password = '' 
            cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
            cursor = cnxn.cursor()

            #Item Revision
            print("Begin Filling Router Data")
            sqlquery = "SELECT [ORDNO],[OPSEQ],[TURNA],[TURNN],[TURNC] FROM [MAPICS].[POWER9].[AMFLIBF].[MOROUT];"
            cursor.execute(sqlquery) 
            row = cursor.fetchone()
            while row:
                sOppCode = str(row[2]) + str(row[3]) + str(row[4])
                self.AddRouter(row[0],row[1],sOppCode)
                row = cursor.fetchone()
            cursor.close()
            cnxn.close()
            self.bFilled = True;

class cManufacturingOrder:
    def __init__(self,sNumber = "", sPartNumber = "",sPartDescription = "",sRevision = "",sEngNumber = "",nMakeDate = 0.0,dUnitCost = 0.0,dSetCost = 0.0,dLabCost = 0.0,dOverheadCost = 0.0,dIssueCost = 0.0,dReceiptCost = 0.0,nOrderQuantity = 0,nRecieptQuantity = 0,bHistory = False,sJobNumber = ""):
        self.sNumber = sNumber
        self.sPartNumber = sPartNumber
        self.sPartDescription = sPartDescription
        self.sRevision = sRevision
        self.sEngNumber = sEngNumber
        self.nMakeDate = nMakeDate
        self.dUnitCost = dUnitCost
        self.dSetCost = dSetCost
        self.dLabCost = dLabCost
        self.dOverheadCost = dOverheadCost
        self.dIssueCost = dIssueCost
        self.dReceiptCost = dReceiptCost
        self.nOrderQuantity = nOrderQuantity
        self.nRecieptQuantity = nRecieptQuantity
        self.bHistory = bHistory
        self.sJobNumber = sJobNumber
        self.PartList = cPartList()
        self.sBuyerNumber = None
        self.sSupplierNumber = None
        if (sRevision == None or sRevision == ""):
            if ("-" in sEngNumber):
                aEngNumber = sEngNumber.split("-")
                nLength = len(aEngNumber)
                sTemp = aEngNumber[nLength - 1]
                if len(sTemp) < 4:
                    self.sRevision = sTemp
            if ("_" in sEngNumber):
                aEngNumber = sEngNumber.split("_")
                nLength = len(aEngNumber)
                sTemp = aEngNumber[nLength - 1]
                if len(sTemp) < 4:
                    self.sRevision = sTemp
    def GetScrapCost(self,dQuantity = 0):
        dCost = None
        dAccountTotal = self.dIssueCost+self.dSetCost+self.dLabCost+self.dOverheadCost
        #only use unit cost below
        if (dAccountTotal > self.dReceiptCost and dAccountTotal != 0 and 1==2):
            dCost = dAccountTotal - self.dReceiptCost
            if (dQuantity != 0 and self.nRecieptQuantity != 0):
                dTotalQuantity = self.nOrderQuantity - self.nRecieptQuantity
                if (dTotalQuantity != 0):
                    dCost = dCost*dQuantity/dTotalQuantity
        if dCost == None:
            dCost = self.dUnitCost
            if (self.dUnitCost != None and self.dUnitCost != 0):
                dCost = self.dUnitCost*dQuantity
        return dCost;
class cManufacturingOrderList:
    def __init__(self):
        self.aManufacturingOrder = []
        self.aMOOpen = []
        self.aMOHistory = []
        self.aMOHistoryRecent = []
        self.bFilled = False;
        self.bHistFilled = False;
        self.bOpenFilled = False;
    def AddMO(self, sNumber = "", sPartNumber = "",sPartDescription = "",sRevision = "",sEngNumber = "",nMakeDate = 0.0,dUnitCost = 0.0,dSetCost = 0.0,dLabCost = 0.0,dOverheadCost = 0.0,dIssueCost = 0.0,dReceiptCost = 0.0,nOrderQuantity = 0,nRecieptQuantity = 0,bHistory = False,sJobNumber = ""):
        newMO = cManufacturingOrder(sNumber, sPartNumber,sPartDescription,sRevision,sEngNumber,nMakeDate,dUnitCost,dSetCost,dLabCost,dOverheadCost,dIssueCost,dReceiptCost,nOrderQuantity,nRecieptQuantity,bHistory,sJobNumber)
        self.aManufacturingOrder.append(newMO);
        if bHistory:
            self.aMOHistory.append(newMO);
            if (nMakeDate > 1140000.0 or nMakeDate == 0.0):
                self.aMOHistoryRecent.append(newMO);
        else:
            self.aMOOpen.append(newMO);
        return newMO
    def FindMO(self, sNumber,bOpen = True, bHistory = True, bRecentHistory = True):      
        if bOpen:
            self.FillMOOpen();
            for mo in self.aMOOpen:
                if (mo != None):
                    if (mo.sNumber == sNumber):
                        return mo;
        if bHistory:
            self.FillMOHistory();
            if bRecentHistory:
                for mo in self.aMOHistoryRecent:
                    if (mo != None):
                        if (mo.sNumber == sNumber):
                            return mo;
            else:
                for mo in self.aMOHistory:
                    if (mo != None):
                        if (mo.sNumber == sNumber):
                            return mo;
        return None
    def FindAllMO(self,sNumber,bOpen = True, bHistory = True, bRecentHistory = True):       
        mos = []
        if bOpen:
            self.FillMOOpen();
            for mo in self.aMOOpen:
                if (mo != None):
                    if (mo.sNumber == sNumber):
                        mos.append(mo);
        if bHistory:
            if len(mos) == 0 or 1==1:
                self.FillMOHistory();
                if bRecentHistory:
                    for mo in self.aMOHistoryRecent:
                        if (mo != None):
                            if (mo.sNumber == sNumber):
                                mos.append(mo);
                else:
                    for mo in self.aMOHistory:
                        if (mo != None):
                            if (mo.sNumber == sNumber):
                                mos.append(mo);
        if len(mos) > 0:
            return mos;
        else:
            return None
    def GetSpecificMOs(self,bHistory = False, bRecentHistory = True):       
        if bHistory:
            self.FillMOHistory();
            if bRecentHistory:
                return self.aMOHistoryRecent;
            else:
                return self.aMOHistory;
        else:
            self.FillMOOpen();
            return self.aMOOpen
    def RemoveMO(self, sNumber):
        curPo = self.FindMO(sNumber);
        del curPo
    def FillMOOpen(self, SupplierList = None,PartList = None):
        if self.bOpenFilled == False:
            server = 'tcp:ATLASSIAN' 
            database = 'master' 
            username = 'jiramapics' 
            password = 'FqAy&6KgqIIE6s$ya6uD' 
            cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
            cursor = cnxn.cursor()

            #Supplier Data
            #Fill PO DATA
            print("Begin Filling Open MO Data")
            sqlquery = "SELECT [FITEM],[FDESC],[ITRV],[ENGNO],[CSTPC],[SETCO],[LABCO],[OVHCO],[ISSCO],[RECCO],[ORQTY],[QTYRC],[ORDNO],[CRDT],[JOBNO] FROM [MAPICS].[POWER9].[AMFLIBF].[MOMAST];"
            cursor.execute(sqlquery) 
            row = cursor.fetchone()   
            while row:         
                sPartNumber = row[0]
                sPartDescription = row[1]
                sRevision = row[2]
                sEngNumber = row[3]
                dPieceCost = float(row[4])
                dSetCost = float(row[5])
                dLabCost = float(row[6])
                dOverheadCost = float(row[7])
                dIssueCost = float(row[8])
                dReceiptCost = float(row[9])
                nOrderQuantity = int(row[10])
                nRecieptQuantity = int(row[11])
                sNumber = row[12]
                nMakeDate = float(row[13])
                sJobNumber = str(row[14])
                bHistory = False
                mo = self.AddMO(sNumber, sPartNumber,sPartDescription,sRevision,sEngNumber,nMakeDate,dPieceCost,dSetCost,dLabCost,dOverheadCost,dIssueCost,dReceiptCost,nOrderQuantity,nRecieptQuantity,bHistory,sJobNumber)
                
                row = cursor.fetchone()
            cursor.close()
            cnxn.close()
            self.bOpenFilled = True
    def FillMOHistory(self, SupplierList = None,PartList = None):
        if self.bHistFilled == False:
            server = 'tcp:ATLASSIAN' 
            database = 'master' 
            username = 'jiramapics' 
            password = 'FqAy&6KgqIIE6s$ya6uD' 
            cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
            cursor = cnxn.cursor()

            #Supplier Data
            #Fill PO DATA
            print("Begin Filling History MO Data")
            sqlquery = "SELECT [FITEM],[FDESC],[ITRV],[ENGNO],[CSTPC],[SETCO],[LABCO],[OVHCO],[ISSCO],[RECCO],[ORQTY],[QTYRC],[ORDNO],[CRDT],[JOBNO] FROM [MAPICS].[POWER9].[AMFLIBF].[MOHMST];"
            cursor.execute(sqlquery) 
            row = cursor.fetchone()   
            while row:
                sPartNumber = row[0]
                sPartDescription = row[1]
                sRevision = row[2]
                sEngNumber = row[3]
                dPieceCost = float(row[4])
                dSetCost = float(row[5])
                dLabCost = float(row[6])
                dOverheadCost = float(row[7])
                dIssueCost = float(row[8])
                dReceiptCost = float(row[9])
                nOrderQuantity = int(row[10])
                nRecieptQuantity = int(row[11])
                sNumber = row[12]
                nMakeDate = float(row[13])
                sJobNumber = str(row[14])
                bHistory = True
                #def AddMO(self, sNumber = "", sPartNumber = "",sPartDescription = "",sRevision = "",sEngNumber = "",nMakeDate = 0.0,dUnitCost = 0.0,dSetCost = 0.0,dLabCost = 0.0,dOverheadCost = 0.0,dIssueCost = 0.0,dReceiptCost = 0.0,nOrderQuantity = 0,nRecieptQuantity = 0,bHistory = False,sJobNumber = ""):
                mo = self.AddMO(sNumber, sPartNumber,sPartDescription,sRevision,sEngNumber,nMakeDate,dPieceCost,dSetCost,dLabCost,dOverheadCost,dIssueCost,dReceiptCost,nOrderQuantity,nRecieptQuantity,bHistory,sJobNumber)

                row = cursor.fetchone()
            cursor.close()
            cnxn.close()
            self.bHistFilled = True
    def FillMapicsData(self,SupplierList = None,PartList = None):
       if self.bFilled == False:
           self.FillMOOpen(SupplierList,PartList);
           self.FillMOHistory(SupplierList,PartList);
           self.bFilled = True;
           self.bHistFilled = True;
           self.bOpenFilled = True;

class cBuyer:
    def __init__(self, sNumber = "",sName = "",sEmail = "", sPhone = "", sFax = ""):
        self.sNumber = sNumber
        self.sName = sName
        self.sEmail = sEmail
        self.sPhone = sPhone
        self.sFax = sFax

class cBuyerList:
    def __init__(self):
        self.aBuyer = []
        self.bFilled = False;
    def AddBuyer(self,sNumber = "",sName = "",sEmail = "", sPhone = "", sFax = ""):
        newBuyer = cBuyer(sNumber,sName,sEmail, sPhone, sFax)
        self.aBuyer.append(newBuyer);
        return newBuyer
    def FindBuyer(self, sNumber = "",sName = "",sEmail = ""):
        self.FillMapicsData();
        for buyer in self.aBuyer:
            if (buyer != None):
                if (sNumber != ""):
                    if (buyer.sNumber == sNumber):
                        return buyer;
                if (sName != ""):
                    if (sName == buyer.sName):
                        return buyer;
                if (sEmail != ""):
                    if (sEmail == buyer.sEmail):
                        return buyer;
    def FillMapicsData(self):
        if self.bFilled == False:
            server = 'tcp:ATLASSIAN' 
            database = 'master' 
            username = 'jiramapics' 
            password = 'FqAy&6KgqIIE6s$ya6uD' 
            cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
            cursor = cnxn.cursor()

            #Item Revision
            print("Begin Filling Buyer Data")
            sqlquery = "SELECT [BUYNO],[BUYNM],[EADR],[BUYPH],[FAXN] FROM [MAPICS].[POWER9].[AMFLIBF].[BUYERF];"
            cursor.execute(sqlquery) 
            row = cursor.fetchone()
            while row:
                self.AddBuyer(row[0],row[1],row[2],row[3],row[4])
                row = cursor.fetchone()
            cursor.close()
            cnxn.close()
            self.bFilled = True;

class cDefect:
    def __init__(self, sShouldBe = "",sActual = "",nQuantityRej = 0,nQuantityRec = 0):
        self.sShouldBe = sShouldBe
        self.sActual = sActual
        self.nQuantityRej = nQuantityRej
        self.nQuantityRec = nQuantityRec
    def GetMeasurement(self):
        return "S/B=" + self.sShouldBe + " Act=" + self.sActual

class cDefectList:
    def __init__(self):
        self.aDefect = []
    def AddDefect(self,sShouldBe = "",sActual = "",nQuantityRej = 0,nQuantityRec = 0):
        newDefect = cDefect(sShouldBe,sActual,nQuantityRej,nQuantityRec)
        self.aDefect.append(newDefect);
        return newDefect
    def GetMeasurementList(self):
        sMeasurementList = ""
        for defect in self.aDefect:
            sMeasurementList = sMeasurementList + defect.GetMeasurement()+ "; ";
        return sMeasurementList;

class cNCMR:
    def __init__(self, nNumber,sPartNumber,sPartDescription,sSupplier,sSupplierNumber,sCAIssued,sPurchaseOrder,sCreateDate,sCreator,nQuantityRej,nQuantityRec):
        self.nNumber = nNumber
        self.sFullNumber = "MRB-"+str(nNumber)
        self.sPartNumber = sPartNumber
        self.sPartDescription = sPartDescription
        self.sCAIssued = sCAIssued
        self.sPurchaseOrder = sPurchaseOrder
        self.sSupplier = sSupplier
        self.sSupplierNumber = sSupplierNumber
        self.sCreateDate = sCreateDate
        self.sCreator = sCreator
        self.nQuantityRej = nQuantityRej
        self.nQuantityRec = nQuantityRej
        self.DefectList = cDefectList()

class cNCMRList:
    def __init__(self):
        self.aNCMR = []
    def AddNCMR(self, nNumber,sPartNumber,sPartDescription,sSupplier,sSupplierNumber,sCAIssued,sPurchaseOrder,sCreateDate,sCreator,nQuantityRej,nQuantityRec):
        newncmr = cNCMR(nNumber,sPartNumber,sPartDescription,sSupplier,sSupplierNumber,sCAIssued,sPurchaseOrder,sCreateDate,sCreator,nQuantityRej,nQuantityRec)
        self.aNCMR.append(newncmr);
        return newncmr
    def FindNCMR(self, nNCMRNumber = -1, sFullNCMRNumber = ""):
        for ncmr in self.aNCMR:
            if (ncmr != None):
                if (nNCMRNumber != -1):
                    if (ncmr.nNumber == nNCMRNumber):
                        return ncmr;
                elif (sFullNCMRNumber != ""):
                    if (sFullNCMRNumber == ncmr.sFullNumber):
                        return ncmr;
        return None
    def RemoveNCMR(self, nNCMRNumber = -1,sFullNCMRNumber = ""):
        curncmr = self.FindNCMR(nNCMRNumber,sFullNCMRNumber)
        del curncmr
    def GetNCMRString(self):
        sNCMRString = ""
        for ncmr in self.aNCMR:
            sNCMRString = sNCMRString + ncmr.sFullNumber + "; "
        return sNCMRString
    def GetMeasurementList(self):
        sMeasurementList = ""
        for ncmr in self.aNCMR:
            sMeasurementList = sMeasurementList + ncmr.DefectList.GetMeasurementList() + "; "
        return sMeasurementList
    def GetDateOfLastNCMR(self):
        sCreateDate = "0"
        for ncmr in self.aNCMR:
            if (int(sCreateDate.replace("-","")) < int(ncmr.sCreateDate.replace("-",""))):
                sCreateDate = ncmr.sCreateDate
        return sCreateDate
    def GetTotalNCMR(self):
        return len(self.aNCMR)
    def GetTotalRejects(self):
        nNumber = 0
        for ncmr in self.aNCMR:
            nNumber = nNumber + ncmr.nQuantityRej
        return nNumber
    def GetTotalReceived(self):
        nNumber = 0
        for ncmr in self.aNCMR:
            nNumber = nNumber + ncmr.nQuantityRec
        return nNumber

class cPart:
    def __init__(self, sNumber,sDescription,sRevision = "",sEngNumber = "",sLocation = "",sHouse = "",sSupplierNumber = "",sSupplierName = "",bNewPart = False):
        self.sNumber = sNumber
        self.sDescription = sDescription
        self.sEngNumber = sEngNumber
        self.NCMRList = cNCMRList()
        self.sLocation = sLocation
        self.sRevision = sRevision
        self.bNewPart = bNewPart
        self.sHouse = sHouse
        self.sSupplierNumber = sSupplierNumber
        self.sSupplierName = sSupplierName
        self.ItemLocationList = cItemLocationList()
        self.fUnitCost = float(0.0)
        if (sRevision == None or sRevision == ""):
            if ("-" in sEngNumber):
                aEngNumber = sEngNumber.split("-")
                nLength = len(aEngNumber)
                sTemp = aEngNumber[nLength - 1]
                if len(sTemp) < 4:
                    self.sRevision = sTemp
            if ("_" in sEngNumber):
                aEngNumber = sEngNumber.split("_")
                nLength = len(aEngNumber)
                sTemp = aEngNumber[nLength - 1]
                if len(sTemp) < 4:
                    self.sRevision = sTemp
    def GetNCMRString(self):
        return self.NCMRList.GetNCMRString()
    def GetMeasurementList(self):
        return self.NCMRList.GetMeasurementList()
    def GetTotalNCMR(self):
        return self.NCMRList.GetTotalNCMR();
    def GetLocationString(self):
        return self.ItemLocationList.GetLocationString();
    

class cPartList:
    def __init__(self):
        self.aPart = []
        self.bFilled = False;
    def AddPart(self, sNumber,sDescription,sRevision = "",sEngNumber = "",sLocation = "",sHouse = "",sSupplierNumber = "",sSupplierName = "",bNewPart = False):
        newPart = cPart(sNumber,sDescription,sRevision,sEngNumber,sLocation,sHouse,sSupplierNumber,sSupplierName,bNewPart)
        self.aPart.append(newPart);
        return newPart
    def FindPart(self,sNumber = ""):
        for part in self.aPart:
            if part != None:
                if (part.sNumber == sNumber):
                    return part;
    def FindAllParts(self,sNumber = ""):
        parts = []
        for part in self.aPart:
            if (part.sNumber == sNumber):
                parts.append(part)
        if len(parts) > 0:
            return parts
        else:
            return None
    def GetSpecificParts(self, bNewPart = True):
        parts = []
        for part in self.aPart:
            if part != None:
                if (bNewPart and part.bNewPart):
                    parts.append(part)
                elif (bNewPart == False and part.bNewPart == False):
                    parts.append(part)
        if len(parts) > 0:
            return parts
        else:
            return None
    def GetPartList(self,bOnlyNCMRPart = False):
        sReturn = ""
        for part in self.aPart:
            if part != None:
                if bOnlyNCMRPart:
                    if part.GetTotalNCMR() > 0:
                        sReturn = sReturn + part.sNumber + "; "
                else:
                    sReturn = sReturn + part.sNumber + "; "
        return sReturn
    def FillList(self):
        if self.bFilled == False:
            server = 'tcp:ATLASSIAN' 
            database = 'master' 
            username = 'jiramapics' 
            password = 'FqAy&6KgqIIE6s$ya6uD' 
            cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
            cursor = cnxn.cursor()

            #Item Revision
            print("Begin Filling Item Revision Data")
            sqlquery = "SELECT [ITNBR],[ITDSC],[PACKC],[ENGNO],[WHSLC] FROM [MAPICS].[POWER9].[AMFLIBF].[ITMRVA];"
            cursor.execute(sqlquery) 
            row = cursor.fetchone()
            while row:
                #def AddPart(self, sNumber,sDescription,sRevision = "",sEngNumber = "",sLocation = "",sHouse = "",sSupplierNumber = "",sSupplierName = "",bNewPart = True):

                self.AddPart(row[0],row[1],row[2],row[3],row[4])
                row = cursor.fetchone()
            cursor.close()
            cnxn.close()
            self.bFilled = True;
class cItemLocation:
    def __init__(self,sNumber,sHouse,sLocation,sPurchaseOrder,sDate):
        self.sNumber = sNumber
        self.sHouse = sHouse
        self.sLocation = sLocation
        self.sPurchaseOrder = sPurchaseOrder
        self.sDescription = ""
        self.sDate = sDate

class cItemLocationList:
    def __init__(self):
        self.aItemLocation = []
        self.aQualityItemLocation = []
        self.bFilled = False;
    def AddItemLocation(self,sNumber,sHouse,sLocation,sPurchaseOrder,sDate = ""):
        newitemlocation = cItemLocation(sNumber,sHouse,sLocation,sPurchaseOrder, sDate)
        self.aItemLocation.append(newitemlocation);
        if (sLocation in ["QC01"]):
            self.aQualityItemLocation.append(newitemlocation);
        return newitemlocation
    def FindAllItemLocation(self,sPurchaseOrder = "",sNumber = ""):
        aReturn = []
        for itemlocation in self.aItemLocation:
            if (itemlocation != None):
                if (sPurchaseOrder != ""):
                    if (itemlocation.sPurchaseOrder == sPurchaseOrder):
                        aReturn.append(itemlocation)
                elif (sNumber != ""):
                    if (sNumber == itemlocation.sNumber):
                        aReturn.append(itemlocation)
        if len(aReturn) > 0:
            return aReturn;
        else:
            return None
        return None
    def FindItemLocation(self,sPurchaseOrder = "",sNumber = ""):
        for itemlocation in self.aItemLocation:
            if (itemlocation != None):
                if (sPurchaseOrder != ""):
                    if (itemlocation.sPurchaseOrder == sPurchaseOrder):
                        return itemlocation;
                elif (sNumber != ""):
                    if (sNumber == itemlocation.sNumber):
                        return itemlocation;
        return None
    def GetLocationString(self):
        sReturn = ""
        for loc in self.aItemLocation:
            sReturn += loc.sLocation + "; "
        return sReturn;
    def FillMapicsData(self,MapicsData = None):
        if self.bFilled == False:
            server = 'tcp:ATLASSIAN' 
            database = 'master' 
            username = 'jiramapics' 
            password = 'FqAy&6KgqIIE6s$ya6uD' 
            cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
            cursor = cnxn.cursor()

            #Item Location Data
            print("Begin Filling Item Location Data")
            sqlquery = "SELECT [ITNBR],[HOUSE],[LLOCN],[ORDNO],[FDATE] FROM [MAPICS].[POWER9].[AMFLIBF].[SLQNTY];"
            cursor.execute(sqlquery) 
            row = cursor.fetchone()
            while row:
                sPartNumber = row[0]
                sHouse = row[1]
                sLocation = row[2]
                sPurchaseOrder = row[3]
                sDate = str(row[4])
                sDateTrim = sDate[1:]
                s_datetime = datetime.datetime.strptime(sDateTrim, '%y%m%d')
                sDateReadable = s_datetime.strftime("%m/%d/%Y")
                location = self.AddItemLocation(sPartNumber,sHouse,sLocation,sPurchaseOrder,sDateReadable)
                if MapicsData != None:
                    POs = MapicsData.PurchaseOrderList.FindAllPO(sPurchaseOrder)
                    if POs != None:
                        for po in POs:
                            if po.sPartNumber == sPartNumber:
                                part = po.PartList.FindPart(sPartNumber)
                                if part != None:
                                    part.ItemLocationList.aItemLocation.append(location)
                row = cursor.fetchone()
            cursor.close()
            cnxn.close()
            self.bFilled = True;

class cPurchaseOrder:
    def __init__(self, sNumber,sPartNumber = "",sPartDescription = "",sSupplierNumber = "",sSupplierName = "",nMakeDate = 0,sEngNumber = "",sLocation = "",sRevision = "",bHistory = False,nQuantity = 0,dCost = 0.0,sBuyerNumber = ""):
        self.sNumber = sNumber
        self.sPartNumber = sPartNumber
        self.sPartDescription = sPartDescription
        self.sSupplierNumber = sSupplierNumber
        self.sSupplierName = sSupplierName
        self.bHistory = bHistory
        self.nMakeDate = nMakeDate
        self.sEngNumber = sEngNumber
        self.sLocation = sLocation
        self.sRevision = sRevision
        self.nQuantity = nQuantity
        self.sBuyerNumber = sBuyerNumber
        self.dCost = dCost
        self.PartList = cPartList()
        if (sRevision == None or sRevision == ""):
            if ("-" in sEngNumber):
                aEngNumber = sEngNumber.split("-")
                nLength = len(aEngNumber)
                sTemp = aEngNumber[nLength - 1]
                if len(sTemp) < 4:
                    self.sRevision = sTemp
            if ("_" in sEngNumber):
                aEngNumber = sEngNumber.split("_")
                nLength = len(aEngNumber)
                sTemp = aEngNumber[nLength - 1]
                if len(sTemp) < 4:
                    self.sRevision = sTemp
    def GetScrapCost(self,dQuantity):
        #"SELECT [EXTPR],[ACTQY] FROM [MAPICS].[POWER9].[AMFLIBF].[POITEM] WHERE ORDNO = '"+ sPo + "';"
        dTotalCost = self.dCost
        dTotalQuantity = self.nQuantity
        if (dTotalCost != None and dTotalQuantity != None):
            dCost = dTotalCost
            if (dQuantity != 0 and dTotalQuantity != 0):
                dCost = dCost*dQuantity/dTotalQuantity;
        return dCost

class cPurchaseOrderList:
    def __init__(self):
        self.aPurchaseOrder = []
        self.aPOOpen = []
        self.aPOHistory = []
        self.aPOHistoryRecent = []
        self.bFilled = False;
        self.bHistFilled = False;
        self.bOpenFilled = False;
    def AddPO(self, sNumber,sPartNumber = "",sPartDescription = "",sSupplierNumber = "",sSupplierName = "",nMakeDate = 0,sEngNumber = "",sLocation = "",sRevision = "",bHistory = False,nQuantity = 0,dCost = 0.0,sBuyerNumber = ""):
        newPO = cPurchaseOrder(sNumber,sPartNumber,sPartDescription,sSupplierNumber,sSupplierName,nMakeDate,sEngNumber,sLocation,sRevision,bHistory,nQuantity,dCost,sBuyerNumber)
        self.aPurchaseOrder.append(newPO);
        if bHistory:
            self.aPOHistory.append(newPO);
            if (nMakeDate > 1140000.0 or nMakeDate == 0.0):
                self.aPOHistoryRecent.append(newPO);
        else:
            self.aPOOpen.append(newPO);
        return newPO
    def FindPO(self, sNumber,bOpen = True, bHistory = True, bRecentHistory = True,MapicsData = None):
        if bOpen:
            self.FillPOOpen(MapicsData);
            for po in self.aPOOpen:
                if (po != None):
                    if (po.sNumber == sNumber):
                        return po;
        if bHistory:
            self.FillPOHistory(MapicsData);
            if bRecentHistory:
                for po in self.aPOHistoryRecent:
                    if (po != None):
                        if (po.sNumber == sNumber):
                            return po;
            else:
                for po in self.aPOHistory:
                    if (po != None):
                        if (po.sNumber == sNumber):
                            return po;
        return None
    def FindAllPO(self,sNumber,bOpen = True, bHistory = True, bRecentHistory = True,MapicsData = None):
        pos = []
        if bOpen:
            self.FillPOOpen(MapicsData);
            for po in self.aPOOpen:
                if (po != None):
                    if (po.sNumber == sNumber):
                        pos.append(po);     
        if bHistory:
            if len(pos) == 0:
                self.FillPOHistory(MapicsData);
            if bRecentHistory:
                for po in self.aPOHistoryRecent:
                    if (po != None):
                        if (po.sNumber == sNumber):
                            pos.append(po);
            else:
                for po in self.aPOHistory:
                    if (po != None):
                        if (po.sNumber == sNumber):
                            pos.append(po);
        if len(pos) > 0:
            return pos;
        else:
            return None
    def GetSpecificPOs(self,bHistory = False, bRecentHistory = True,MapicsData = None):
        if bHistory:
            self.FillPOHistory(MapicsData);
            if bRecentHistory:
                return self.aPOHistoryRecent;
            else:
                return self.aPOHistory;
        else:
            self.FillPOOpen(MapicsData);
            return self.aPOOpen
    def RemovePO(self, sNumber):
        curPo = self.FindPO(sNumber);
        del curPo
    def FillPOOpen(self, MapicsData = None):
        if self.bOpenFilled == False:
            server = 'tcp:ATLASSIAN' 
            database = 'master' 
            username = 'jiramapics' 
            password = 'FqAy&6KgqIIE6s$ya6uD' 
            cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
            cursor = cnxn.cursor()

            #Supplier Data
            #Fill PO DATA
            print("Begin Filling Open PO Data")
            #def (sNumber,sPartNumber,sPartDescription,sSupplierNumber,sSupplierName,nMakeDate = 0,sEngNumber = "",sLocation = "",sRevision = "",bHistory = False):
    
            sqlquery = "SELECT [ITNBR],[ITDSC],[PACKC],[ENGNO],[BUYNO],[VNDNR],[ORDNO],[MDATE],[WHSLC],[EXTPR],[QTYOR] FROM [MAPICS].[POWER9].[AMFLIBF].[POITEM];"
            cursor.execute(sqlquery) 
            row = cursor.fetchone()   
            while row:
                sItemNo = row[0]
                sItemDesc = row[1]
                sRev = row[2]
                sEngNo = row[3]
                sBuyerNo = row[4]
                sVendorNo = row[5]
                sOrderNo = row[6]
                sMDate = float(row[7])
                sLocation= row[8]
                dCost = float(row[9])
                nQuantity = int(row[10])
            

                po = self.AddPO(sOrderNo,sItemNo,sItemDesc,sVendorNo,"",sMDate,sEngNo,sLocation,sRev,False,nQuantity,dCost,sBuyerNo)
                if MapicsData != None:
                    Supplier = MapicsData.SupplierList.FindSupplier(sVendorNo)
                    if Supplier == None:
                        Supplier = MapicsData.SupplierList.AddSupplier(sVendorNo,"","",False,"","")
                    #def AddPart(self, sNumber,sDescription,sRevision = "",sEngNumber = "",sLocation = "",sHouse = "",sSupplierNumber = "",sSupplierName = "",bNewPart = True):
                    newPart = Supplier.PartList.FindPart(sItemNo);
                    if newPart == None:
                        newPart = Supplier.PartList.AddPart(sItemNo,sItemDesc,sRev,sEngNo,sLocation,"",sVendorNo,Supplier.sName,False)
                        MapicsData.PartList.aPart.append(newPart);
                    po.PartList.aPart.append(newPart)
                    po.sSupplierName = Supplier.sName
                    Supplier.PurchaseOrderList.aPurchaseOrder.append(po);
                row = cursor.fetchone()
            cursor.close()
            cnxn.close()
            self.bOpenFilled = True;
    def FillManualPOOpen(self, MapicsData = None):
        if self.bOpenFilled == False:
            server = 'tcp:ATLASSIAN' 
            database = 'master' 
            username = 'jiramapics' 
            password = 'FqAy&6KgqIIE6s$ya6uD' 
            cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
            cursor = cnxn.cursor()

            #Supplier Data
            #Fill PO DATA
            print("Begin Filling Open PO Data")
            #def (sNumber,sPartNumber,sPartDescription,sSupplierNumber,sSupplierName,nMakeDate = 0,sEngNumber = "",sLocation = "",sRevision = "",bHistory = False):
    
            sqlquery = "SELECT [Item],[Item description],[PACKC],[ENGNO],[Buyer],[VNDNR],[Order],[Release],[WHSLC],[Extended amount],[Order Quantity Requested] FROM [DataStorage].[dbo].[POITEM$];"
            cursor.execute(sqlquery) 
            #Name = Vendor Name
            row = cursor.fetchone()   
            while row:
                sItemNo = row[0]
                sItemDesc = row[1]
                sRev = ""
                sEngNo = ""
                sBuyerNo = row[4]
                sVendorNo = row[5]
                sOrderNo = row[6]
                sMDate = float(row[7])
                sLocation= row[8]
                dCost = float(row[9])
                nQuantity = int(row[10])
            

                po = self.AddPO(sOrderNo,sItemNo,sItemDesc,sVendorNo,"",sMDate,sEngNo,sLocation,sRev,False,nQuantity,dCost,sBuyerNo)
                if MapicsData != None:
                    Supplier = MapicsData.SupplierList.FindSupplier(sVendorNo)
                    if Supplier == None:
                        Supplier = MapicsData.SupplierList.AddSupplier(sVendorNo,"","",False,"","")
                    #def AddPart(self, sNumber,sDescription,sRevision = "",sEngNumber = "",sLocation = "",sHouse = "",sSupplierNumber = "",sSupplierName = "",bNewPart = True):
                    newPart = Supplier.PartList.FindPart(sItemNo);
                    if newPart == None:
                        newPart = Supplier.PartList.AddPart(sItemNo,sItemDesc,sRev,sEngNo,sLocation,"",sVendorNo,Supplier.sName,False)
                        MapicsData.PartList.aPart.append(newPart);
                    po.PartList.aPart.append(newPart)
                    po.sSupplierName = Supplier.sName
                    Supplier.PurchaseOrderList.aPurchaseOrder.append(po);
                row = cursor.fetchone()
            cursor.close()
            cnxn.close()
            self.bOpenFilled = True;
    def FillPOHistory(self, MapicsData = None):
        if self.bHistFilled == False:
            server = 'tcp:ATLASSIAN' 
            database = 'master' 
            username = 'jiramapics' 
            password = 'FqAy&6KgqIIE6s$ya6uD' 
            cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
            cursor = cnxn.cursor()

            #Supplier Data
            #Fill PO DATA
            print("Begin Filling History PO Data")
            #sqlquery = "SELECT [ITNBR],[ITDSC],[PACKC],[ENGNO],[BUYNO],[VNDNR],[ORDNO],[MDATE] FROM [MAPICS].[POWER9].[AMFLIBF].[POHISTI] WHERE MDATE > '1150000';"
            sqlquery = "SELECT [ITNBR],[ITDSC],[PACKC],[ENGNO],[BUYNO],[VNDNR],[ORDNO],[MDATE],[WHSLC],[EXTPR],[QTYOR] FROM [MAPICS].[POWER9].[AMFLIBF].[POHISTI];"
            cursor.execute(sqlquery) 
            row = cursor.fetchone()   
            while row:
                sItemNo = row[0]
                sItemDesc = row[1]
                sRev = row[2]
                sEngNo = row[3]
                sBuyerNo = row[4]
                sVendorNo = row[5]
                sOrderNo = row[6]
                sMDate = float(row[7])
                sLocation= row[8]
                dCost = float(row[9])
                nQuantity = int(row[10])
                po = self.AddPO(sOrderNo,sItemNo,sItemDesc,sVendorNo,"",sMDate,sEngNo,sLocation,sRev,True,nQuantity,dCost,sBuyerNo)
                if MapicsData != None:
                    Supplier = MapicsData.SupplierList.FindSupplier(sVendorNo)
                    if Supplier == None:
                        Supplier = MapicsData.SupplierList.AddSupplier(sVendorNo,"","",False,"","")
                    #def AddPart(self, sNumber,sDescription,sRevision = "",sEngNumber = "",sLocation = "",sHouse = "",sSupplierNumber = "",sSupplierName = "",bNewPart = True):
                    newPart = Supplier.PartList.FindPart(sItemNo);

                    if newPart == None:
                        newPart = Supplier.PartList.AddPart(sItemNo,sItemDesc,sRev,sEngNo,sLocation,"",sVendorNo,Supplier.sName,False)
                        MapicsData.PartList.aPart.append(newPart);
                    po.PartList.aPart.append(newPart)
                    po.sSupplierName = Supplier.sName
                    Supplier.PurchaseOrderList.aPurchaseOrder.append(po);
                #po = self.AddPO(sOrderNo,sItemNo,sItemDesc,sVendorNo,"",sMDate,sEngNo,sLocation,sRev,True)
                row = cursor.fetchone()
            cursor.close()
            cnxn.close()
            self.bHistFilled = True;
    def FillMapicsData(self,MapicsData = None):
       if self.bFilled == False:
            self.FillPOOpen(MapicsData);
            self.FillPOHistory(MapicsData);
            self.bFilled = True;
            self.bHistFilled = True;
            self.bOpenFilled = True;
   
class cSupplier:
    def __init__(self, sNumber, sName,sAddress = "",bHasNCMR = False,sPhone = "",sFax = "",sEmail = "",sEmailList = ""):
        self.sName = sName
        self.sNumber = sNumber
        self.sAddress = sAddress
        self.sPhone = sPhone
        self.sFax = sFax
        self.PartList = cPartList();
        self.bHasNCMR = bHasNCMR;
        self.NCMRList = cNCMRList();
        self.PurchaseOrderList = cPurchaseOrderList();
        self.sEmail = sEmail
        self.sEmailList = sEmailList

class cSupplierList:
    def __init__(self):
        self.aSupplier = []
        self.bFilled = False;
    def AddSupplier(self, sNumber,sName,sAddress = "",bHasNCMR = False,sPhone = "",sFax = "",sEmail = "",sEmailList = ""):
        newSupplier = cSupplier(sNumber,sName,sAddress,bHasNCMR,sPhone,sFax,sEmail,sEmailList)
        self.aSupplier.append(newSupplier);
        return newSupplier
    def FindSupplier(self, sNumber = "", sName = "",bFill = True):
        if bFill:
            self.FillMapicsData()
        for supplier in self.aSupplier:
            if (supplier != None):
                if (sNumber != ""):
                    if (supplier.sNumber == sNumber):
                        return supplier;
                elif (sName != ""):
                    if (supplier.sName == sName):
                        return supplier;
        return None
    def RemoveSupplier(self, sNumber = "", sName = ""):
        curSuppplier = self.FindSupplier(sNumber,sName);
        del curSupplier
    def FillMapicsData(self):
        if self.bFilled == False:
            server = 'tcp:ATLASSIAN' 
            database = 'master' 
            username = 'jiramapics' 
            password = 'FqAy&6KgqIIE6s$ya6uD' 
            cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
            cursor = cnxn.cursor()
            #Supplier Data
            print("Begin Filling Supplier Data")
            sqlquery = "SELECT [VNAME],[VADD1],[VADD2],[VCITY],[VSTAC],[VZIPC],[VETEL],[FAXTN],[VNDNR],[EADR] FROM [MAPICS].[POWER9].[AMFLIBF].[VENNAM];"
            cursor.execute(sqlquery) 
            row = cursor.fetchone()
            while row:
                sVendorName = row[0]    
                if (row[2] != ''):
                    sVendorAddress = row[0] + "\r\n" + row[1] + "\r\n" + row[2] + "\r\n" + row[3] + ", " + row[4] + " " + row[5]
                else:
                    sVendorAddress = row[0] + "\r\n" + row[1] + "\r\n" + row[3] + ", " + row[4] + " " + row[5]
                sVendorPhone = row[6]
                sVendorFax = row[7]
                sVendorNo = row[8]
                sVendorEmail = row[9]
                if sVendorEmail != "" and sVendorEmail != None:
                    sVendorEmail = sVendorEmail.replace(" ","")
                    sVendorEmail+=","
                #supplier = self.findsupplier(sVendoNo)
                #if supplier == None:
                supplier = self.AddSupplier(sVendorNo,sVendorName,sVendorAddress,False,sVendorPhone,sVendorFax,sVendorEmail,sVendorEmail)
                row = cursor.fetchone()
            #Fill Contact Data
            sqlquery = "SELECT [Supplier],[Contact Name],[Email],[Phone],[Vendor Number] FROM [DataStorage].[dbo].[Contact list];"
            cursor.execute(sqlquery) 
            row = cursor.fetchone()
            while row:
                sVendorNo = row[4]
                sEmail = row[2]
                supplier = self.FindSupplier(sVendorNo,"",False)
                if supplier != None and sEmail != None:
                    sEmail = sEmail.replace(" ","")
                    if not sEmail.lower() in supplier.sEmailList.lower():
                        supplier.sEmailList+=sEmail+","
                        supplier.sEmailList.replace("ach@accuratedial.com,","")
                row = cursor.fetchone()
            cursor.close()
            cnxn.close()
            self.bFilled = True;
