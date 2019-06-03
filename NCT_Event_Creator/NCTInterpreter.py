#######################################################################
# Excel Interpreter
# Reads data from the Excel sheet and fills in the required data into the DataContainer
#######################################################################
import tkinter
from tkinter import messagebox
from GSARExcelSheet import *
from NCTData import *

Chargebacks8D = ["[825772]: CHARGEBACK - EXT-DEVICE OPEN/SHORT",
                 "[825775]: CHARGEBACK - COMMUNICATION FAULT",
                 "[825777]: CHARGEBACK - CONFIGURATION ERROR",
                 "[825778]: CHARGEBACK - EDR_FULL_AND_LOADED",
                 "[825779]: CHARGEBACK - ENS_OPEN_SHORT_TO_GND",
                 "[825780]: CHARGEBACK - IGNITION LOW",
                 "[825781]: CHARGEBACK - KNOWN ISSUE",
                 "[825782]: CHARGEBACK - NTF",
                 "[825783]: CHARGEBACK - SERVICE MODULE",
                 "[825784]: CHARGEBACK - VIN NOT PROGRAMMED",
                 "[825785]: CHARGEBACK - WATER CONTAMINATION",
                 "[825786]: CHARGEBACK - MODULE DAMAGED",
                 "[825787]: CHARGEBACK - WRONG PART RETURNED",
                 "[833660]: CHARGEBACK - OCS FAULT",
                 "[837056]: CHARGEBACK - ABS",
                 "[843433]: CHARGEBACK - SYSTEMS",
                 "[852906]: CHARGEBACK - OUT OF WARRANTY",
                 "[864387]: CHARGEBACK - Bosch Gyro Channel Fault"
                 ]

#Returns None if couldnt cast to int, otherwise returns the casted int.
def __safeIntCast(value):
    try:
        return int(value)
    except ValueError:
        return None

#TODO add "Progress" output. E.X. print("Started Processing") and print("Generating NCT Number...")
#Given a cell from the excel sheet this function will interpret and store the appropriate data.
#Returns True if data was found, False if the cell was empty
def __interpretData(ws:GSARExcelWorksheet, dataContainer:NCTDataContainer, row:int, type:GSARExcelDataLocation):
    value = ws.GetValue(row, type)
    if value is None:
        return False
    if type == GSARExcelDataLocation.VIN:
        dataContainer.setData(NCTDataType.VIN, value)
    elif type == GSARExcelDataLocation.CSN:
        dataContainer.setData(NCTDataType.SerialNumber, value)
    elif type == GSARExcelDataLocation.VeoneerProductionDate:
        dataContainer.setData(NCTDataType.ProductionDate, value)
    elif type == GSARExcelDataLocation.CustomerPN:
        dataContainer.setData(NCTDataType.CustomerPN, value)
        dataContainer.setData(NCTDataType.VeoneerPN, value)
    elif type == GSARExcelDataLocation.VehicleLine:
        dataContainer.setData(NCTDataType.VehicleLine, value)
    elif type == GSARExcelDataLocation.DTCs:
        dataContainer.setData(NCTDataType.ActiveDTC, value)
    elif type == GSARExcelDataLocation.HistoricDTCs:
        dataContainer.setData(NCTDataType.HistoricDTC, value)#TODO turns out \n and \r characters are evil too!
    #elif type == GSARExcelDataLocation.Chargeback:
        #This field is ignored.
    elif type == GSARExcelDataLocation.PreAnalysisResults:
        dataContainer.setData(NCTDataType.PreAnalysisResults, value)
    elif type == GSARExcelDataLocation.DateOfArrival:
        dataContainer.setData(NCTDataType.DateOfArrival, value)
        #TODO set "Analysis Due Date" 30 days after arrival
    elif type == GSARExcelDataLocation.ModelYear:
        dataContainer.setData(NCTDataType.ModelYear, value)
    elif type == GSARExcelDataLocation.CustomerPickList:
        dataContainer.setData(NCTDataType.Customer, value.capitalize())
    elif type == GSARExcelDataLocation.VehicleProductionDate:
        dataContainer.setData(NCTDataType.VehicleProductionDate, value)
    elif type == GSARExcelDataLocation.WSD:
        dataContainer.setData(NCTDataType.WSD, value)
        dataContainer.setData(NCTDataType.MonthOfWarranty, value)
    elif type == GSARExcelDataLocation.RejectDate:
        dataContainer.setData(NCTDataType.CustomerRejectDate, value)
        dataContainer.setData(NCTDataType.DateOfFailure, value)
        dataContainer.setData(NCTDataType.First8DSentDate, value)
    #elif type == GSARExcelDataLocation.TIS_WSD:
        #I don't know what this is.
    #elif type == GSARExcelDataLocation.MileageMiles:
        #Not used
    elif type == GSARExcelDataLocation.OdometerKM:
        safeCast = __safeIntCast(value)
        if not (safeCast is None):
            dataContainer.setData(NCTDataType.OdometerKM, safeCast)
    elif type == GSARExcelDataLocation.DealerState:
        dataContainer.setData(NCTDataType.DealerState, value)
    elif type == GSARExcelDataLocation.Dealer:
        dataContainer.setData(NCTDataType.DealerName, value)
    elif type == GSARExcelDataLocation.ClaimInfo:
        dataContainer.setData(NCTDataType.CustomerComments, value)
        #TODO customer complaint category
        #if ("lamp" in value.lower()) or ("light" in value.lower()):
        #    dataContainer.setData(NCTDataType.CustomerCo)
    elif type == GSARExcelDataLocation.DealerCountry:
        if value in "CAN":
            value = "Canada"
        dataContainer.setData(NCTDataType.DealerCountry, value)
    elif type == GSARExcelDataLocation.CustomerFailureCode:
        if isinstance(value, int):
            value = str(value)
        dataContainer.setData(NCTDataType.CustomerFailureCode, value)
    elif type == GSARExcelDataLocation.ComplaintSummary: #TODO dealer comments?
        dataContainer.setData(NCTDataType.ComplaintSummary, value)
    elif type == GSARExcelDataLocation.ProductFamily:
        dataContainer.setData(NCTDataType.ProductFamily, value)
    elif type == GSARExcelDataLocation.VeoneerFacility:
        value = value.lower()
        if value in "cmm - veoneer canada markham":
            value = "CMM - Veoneer Canada Markham"
        elif value in "frm - veoneer france rouen":
            value = "FRM - Veoneer France Rouen"
        elif value in "cfm - veoneer china fengxian":
            value = "CFM - Veoneer China Fengxian"
        #TODO what if it is none of these???
        dataContainer.setData(NCTDataType.VeoneerFacility, value)
    elif type == GSARExcelDataLocation.CustomerPlant:
        dataContainer.setData(NCTDataType.CustomerPlant, value)
    elif type == GSARExcelDataLocation.Summary:
        dataContainer.setData(NCTDataType.ProblemDescription, value)
    elif type == GSARExcelDataLocation.IssueCategory:
        dataContainer.setData(NCTDataType.IssueCategory, value.capitalize())
    elif type == GSARExcelDataLocation.Project:
        dataContainer.setData(NCTDataType.PermissionGroup, "/Analysis/RCS Analysis")
    elif type == GSARExcelDataLocation.VeoneerOwner:
        dataContainer.setData(NCTDataType.VeoneerOwner, value[value.find("(") + 1:value.find(")")])
    elif type == GSARExcelDataLocation.AssignedUser:
        dataContainer.setData(NCTDataType.AssignedUser, value[value.find("(") + 1:value.find(")")])
    elif type == GSARExcelDataLocation.ProductArea:
        dataContainer.setData(NCTDataType.Fault8DCategory, value)
        dataContainer.setData(NCTDataType.ProductLine, value)
    elif type == GSARExcelDataLocation.Category8D:
        chargebackValue = ws.GetValue(row, GSARExcelDataLocation.Chargeback)
        if not (chargebackValue is None):
            if chargebackValue.lower() in "y":
                for FTF in Chargebacks8D:
                    if value in FTF:
                        value = FTF
                dataContainer.setData(NCTDataType.Fault8D, value)
                dataContainer.setData(NCTDataType.CustomerStatus, "Closed")
                dataContainer.setData(NCTDataType.CurrentLocation, "CHARGEBACK")
            else:
                dataContainer.setData(NCTDataType.Fault8D, "[840805]: ECU - UNDER ANALYSIS")
                dataContainer.setData(NCTDataType.CurrentLocation, "UST")
    return True

#Returns an array of NCTDataContainer with the interpreted data.
def LoadAndInterpretGSARData():
    ws = GSARExcelWorksheet() #Load the GSAR data from a file.
    InterpretedData = []

    for cellRow in range(2, ws.MaxRow + 1):
        print("Parsing Row#: " + str(cellRow))
        #Each row represents a different NCT that needs to be created.
        NCTData = NCTDataContainer(cellRow) #This is what will be storing the data that gets parsed.
        DataInRow = False
        for type in GSARExcelDataLocation:
            #Attempt to read every data type from the excel sheet.
            if __interpretData(ws, NCTData, cellRow, type) == True:
                DataInRow = True
            #print(repr(type))
        if not DataInRow:
            break
        else:
            InterpretedData.append(NCTData)
    return InterpretedData




