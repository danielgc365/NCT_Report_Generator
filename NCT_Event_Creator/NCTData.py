#######################################################################
# NCT Data
# Stores the data to be entered into NCT
#
# Enum: NCTDataType - All the different fields available in NCT. The order that they are listed here is
#                   the order that the data will be entered into NCT (e.x. IssueCategory must be set for some fields to become available).
#                   To add a new field from NCT, simply add the field name to the DataType enum below.
#                   
#######################################################################

from enum import Enum
#from subprocess import check_output
import subprocess
import os
import sys
import tkinter
from tkinter import messagebox
import datetime

#The order that they are listed here affect the order they are entered into NCT
class NCTDataType(Enum):
    ProblemDescription = "--field=\'Summary="
    IssueCategory = "--field=\'FA_Issue Category="
    PermissionGroup = "--field=\'Project="
    VeoneerOwner = "--field=\'CQE="
    AssignedUser = "--field=\'Assigned User="
    CustomerRejectDate = "--field=\'FA_Customer Reject Date="
    DateOfFailure = "--field=\'Error discovered date="
    First8DSentDate = "--field=\'FA_First 8D Sent Date="
    Fault8DCategory = "--field=\'AI_Product Area="
    Fault8D = "--field=\'AI_Fault Type Others="
    CustomerComments = "--field=\'AI_Claim info="
    CustomerFailureCode = "--field=\'FA_Customer Failure Code="
    Customer = "--field=\'AI_Customer (pick list)="
    ComplaintSummary = "--field=\'AI_Complaint Summary="
    DealerState = "--field=\'FA_Dealer State [US Only]="
    DealerName = "--field=\'FA_Dealer="
    DealerCountry = "--field=\'FA_Dealer Country="
    VIN = "--field=\'AI_Chassi number="
    SerialNumber = "--field=\'AI_CSN number="
    ProductionDate = "--field=\'AI_Production date="
    CustomerPN = "--field=\'AI_Customer Part No="
    VeoneerPN = "--field=\'AI_Part No="
    VehicleLine = "--field=\'FA_Vehicle Line (Code)="
    ActiveDTC = "--field=\'AI_DTC="
    HistoricDTC = "--field=\'FA_Historic DTCs="
    PreAnalysisResults = "--field=\'Pre-Analysis Results="
    DateOfArrival = "--field=\'AI_Date of Arrival="
    ModelYear = "--field=\'FA_MY (Model Year)="
    VehicleProductionDate = "--field=\'Vehicle production date="
    WSD = "--field=\'Vehicle registration date="
    MonthOfWarranty = "--field=\'FA_Month Of Warranty="
    OdometerKM = "--field=\'AI_Km="
    VeoneerFacility = "--field=\'AI_Country="
    CustomerPlant = "--field=\'AI_Error discovered (place)="
    ProductLine = "--field=\'FA_Product Line="
    ProductCategory = "--field=\'FA_Product Category="
    ProductFamily = "--field=\'FA_Product Family="
    CurrentLocation = "--field=\'AI_Current Location="
    CustomerStatus = "--field=\'AI_Customer Status="

RequiredFields = [  
    NCTDataType.VIN, 
    NCTDataType.SerialNumber, 
    NCTDataType.ProblemDescription, 
    NCTDataType.IssueCategory, 
    NCTDataType.PermissionGroup, 
    NCTDataType.VeoneerOwner, 
    NCTDataType.AssignedUser, 
    NCTDataType.CustomerRejectDate, 
    NCTDataType.Fault8DCategory, 
    NCTDataType.Fault8D, 
    NCTDataType.Customer,  
    NCTDataType.CustomerPlant, 
    NCTDataType.VeoneerPN, 
    NCTDataType.CustomerPN, 
    NCTDataType.VeoneerFacility, 
    NCTDataType.ProductLine, 
    NCTDataType.ProductCategory, 
    NCTDataType.ProductFamily, 
    NCTDataType.First8DSentDate,
    NCTDataType.CustomerStatus,
    NCTDataType.CurrentLocation
]

class NCTDataContainer:
    def __init__(self, excelRow:int):
        self.data = {}
        self.ExcelSheetRow = excelRow #Used for debugging
    
    def GetExcelSheetRow(self):
        return self.ExcelSheetRow

    def setData(self, dataType:NCTDataType, value):
        self.data[dataType] = value

    def getData(self, dataType:NCTDataType):
        if dataType in self.data:
            return self.data[dataType]
        else:
            return None

    #Returns true if all the required fields have data, otherwise returns the missing fields.
    def CheckRequiredFields(self):
        MissingFields = []
        for requirement in RequiredFields:
            if self.getData(requirement) == None:
                MissingFields.append(requirement)
        if not MissingFields: #If the list is empty
            return True    
        else:
            return MissingFields

    def __generateNCTString(self):
        creation_string = "im createissue --type=\'Non Conforming Tracking - NCT\'"#TODO There are specials characters like "&" that need to be removed.
        for field in NCTDataType:
            #Store all fields from the DataContainer
            value = self.getData(field)
            if not (value is None):
                if isinstance(value, datetime.datetime):
                    value = value.strftime("%b %d, %Y")
                elif not isinstance(value, str):
                    value = str(value)
                creation_string = creation_string + " " + field.value + value + "\'"
        return creation_string

    def __getNextDebugFileName(self):
        for i in range(1, 100000):
            name = "Debug" + str(i).zfill(5)
            if not os.path.isfile(name + ".txt"):
                return name

    def __saveDebugInfo(self, ex:Exception):
        fileName = self.__getNextDebugFileName()
        try:
            debugFile = open(fileName + ".txt","a")
            debugFile.write(str(type(ex)) + "\n")
            debugFile.write(str(ex))
            debugFile.close()
            return fileName
        except:
            print("Error writing debug file: " + fileName)
            return None

    #Returns the NCT number if requirements were met and NCT was created.
    #Returns missing requirements if NCT was not created.
    def CreateNCT(self):
        #TODO We need to assume the product category
        self.setData(NCTDataType.ProductCategory, "ECU")
        requirementsMet = self.CheckRequiredFields()
        if requirementsMet == True:
            #Generate the NCT
            creation_string = self.__generateNCTString()
            creation_string = creation_string.replace('\r', '').replace('\t', '').replace('\n', '').replace('&', '')
            try:
                #byteList = check_output(creation_string, shell=True)#[nct for nct in check_output(creation_string).split() if nct.isdigit()]
                result = subprocess.run(creation_string, capture_output=True) #stdout=subprocess.PIPE
                print(result.stderr) #More for debugging, passing on the result
                #print(result.stderr.flush())#TODO null checks
                #print(result.stderr)
                lines = result.stderr.decode("utf-8")[:-2].replace('\r', '').split('\n')
                #print(lines)
                #resultString = result.stderr.decode("utf-8")[:-2] #Convert bytes to string and trim newline and carriage return chars
                for line in lines:
                    if line.startswith("Created Non Conforming Tracking - NCT "):
                        NCT_number = int(line[len("Created Non Conforming Tracking - NCT "):])
                        print("Created " + str(NCT_number))
                        return NCT_number
                return [] #Failed NCT creation
            except Exception as inst:
                fileName = self.__saveDebugInfo(inst)
                if not (fileName is None):
                    messagebox.showinfo("Debug Saved", "Saved debug info to \"" + fileName +".txt\"")
                else:
                    messagebox.showinfo("Debug Error", "Failed to save debug info.")
                return []
        else:
            return requirementsMet
