#####################################################################
# Creates Excel file from NCT query
# TODO Expand Pivot table with all necessary fields.
# TODO clear gen_py data if it fails
#####################################################################
import ReportGenerator
from os import path
import win32com.client
import time
from tkinter import Tk
from tkinter import filedialog

#####################################################################
# Functions
#####################################################################


def populate_excel(main_list):
    # Given a list this function will populate an excel file
    global CountProgressBar
    for i, NCT_Row in enumerate(main_list):
        for j, NCT_Item in enumerate(NCT_Row):
            RawDataSheet.Cells(i + 1, j + 1).Value = NCT_Item
        CountProgressBar += 1
        print("{0:.1f}".format(100 * (CountProgressBar / len(NCT_list))))


def pivot_fields(field, pivot_table, row=False, column=False, data=False):
    # Given a field, and the location of field, this function will populate the pivot table with it
    if row:
        pivot_table.PivotFields(field).Orientation = win32c.xlRowField
    elif column:
        pivot_table.PivotFields(field).Orientation = win32c.xlColumnField
    elif data:
        data = pivot_table.AddDataField(pivot_table.PivotFields(field))
        data.Caption = "Count of " + field
        data.Function = win32c.xlCount
        return None
    pivot_table.PivotFields(field).Position = 1


def add_worksheet(workbook, sheet_name):
    # Given a workbook and a sheet name function will create the sheet
    global SheetNumber
    SheetNumber += 1
    workbook.Sheets.Add()
    workbook.Sheets("Sheet"+str(SheetNumber)).Name = sheet_name
    return workbook.Worksheets(sheet_name)


def create_pivot_table(sheet, range_origin, range_end):
    # Given sheet and range for which to construct pivot table one is created
    target_range = sheet.Cells(1, 1)
    source_range = RawDataSheet.Range(range_origin, range_end)
    pivot_cache = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=source_range,
                                          Version=win32c.xlPivotTableVersion14)
    pivot_table_func = pivot_cache.CreatePivotTable(TableDestination=target_range, TableName=sheet.Name,
                                                    DefaultVersion=win32c.xlPivotTableVersion14)
    pivot_chart = sheet.Shapes.AddChart2(201)
    pivot_chart.Width = 720
    pivot_chart.Height = 400

    return pivot_table_func


def create_slicer(slicer_pivot_table, slicer_field, slicer_sheet_target):
    slicer_cache = wb.SlicerCaches.Add2(slicer_pivot_table, slicer_field)
    slicer_obj = slicer_cache.Slicers.Add(slicer_sheet_target)
    return slicer_cache


def save_file():
    # Saves Excel File
    filepath = filedialog.asksaveasfilename(title="Save as...", filetypes=(("Microsoft Excel Worksheet", ".xlsx"),
                                                                           ("All Files", "*.*")))
    # Try - Except in case Excel is open TODO Improve
    wb.SaveAs("C:\\PythonProjects\\")
    return None

#####################################################################
# Executable Code
#####################################################################
# Bring GUI to foreground


root = Tk()
root.withdraw()
root.lift()
root.attributes('-topmost', True)
root.after_idle(Tk.attributes, '-topmost', False)

# Define some variables
t0 = time.time()
CountProgressBar = 0
SheetNumber = 0
NCT_list = ReportGenerator.report_creation()

# Excel File Creation

Excel = win32com.client.gencache.EnsureDispatch('Excel.Application')  # Import Excel application
win32c = win32com.client.constants  # Import specific Excel constants
wb = Excel.Workbooks.Add()  # Create the workbook
RawDataSheet = add_worksheet(wb, "RawData")  # Create raw data sheet

# Make Excel Visible for troubleshooting
Excel.Visible = 0

# Populate Excel File
populate_excel(NCT_list)

################################################
print("Populated Excel Workbook")
print(NCT_list[1])
################################################

# Create the range of where data is in sheet 1
RawData_RangeOrigin = RawDataSheet.Cells(1, 1)
RawData_RangeEnd = RawDataSheet.Cells(1 + len(NCT_list) - 1, 1 + len(NCT_list[0]) - 1)
RawDataSheet.Columns.AutoFit()

# Creates pivot tables. TODO Create rest of tables
PivotTableSheetOrigin = add_worksheet(wb, "Root Cause Origin")
PivotOrigin = create_pivot_table(PivotTableSheetOrigin, RawData_RangeOrigin, RawData_RangeEnd)

# Populate fields on Pivot table for Root Cause Origin TODO Create rest of fields
pivot_fields("ID", PivotOrigin, data=True)
pivot_fields("Root Cause Origin", PivotOrigin, row=True)
pivot_fields("Veoneer Owner", PivotOrigin, column=True)

# Create Slicers
ProductFamilySlicer = create_slicer(PivotOrigin, "Customer / OEM", PivotTableSheetOrigin)
CustomerSlicer = create_slicer(PivotOrigin, "Dealer", PivotTableSheetOrigin)
VeoneerFacilitySlicer = create_slicer(PivotOrigin, "Veoneer Facility", PivotTableSheetOrigin)


# Save Excel and close it
t1 = time.time()
print(t1-t0)
wb.SaveAs("C:\\PythonProjects\\NCT_April30v2.xlsx")
#save_file()
#Excel.Application.Quit()
