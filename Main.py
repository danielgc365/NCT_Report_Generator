import ReportGenerator
import ExcelFileCreator
import time

t0 = time.time()
NCT_Master = ReportGenerator.report_creation()
t1 = time.time()

print(t1-t0)



