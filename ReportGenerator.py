######################################################################################################
# Creates the NCT report, cleans it up by commonizing Root Cause Origin/Statement fields
# Inputs: String for query. Main fields being: user, password, query type
# Outputs: Dictionary with the different fields for each NCT in the system.
#         Will also output any errors during NCT Fetch
# Created by: Daniel Gomez
# Date: 04/19/2019
# Revision: 01 [Initial Release]
# [TODO] Make variables for query items and finish documentation
######################################################################################################

import time
from subprocess import check_output

query_string = "im issues " \
               "--user='daniel.gomez' " \
               "--password=\"Dg207433955Veoneer\" " \
               "--queryDefinition='((field[\"FA_Issue Category\"] = \"Warranty\") and" \
                                " (field[\"AI_Customer (pick list)\"]=\"Ford\",\"GM\",\"Nissan\",\"HKMC\") and" \
                                " (field[\"Project\"]=\"/Analysis/RCS Analysis\") and" \
                                " (field[\"Created Date\"] in the last 12 months)" \
                                ")' " \
               "--fields=\"ID\",\"FTF_Root Cause Origin\","",\"FTF_Root Cause Statement\",\"Created Date\"," \
               "\"FA_Customer Reject Date\"," \
               "\"AI_Date of Arrival\",\"FA_Issue Category\",\"AI_Country\",\"Summary\",\"Project\",\"CQE\"," \
               "\"Assigned User\",\"AI_NCM number\",\"FA_NCM Open Date\",\"AI_Fault Type Others\",\"AI_Km\"," \
               "\"FTF_Supplier Name Short\",\"FA_Supplier Name Short FVA\",\"FTF_Device Name\",\"FA_Device Name FVA\"," \
               "\"FA_Customer Failure Code\",\"AI_Customer (pick list)\",\"FA_Dealer\"," \
               "\"FA_Dealer Country\",\"FA_Dealer State [US Only]\",\"AI_Error discovered (place)\"," \
               "\"Error discovered date\",\"AI_Part No\",\"AI_Customer Part No\",\"AI_CSN number\"," \
               "\"AI_Chassi number\",\"AI_Production date\",\"FA_Product Line\",\"FA_Product Category\"," \
               "\"FA_Product Family\",\"FA_Vehicle Line (Code)\",\"FA_MY (Model Year)\",\"Vehicle registration date\"" \
               ",\"Vehicle production date\"" \
               " --fieldsDelim=\"$\""


def report_creation (query_string = query_string):
    header_string = ["ID", "Root Cause Origin", "Root Cause Statement", "Created Date", "Reject Date",
                     "Arrival Date", "Issue Category", "Veoneer Facility", "Summary", "Project", "Veoneer Owner",
                     "Assigned User", "NCM Number", "NCM Open Date", "8D", "Odometer (km)", "Supplier Name",
                     "Device Name", "Customer Failure Code", "Customer / OEM", "Dealer",
                     "Dealer Country", "Dealer State", "Customer Plant / Failure Location", "Date of Failure",
                     "Veoneer PN", "Customer PN", "CSN/Serial Number", "VIN Number", "Veoneer Production Date",
                     "Product Line", "Product Category", "Product Family", "Vehicle Line (Code)", "Model Year",
                     "Warranty Start Date", "Vehicle Production Date"]
    report_string = (check_output(query_string).decode(errors='ignore'))
    if "Reconnecting" in report_string:
        report_creation(query_string)
    NCT_Master = []
    for line in report_string.splitlines():
        NCT_Event = []
        for item in line.split("$"):
            NCT_Event.append(item)
        # Commonize Supplier Name TODO optimize
        if NCT_Event[15] or NCT_Event[16] != "":
            if NCT_Event[15] == "":
                NCT_Event[15] = NCT_Event[16]

        # Commonize Device Name TODO optimize
        if NCT_Event[17] or NCT_Event[18] != "":
            if NCT_Event[17] == "":
                NCT_Event[17] = NCT_Event[18]

        # Perform 'pop' of unused indexes
        NCT_Event.pop(16)
        NCT_Event.pop(17)
        NCT_Master.append(NCT_Event)
    NCT_Master.insert(0, header_string)
    return NCT_Master

########################### NCT Event List Index ###########################
# Index 0 = ID
# Index 1 = Root Cause Origin
# Index 2 = Root Cause Statement
# Index 3 = Reject Date
# Index 4 = Date of Arrival
# Index 5 = Issue Category
# Index 6 = Veoneer Facility
# Index 7 = Summary
# Index 8 = Project
# Index 9 = Veoneer Owner
# Index 10 = Assigned User
# Index 11 = NCM Number
# Index 12 = NCM Open Date
# Index 13 = 8D [TODO] Improve
# Index 14 = Odometer (km)
# Index 15 = Supplier Name
# Index 16 = Extra Supplier Name
# Index 17 = Device Name
# Index 18 = Extra Device Name
# Index
# Index
# Index
# Index
# Index
###############################################################################





















