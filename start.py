'''
The workflow is in 3 steps :
    1. Run a script to export a json of the tenant
    2. Unwrap the json and drill down the content while sorting it
        * One dict for each workspace, report, user and dataset
        * One dict of many nested dicts that keeps track of the mining
    3. Organise this data into seperate excel sheets and name it accordingly
    
EXCEL NEEDS TO BE ALREADY OPENED FIRST
'''

import os
import sys

from toolbox.scan_export import scan_export
from toolbox.ExcelManager import ExcelManager
from toolbox.run_powershell_script import run_powershell_script
from toolbox.unwrap_json import unwrap_json

# Split the arguments list in pairs
all_args = sys.argv[1:] # We skip the first argument which is the file name
n = 2
list_args = [all_args[i * n:(i + 1) * n] for i in range((len(all_args) + n - 1) // n )]  
print(list_args)

ps_bool = False

for arg in list_args:
    match arg[0]:
        # Does the script need to run the data export ?
        case '--export' | "-e":
            if arg[1] not in ["yes", "y", "no", "n"]:
                print("Invalid argument. Please enter [y]es or [n]o.")
                pass
            else:
                export_bool = (arg[1] == "y" or arg[1] == "yes")
        case _:
            print("Please enter a valid option.")
                            
# Defining paths
script_directory = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_directory, "json_extract.ps1")
export_directory = os.path.join(script_directory, "export")
export_path = os.path.join(script_directory, 'export.json')

if export_bool: run_powershell_script(file_path)
    
json_data = unwrap_json(export_path)
res_final = scan_export(json_data)


## Formatting the excel sheet
ExcelM = ExcelManager(res_final)

ExcelM.summary_stats()
ExcelM.datasets_and_sources()
ExcelM.not_empty_personal_workspaces()
ExcelM.shared_workspaces_content()
ExcelM.datasets_info()
ExcelM.reports_info()
ExcelM.iam_done()

print("Done.")