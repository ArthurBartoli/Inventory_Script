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

# If the list of arguments is empty, we enter a dummy one
if not all_args: list_args = [["no arguments"]]

print(list_args)
export_bool = True
for arg in list_args:
    match arg[0]:
        case '--export' | "-e": # Does the script need to run the data export ?
            arg_value = arg[1].lower()
            if arg_value not in ["yes", "y", "no", "n"]:
                print("Invalid argument. Please enter [y]es or [n]o.")
                pass
            else:
                export_bool = (arg_value == "y" or arg_value == "yes")
        case "--help" | "-h":
            print("You have entered the help command, all further commands are ignored.")
            print("Here is a list of available options :")
            print("\t* --export|-e [Y]es|[N]o")
            print("\t\tWhether or not to run the powershell script which exports all data")
            print("\t\tfrom tenant. If the export is already done once, it is not necessary.")
        case "no arguments":
            pass
        case _:
            print("## Could not recognise the argument, please enter a valid option. ##")

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