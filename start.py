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

from toolbox.scan_export import scan_export
from toolbox.ExcelManager import ExcelManager
from toolbox.run_powershell_script import run_powershell_script
from toolbox.unwrap_json import unwrap_json

                            
# Defining paths
script_directory = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_directory, "json_extract.ps1")
export_directory = os.path.join(script_directory, "export")
export_path = os.path.join(script_directory, 'export.json')

# run_powershell_script(file_path)
    
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