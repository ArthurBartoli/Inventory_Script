'''
The workflow is in 3 steps :
    1. Run a script to export a json of the tenant
    2. Unwrap the json and drill down the content while sorting it
        * One dict for each workspace, report, user and dataset
        * One dict of many nested dicts that keeps track of the mining
    3. Organise this data into seperate excel sheets and name it accordingly
    
EXCEL NEEDS TO BE ALREADY OPENED FIRST
'''

import subprocess
import os
import json

from toolbox.workspace_search import scan_export  
from toolbox.ExcelManager import ExcelManager                                     

# Defining paths
script_directory = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_directory, "json_extract.ps1")
export_directory = os.path.join(script_directory, "export")

# Function definition as part of legacy (a decorator was used)
def run_powershell_script(script_path):
    command = ["powershell", "-ExecutionPolicy", "Bypass", "-File", script_path]
    subprocess.run(command)

run_powershell_script(file_path)

file_path = os.path.join(script_directory, 'export.json')
# Unwrap the json
with open(file_path, 'r', encoding='utf-16') as json1_file:
    json_str = json1_file.read()
    json_data = json.loads(json_str)

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