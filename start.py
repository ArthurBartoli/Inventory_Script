'''
The workflow is in 3 steps :
    1. Run a script to export a json of the tenant
    2. Unwrap the json and drill down the content while storing it in variables
        * One dict for each workspace, report, user and dataset
        * One dict of many nested dicts that keeps track of the mining
    3. Organise this data into seperate excel sheets and name it accordingly
    
EXCEL NEEDS TO BE ALREADY OPENED FIRST
'''

import subprocess
import os
import time
import json
import win32com.client as win32
from pprint import pprint

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


# Set the result storage
keys = ["search", "reports", "workspaces", "datasets", "users", "dataflows", "dashboards"]
res_final = {k: {"PersonalWorkspace": {}, "SharedWorkspace": {}} for k in keys}



### Search the workspaces
for t in range(len(json_data)):
    workspaceId = str(json_data[t]["Id"])
    workspaceName = str(json_data[t]["Name"])
    
    # The workspace that is being searched is a personal workspace
    if "PersonalWorkspace" in workspaceName or "My workspace" in workspaceName:
        
        res_final["workspaces"]["PersonalWorkspace"][workspaceId] = {'workspaceName': workspaceName}
        res_final["search"]["PersonalWorkspace"][workspaceId] = {'workspaceName': workspaceName}
        
        # Set data for browsing
        keys = ["Reports", "Dashboards", "Datasets", "Users", "Dataflows"]
        data = {key: json_data[t][key] for key in keys}
        
        res_final["search"]["PersonalWorkspace"][workspaceId]["reports"] = {}
        ## We search through the reports
        if data["Reports"]:
            for report in data["Reports"]:
                
                
                
                # We keep the name for the entry
                tmpName = report["Id"]
                del report["Id"]
                
                # We create an entry for each report
                res_final["search"]["PersonalWorkspace"][workspaceId]["reports"][tmpName] = report
                if tmpName not in res_final["reports"]["PersonalWorkspace"]:
                    res_final["reports"]["PersonalWorkspace"][tmpName] = report
        
        
        res_final["search"]["PersonalWorkspace"][workspaceId]["dashboards"] = {}
        ## We search through the dashboards
        if data["Dashboards"]:
            for dashboard in data["Dashboards"]:
                
                # We keep the name for the entry
                tmpName = dashboard["Id"]
                del dashboard["Id"]
                
                # We create an entry for each report
                res_final["search"]["PersonalWorkspace"][workspaceId]["dashboards"][tmpName] = dashboard
                if tmpName not in res_final["dashboards"]["PersonalWorkspace"]:
                    res_final["dashboards"]["PersonalWorkspace"][tmpName] = dashboard
        
        
        res_final["search"]["PersonalWorkspace"][workspaceId]["datasets"] = {}
        ## We search through the datasets
        if data["Datasets"]:
            for dataset in data["Datasets"]:
                
                # We keep the name for the entry
                tmpName = dataset["Id"]
                del dataset["Id"]
                
                # We create an entry for each report
                res_final["search"]["PersonalWorkspace"][workspaceId]["datasets"][tmpName] = dataset
                if tmpName not in res_final["datasets"]["PersonalWorkspace"]:
                    res_final["datasets"]["PersonalWorkspace"][tmpName] = dataset
            
        
        res_final["search"]["PersonalWorkspace"][workspaceId]["users"] = {}
        ## We search through the users
        if data["Users"] :
            for user in data["Users"]:
                
                # We keep the name for the entry
                tmpName = user["UserPrincipalName"]
                del user["UserPrincipalName"]
                
                # We create an entry for each report
                res_final["search"]["PersonalWorkspace"][workspaceId]["users"][tmpName] = user
                if tmpName not in res_final["users"]["PersonalWorkspace"]:
                    del user["AccessRight"]
                    res_final["users"]["PersonalWorkspace"][tmpName] = user
            
            
        res_final["search"]["PersonalWorkspace"][workspaceId]["dataflows"] = {}
        ## We search through the dataflows
        if data["Dataflows"] :
            for dataflow in data["Dataflows"]:
                
                # We keep the name for the entry
                tmpName = dataflow["Id"]
                del dataflow["Id"]
                
                # We create an entry for each report
                res_final["search"]["PersonalWorkspace"][workspaceId]["dataflows"][tmpName] = dataflow
                if tmpName not in res_final["dataflows"]["PersonalWorkspace"]:
                    res_final["dataflows"]["PersonalWorkspace"][tmpName] = dataflow
                    
            
    # The workspace that is being searched is a shared workspace
    else: 
        res_final["workspaces"]["SharedWorkspace"][workspaceId] = {'workspaceName': workspaceName}
        res_final["search"]["SharedWorkspace"][workspaceId] = {'workspaceName': workspaceName}
        
        # Set data for browsing
        keys = ["Reports", "Dashboards", "Datasets", "Users", "Dataflows"]
        data = {key: json_data[t][key] for key in keys}
        
        res_final["search"]["SharedWorkspace"][workspaceId]["reports"] = {}
        ## We search through the reports
        if data["Reports"]:
            for report in data["Reports"]:
                
                # We keep the name for the entry
                tmpName = report["Id"]
                del report["Id"]
                
                # We create an entry for each report
                res_final["search"]["SharedWorkspace"][workspaceId]["reports"][tmpName] = report
                if tmpName not in res_final["reports"]["SharedWorkspace"]:
                    res_final["reports"]["SharedWorkspace"][tmpName] = report
        
        
        res_final["search"]["SharedWorkspace"][workspaceId]["dashboards"] = {}
        ## We search through the dashboards
        if data["Dashboards"]:
            for dashboard in data["Dashboards"]:
                
                # We keep the name for the entry
                tmpName = dashboard["Id"]
                del dashboard["Id"]
                
                # We create an entry for each report
                res_final["search"]["SharedWorkspace"][workspaceId]["dashboards"][tmpName] = dashboard
                if tmpName not in res_final["dashboards"]["SharedWorkspace"]:
                    res_final["dashboards"]["SharedWorkspace"][tmpName] = dashboard
        
        
        res_final["search"]["SharedWorkspace"][workspaceId]["datasets"] = {}
        ## We search through the datasets
        if data["Datasets"]:
            for dataset in data["Datasets"]:
                
                # We keep the name for the entry
                tmpName = dataset["Id"]
                del dataset["Id"]
                
                # We create an entry for each report
                res_final["search"]["SharedWorkspace"][workspaceId]["datasets"][tmpName] = dataset
                if tmpName not in res_final["datasets"]["SharedWorkspace"]:
                    res_final["datasets"]["SharedWorkspace"][tmpName] = dataset
            
        
        res_final["search"]["SharedWorkspace"][workspaceId]["users"] = {}
        ## We search through the users
        if data["Users"] :
            for user in data["Users"]:
                
                # We keep the name for the entry
                tmpName = user["UserPrincipalName"]
                del user["UserPrincipalName"]
                
                # We create an entry for each report
                res_final["search"]["SharedWorkspace"][workspaceId]["users"][tmpName] = user
                if tmpName not in res_final["users"]["SharedWorkspace"]:
                    del user["AccessRight"]
                    res_final["users"]["SharedWorkspace"][tmpName] = user
            
            
        res_final["search"]["SharedWorkspace"][workspaceId]["dataflows"] = {}
        ## We search through the dataflows
        if data["Dataflows"] :
            for dataflow in data["Dataflows"]:
                
                # We keep the name for the entry
                tmpName = dataflow["Id"]
                del dataflow["Id"]
                
                # We create an entry for each report
                res_final["search"]["SharedWorkspace"][workspaceId]["dataflows"][tmpName] = dataflow
                if tmpName not in res_final["dataflows"]["SharedWorkspace"]:
                    res_final["dataflows"]["SharedWorkspace"][tmpName] = dataflow


## Formatting the excel sheet

## Summary stats
# Crate new excel instance
excel = win32.gencache.EnsureDispatch('Excel.Application')
# Make Excel visible
excel.Visible = True
# New worksheet
workbook = excel.Workbooks.Add()
# Select the first worksheet
sheet = workbook.Worksheets(1)
sheet.Name = "Workspaces"  
# Create data
data = [
    ["type of workspace", "reports", "workspaces", "datasets", "users", "dataflows", "dashboards"],
    ["Personal group"],
    ["Users workspaces"]
]

for k in data[0][1:]:
    data[1].append(len(res_final[k]["PersonalWorkspace"].keys()))
    data[2].append(len(res_final[k]["SharedWorkspace"].keys()))

for i, row in enumerate(data):
    for j, value in enumerate(row):
        try:
            sheet.Cells(i + 1, j + 1).Value = value
        except win32.pywintypes.com_error:
            time.sleep(1)
            sheet.Cells(i + 1, j + 1).Value = value

# Create table
table_range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(len(data), len(data[0])))
table = sheet.ListObjects.Add(SourceType=win32.constants.xlSrcRange, XlListObjectHasHeaders=1)

##############

## Datasets and their sources/gateway
sheet2 = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
sheet2.Name = "Gateways & Data sources"  

file_path = os.path.join(script_directory, 'DatasourceAndGateway.json')

# Unwrap the data source export
with open(file_path, 'r', encoding='utf-16') as json2_file:
    json_str = json2_file.read()
    json_data = json.loads(json_str)
    print(json_data)

data = [
    ["DataSourceType", "DatasetId", "DatasetName", "ConnectionDetails", "DataSourceId", "ConnectionString", "GatewayId", "DataSourceName"]
]

# Build table
try :
    for i, item in enumerate(json_data):
        data.append([])
        for k in data[0]:
            if k == "ConnectionDetails":
                data[i+1].append(str(item[k]))
            else: 
                data[i+1].append(item[k])
# If the tenant is REALLY empty, json_data is the lowest-level dict in the whole export data structure
except TypeError:
    for i, item in enumerate(json_data):
        data.append([])
        for k in data[0]:
            if k == "ConnectionDetails":
                data[i+1].append(str(json_data[k]))
            else: 
                data[i+1].append(json_data[k])

# Write table
for i, row in enumerate(data):
    for j, value in enumerate(row):
        try:
            sheet2.Cells(i + 1, j + 1).Value = value
        except win32.pywintypes.com_error:
            time.sleep(3)
            sheet2.Cells(i + 1, j + 1).Value = value

# Create table
table_range = sheet2.Range(sheet2.Cells(1, 1), sheet2.Cells(len(data), len(data[0])))
table = sheet2.ListObjects.Add(SourceType=win32.constants.xlSrcRange, XlListObjectHasHeaders=1)

##############

## Datasets and their sources/gateway
sheet3 = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
sheet3.Name = "Personal workspaces"  

data = [
    ["UPN", "# Dashboards", "# Reports", "# Datasets", "Workspace Name"]
]


personal_workspaces = res_final["search"]["PersonalWorkspace"]
# Build data
for k in personal_workspaces.keys():
    pprint(personal_workspaces[k]['users'])
    if personal_workspaces[k]["reports"] or personal_workspaces[k]["dashboards"] or personal_workspaces[k]["datasets"]: # If there are any reports/datasets in this personal workspace
        try:
            data.append([
                list(personal_workspaces[k]['users'].keys())[0], # Get user UPN
                len(personal_workspaces[k]["dashboards"].keys()), # Get number of dashboards
                len(personal_workspaces[k]["reports"].keys()), # Get number of dashboards
                len(personal_workspaces[k]["datasets"].keys()), # Get number of dashboards
                personal_workspaces[k]["workspaceName"]
            ])
        except IndexError: # Yes, sometimes the API returns no users
            data.append([
                "NO USER REGISTERED", # Get user UPN
                len(personal_workspaces[k]["dashboards"].keys()), # Get number of dashboards
                len(personal_workspaces[k]["reports"].keys()), # Get number of dashboards
                len(personal_workspaces[k]["datasets"].keys()), # Get number of dashboards
                personal_workspaces[k]["workspaceName"]
            ])

# Write table
for i, row in enumerate(data):
    for j, value in enumerate(row):
        try:
            sheet3.Cells(i + 1, j + 1).Value = value
        except win32.pywintypes.com_error:
            time.sleep(1)
            sheet3.Cells(i + 1, j + 1).Value = value

# Create table
table_range = sheet3.Range(sheet3.Cells(1, 1), sheet3.Cells(len(data), len(data[0])))
table = sheet3.ListObjects.Add(SourceType=win32.constants.xlSrcRange, XlListObjectHasHeaders=1)

#############

sheet4 = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
sheet4.Name = "Shared workspaces"  

nb_workspaces = len(res_final["search"]["SharedWorkspace"])
str_nb_workspaces = f"There are {nb_workspaces} shared workspaces" 

data = [
    [str_nb_workspaces],
    []
]
header = ["UPN", "Identifier", "Role", "Principal type"]

shared_workspaces = res_final["search"]["SharedWorkspace"]

lengths = [] # Starting position for table creation
# Build data
for k in shared_workspaces.keys():
    data.append([shared_workspaces[k]["workspaceName"]])
    data.append(header)
    for i, user in enumerate(shared_workspaces[k]["users"]):

        try:
            princType = (shared_workspaces[k]["users"][user]["PrincipalType"] == 2) * "User" + (shared_workspaces[k]["users"][user]["PrincipalType"] == 1) * "Group"
            data.append([
                user,
                shared_workspaces[k]["users"][user]["Identifier"],
                shared_workspaces[k]["users"][user]["AccessRight"],
                princType
            ])
        except KeyError:
            princType = (shared_workspaces[k]["users"][user]["PrincipalType"] == 2) * "User" + (shared_workspaces[k]["users"][user]["PrincipalType"] == 1) * "Group"
            data.append([
                user,
                shared_workspaces[k]["users"][user]["Identifier"],
                ["N/A THROUGH API"],
                princType
            ])
    lengths.append(i+1) # +1 for headers
    data.append([])
    for item in ["reports", "datasets", "dataflows"]:
        count = len(shared_workspaces[k][item])
        data.append([f"{count} {item}"])
    data.append([])

# Write table
for i, row in enumerate(data):
    for j, value in enumerate(row):
        try:
            sheet4.Cells(i + 1, j + 1).Value = value
        except win32.pywintypes.com_error:
            time.sleep(1)
            sheet4.Cells(i + 1, j + 1).Value = value


# Create table
# TODO: Faire fonctionner ça
# last_stop = 0
# for table_length in lengths:
#     if last_stop == 0:
#         table_start = 3 
#     else: table_start = last_stop + 7
#     print(f"Le tableau va de {table_start},1 à {table_start+table_length},3")
#     table_range = sheet4.Range(sheet4.Cells(table_start, 1), sheet4.Cells(table_start + table_length, 3))
#     table = sheet4.ListObjects.Add(SourceType=win32.constants.xlSrcRange, XlListObjectHasHeaders=1)
#     last_stop = table_start + table_length   

##############

## Datasets and their sources/gateway
sheet5 = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
sheet5.Name = "Dataset info"  

data = [
    ["DatasetId", "DatasetName", "Gateway Required"]
]

datasets = res_final["datasets"]["SharedWorkspace"]
# Build data
for k in datasets.keys():
    data.append([k, datasets[k]["Name"], datasets[k]["IsOnPremGatewayRequired"]])

# Write table
for i, row in enumerate(data):
    for j, value in enumerate(row):
        try:
            sheet5.Cells(i + 1, j + 1).Value = value
        except win32.pywintypes.com_error:
            time.sleep(1)
            sheet5.Cells(i + 1, j + 1).Value = value

# Create table
table_range = sheet5.Range(sheet5.Cells(1, 1), sheet5.Cells(len(data), len(data[0])))
table = sheet5.ListObjects.Add(SourceType=win32.constants.xlSrcRange, XlListObjectHasHeaders=1)

##############

## Reports
sheet6 = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
sheet6.Name = "Report info"  

data = [
    ["ReportId", "ReportName"]
]

reports = res_final["reports"]["SharedWorkspace"]
# Build data
for k in reports.keys():
    data.append([k, reports[k]["Name"]])

# Write table
for i, row in enumerate(data):
    for j, value in enumerate(row):
        try:
            sheet6.Cells(i + 1, j + 1).Value = value
        except win32.pywintypes.com_error:
            time.sleep(1)
            sheet6.Cells(i + 1, j + 1).Value = value

# Create table
table_range = sheet6.Range(sheet6.Cells(1, 1), sheet6.Cells(len(data), len(data[0])))
table = sheet6.ListObjects.Add(SourceType=win32.constants.xlSrcRange, XlListObjectHasHeaders=1)

##############

# We just pick the last two character of any UPN
country_initials = list(res_final["users"]["PersonalWorkspace"].keys())[0][-2:]

workbook.SaveAs(os.path.join(export_directory, country_initials + '_inventory.xlsx'))
excel.Quit()

print("Done.")