import win32com.client as win32
import time
import os
import json

class ExcelManager:
    '''Generates the data from the export and files it in an excel using its API'''
    
    def __init__(self, data: dict):
        # Create new excel instance
        self.excel_app = win32.gencache.EnsureDispatch('Excel.Application')
        # Make Excel visible (because it's cool)
        self.excel_app.Visible = True
        self.export_data = data
        ## New worksheet
        self.workbook = self.excel_app.Workbooks.Add()
        ## Defining paths
        current_file_path = os.path.abspath(__file__)
        current_directory = os.path.dirname(current_file_path)
        parent_directory = os.path.normpath(os.path.join(current_directory, '..'))
        self.script_directory = parent_directory
        self.file_path = os.path.join(self.script_directory, "json_extract.ps1")
        self.export_directory = os.path.join(self.script_directory, "export")

    def summary_stats(self):
        '''
        Summary stats are counts of reports, workspaces, datasets, users, dataflows and dashboards.
        This counts is divided in personal groups (personal workspaces) and users workspaces (shared workspaces)
        '''

        # Select the first worksheet
        sheet = self.workbook.Worksheets(1)
        sheet.Name = "Workspaces"  
        
        # Build data structure
        data = [
            ["type of workspace", "reports", "workspaces", "datasets", "users", "dataflows", "dashboards"],
            ["Personal group"],
            ["Users workspaces"]
        ]

        for k in data[0][1:]:
            data[1].append(len(self.export_data[k]["PersonalWorkspace"].keys()))
            data[2].append(len(self.export_data[k]["SharedWorkspace"].keys()))

        for i, row in enumerate(data):
            for j, value in enumerate(row):
                try:
                    sheet.Cells(i + 1, j + 1).Value = value
                except win32.pywintypes.com_error:  
                    # The script may write into an excel that is not fully opened
                    # So we retry and wait a second for it to boot up correctly
                    time.sleep(1)
                    sheet.Cells(i + 1, j + 1).Value = value

        # Create table
        table_range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(len(data), len(data[0])))
        table = sheet.ListObjects.Add(SourceType=win32.constants.xlSrcRange, XlListObjectHasHeaders=1)
        
    def datasets_and_sources(self):
        ## Datasets and their sources/gateway
        sheet2 = self.workbook.Worksheets.Add(After=self.workbook.Worksheets(self.workbook.Worksheets.Count))
        sheet2.Name = "Gateways & Data sources"  

        file_path = os.path.join(self.script_directory, 'DatasourceAndGateway.json')

        # Unwrap the data source export
        with open(file_path, 'r', encoding='utf-16') as json2_file:
            json_str = json2_file.read()
            json_data = json.loads(json_str)
            print(json_data)

        # Creating data structure
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
        
    def not_empty_personal_workspaces(self):
        '''
        Detects the personal workspaces which are not empty, makes a count of ressources
        and gets the related UPN.
        '''
        
        ## Datasets and their sources/gateway
        sheet3 = self.workbook.Worksheets.Add(After=self.workbook.Worksheets(self.workbook.Worksheets.Count))
        sheet3.Name = "Personal workspaces"  

        # Build data structure
        data = [
            ["UPN", "# Dashboards", "# Reports", "# Datasets", "Workspace Name"]
        ]

        personal_workspaces = self.export_data["search"]["PersonalWorkspace"]
        # Build data
        for k in personal_workspaces.keys():
            if personal_workspaces[k]["reports"] or personal_workspaces[k]["dashboards"] or personal_workspaces[k]["datasets"]: # If there are any reports/datasets in this personal workspace
                try:
                    data.append([
                        list(personal_workspaces[k]['users'].keys())[0], # Get user UPN
                        len(personal_workspaces[k]["dashboards"].keys()), # Get number of dashboards
                        len(personal_workspaces[k]["reports"].keys()), # Get number of dashboards
                        len(personal_workspaces[k]["datasets"].keys()), # Get number of dashboards
                        personal_workspaces[k]["workspaceName"]
                    ])
                except IndexError: # Yes, sometimes the API returns no users...
                    data.append([
                        "NO UPN AVAILABLE", # Get user UPN
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

    def shared_workspaces_content(self):
        '''
        Counts the workspaces and their ressources. 
        A list of users and their respective rights is provided.
        '''
        
        sheet4 = self.workbook.Worksheets.Add(After=self.workbook.Worksheets(self.workbook.Worksheets.Count))
        sheet4.Name = "Shared workspaces"  

        nb_workspaces = len(self.export_data["search"]["SharedWorkspace"])
        str_nb_workspaces = f"There are {nb_workspaces} shared workspaces" 

        data = [
            [str_nb_workspaces],
            []
        ]
        header = ["UPN", "Identifier", "Role", "Principal type"]

        shared_workspaces = self.export_data["search"]["SharedWorkspace"]

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
        
    def datasets_info(self):
        '''
        All datasets, names and whether a gateway is required
        /!\ The API returns wrong info about the gateway...
        '''
        
        sheet5 = self.workbook.Worksheets.Add(After=self.workbook.Worksheets(self.workbook.Worksheets.Count))
        sheet5.Name = "Dataset info"  

        data = [
            ["DatasetId", "DatasetName", "Gateway Required"]
        ]

        datasets = self.export_data["datasets"]["SharedWorkspace"]
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
        
    def reports_info(self):
        '''All reports and their relative information'''
        sheet6 = self.workbook.Worksheets.Add(After=self.workbook.Worksheets(self.workbook.Worksheets.Count))
        sheet6.Name = "Report info"  

        data = [
            ["ReportId", "ReportName"]
        ]

        reports = self.export_data["reports"]["SharedWorkspace"]
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
        
    def iam_done(self):
        # We just pick the last two character of any UPN
        country_initials = list(self.export_data["users"]["PersonalWorkspace"].keys())[0][-2:]

        self.workbook.SaveAs(os.path.join(self.export_directory, country_initials + '_inventory.xlsx'))
        self.excel_app.Quit()



                    
