import win32com.client as win32
import os

from toolbox.safe_write_to_excel import safe_write_to_excel
from toolbox.safe_write_to_excel import data_writing_to_excel
from toolbox.unwrap_json import unwrap_json
from toolbox.date_reader import date_reader, closest_date

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
        headers = ["type of workspace", "reports", "workspaces", "datasets", "users", "dataflows", "dashboards"]
        data = [
            headers,
            ["Personal group"],
            ["Users workspaces"]
        ]

        data[1].extend(len(self.export_data[key]["PersonalWorkspace"].keys()) for key in headers[1:])
        data[2].extend(len(self.export_data[key]["SharedWorkspace"].keys()) for key in headers[1:])

        # Write table
        data_writing_to_excel(sheet, data)

        # Create table
        table_range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(len(data), len(data[0])))
        table = sheet.ListObjects.Add(SourceType=win32.constants.xlSrcRange, XlListObjectHasHeaders=1)
        
    def datasets_and_sources(self):
        ## Datasets and their sources/gateway
        sheet2 = self.workbook.Worksheets.Add(After=self.workbook.Worksheets(self.workbook.Worksheets.Count))
        sheet2.Name = "Gateways & Data sources"  

        file_path = os.path.join(self.script_directory, 'DatasourceAndGateway.json')

        # Unwrap the data source export
        json_data = unwrap_json(file_path)

        # Creating data structure
        header = ["DataSourceType", "DatasetId", "DatasetName", "ConnectionDetails", "DataSourceId", "ConnectionString", "GatewayId", "DataSourceName"]
        data = [header]
        
        # Function to process each dataset
        def process_dataset(item):
            return [str(item[k]) if k == "ConnectionDetails" else item[k] for k in header]

        # Build table
        try :
            for item in json_data:
                data.append(process_dataset(item))
        # If the tenant is REALLY empty, a TypeError is returned.
        # In that case, json_data is the lowest-level dict in the whole export data structure
        except TypeError:
            data.append(process_dataset(json_data))

        # Write table
        data_writing_to_excel(sheet2, data)

        # Create Excel table from the data range
        table_range = sheet2.Range(sheet2.Cells(1, 1), sheet2.Cells(len(data), len(header)))
        table = sheet2.ListObjects.Add(SourceType=win32.constants.xlSrcRange, Source=table_range, XlListObjectHasHeaders=win32.constants.xlYes)
        
    def not_empty_personal_workspaces(self):
        '''
        Detects the personal workspaces which are not empty, makes a count of ressources
        and gets the related UPN.
        '''
        
        ## Datasets and their sources/gateway
        sheet3 = self.workbook.Worksheets.Add(After=self.workbook.Worksheets(self.workbook.Worksheets.Count))
        sheet3.Name = "Personal workspaces"  

        # Build data structure
        header = ["UPN", "# Dashboards", "# Reports", "# Datasets", "Workspace Name", "Date of last modification", "UPN suffix"]
        data = [header]

        # Access personal workspaces data
        personal_workspaces = self.export_data["search"]["PersonalWorkspace"]
        
        # Build data rows for non-empty workspaces
        for workspace in personal_workspaces.values():
            if any(workspace[key] for key in ["reports", "dashboards", "datasets"]):
                upn = next(iter(workspace['users'].keys()), "NO UPN AVAILABLE") # Yes, sometimes the API returns no users...
                
                try:
                    allCreatedDate = []
                    for dataset in workspace["datasets"].values():
                        allCreatedDate.append(date_reader(dataset["CreatedDate"]))
                    lastCreationDate = closest_date(allCreatedDate)
                except ValueError:
                    lastCreationDate = "NA"
                
                row = [
                    upn,
                    len(workspace["dashboards"]),
                    len(workspace["reports"]),
                    len(workspace["datasets"]),
                    workspace["workspaceName"],
                    lastCreationDate,
                    upn.split('.')[-1]
                ]
                data.append(row)
        
        # Write table
        data_writing_to_excel(sheet3, data)

        # Create table
        table_range = sheet3.Range(sheet3.Cells(1, 1), sheet3.Cells(len(data), len(header)))
        table = sheet3.ListObjects.Add(SourceType=win32.constants.xlSrcRange, Source=table_range, XlListObjectHasHeaders=win32.constants.xlYes)

    def shared_workspaces_content(self):
        '''
        Counts the workspaces and their ressources. 
        A list of users and their respective rights is provided.
        '''
    
        # Create a new worksheet for shared workspaces
        sheet4 = self.workbook.Worksheets.Add(After=self.workbook.Worksheets(self.workbook.Worksheets.Count))
        sheet4.Name = "Shared workspaces"  

        # Display the number of shared workspaces
        workspace_count = f"There are {len(self.export_data['search']['SharedWorkspace'])} shared workspaces"
        safe_write_to_excel(sheet4, 1, 1, workspace_count)

        current_row = 3  # Start writing data from row 3

        # Iterate through each shared workspace
        for workspace in self.export_data["search"]["SharedWorkspace"].values():
            safe_write_to_excel(sheet4, current_row, 1, workspace["workspaceName"])
            current_row += 1

            # Headers for the users' list
            header = ["UPN", "Identifier", "Role", "Principal type"]
            for i, head in enumerate(header, start=1):
                safe_write_to_excel(sheet4, current_row, i, head)
            current_row += 1

            # Users and their details
            for user, details in workspace["users"].items():
                princType = "User" if details["PrincipalType"] == 2 else "Group" if details["PrincipalType"] == 1 else "Unknown"
                accessRight = details.get("AccessRight", "N/A THROUGH API")
                values = [user, details.get("Identifier", ""), accessRight, princType]
                for i, value in enumerate(values, start=1):
                    safe_write_to_excel(sheet4, current_row, i, value)
                current_row += 1

            # Space before resource counts
            current_row += 1

            # Counts for reports, datasets, and dataflows
            for item in ["reports", "datasets", "dataflows"]:
                safe_write_to_excel(sheet4, current_row, 1, f"{len(workspace[item])} {item}")
                current_row += 1

            # Additional space after each workspace's details
            current_row += 1

                        
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
        /!\ The API might return wrong information about the gateway
        '''
        
        sheet5 = self.workbook.Worksheets.Add(After=self.workbook.Worksheets(self.workbook.Worksheets.Count))
        sheet5.Name = "Dataset info"  

        header = ["DatasetId", "DatasetName", "Gateway Required"]
        data = [header]

        datasets = self.export_data["datasets"]["SharedWorkspace"]
        # Build data
        data.extend([[k, v["Name"], v["IsOnPremGatewayRequired"]] for k, v in datasets.items()])

        # Write table
        data_writing_to_excel(sheet5, data)

        # Create table
        table_range = sheet5.Range(sheet5.Cells(1, 1), sheet5.Cells(len(data), len(data[0])))
        table = sheet5.ListObjects.Add(SourceType=win32.constants.xlSrcRange, Source=table_range, XlListObjectHasHeaders=win32.constants.xlYes)
        
    def reports_info(self):
        '''All reports and their relative information'''
        
        # Add worksheet
        sheet6 = self.workbook.Worksheets.Add(After=self.workbook.Worksheets(self.workbook.Worksheets.Count))
        sheet6.Name = "Report info"  
        
        # Build data structure
        header = ["ReportId", "ReportName"]
        data = [header]

        reports = self.export_data["reports"]["SharedWorkspace"]
        # Build data
        data.extend([[k, v["Name"]] for k, v in reports.items()])

        # Write table
        data_writing_to_excel(sheet6, data)

        # Create table
        table_range = sheet6.Range(sheet6.Cells(1, 1), sheet6.Cells(len(data), len(data[0])))
        table = sheet6.ListObjects.Add(SourceType=win32.constants.xlSrcRange, Source=table_range, XlListObjectHasHeaders=win32.constants.xlYes)
        
    def iam_done(self):
        # We just pick the last two character of any UPN
        country_initials = list(self.export_data["users"]["PersonalWorkspace"].keys())[0][-2:]

        self.workbook.SaveAs(os.path.join(self.export_directory, country_initials + '_inventory.xlsx'))
        self.excel_app.Quit()
