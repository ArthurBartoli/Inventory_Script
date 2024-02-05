import win32com.client as win32
import time

class ExcelManager:
    
    def __init__(self):
        # Crate new excel instance
        self.excel_app = win32.gencache.EnsureDispatch('Excel.Application')

    def summary_stats(res_final: dict):
        '''
        Summary stats are counts of reports, workspaces, datasets, users, dataflows and dashboards.
        This counts is divided in personal groups (personal workspaces) and users workspaces (shared workspaces)
        '''

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
