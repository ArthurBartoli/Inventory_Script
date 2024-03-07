import time
import win32com.client as win32

def safe_write_to_excel(sheet, row, col, value):
    '''
    The script may write into an excel that is not fully opened, 
    so we retry and wait a second for it to boot up correctly.
    '''
    try:
        sheet.Cells(row, col).Value = value
    except win32.pywintypes.com_error:
        time.sleep(3)  # Wait a bit for Excel to be ready
        sheet.Cells(row, col).Value = value  # Retry writing

def data_writing_to_excel(sheet, data):
    '''Iterates over rows and col to safely write data into the excel,
    matching Excel's indexing.'''
    
    # Using start=1 to match Excel's 1-based indexing
    for i, row in enumerate(data, start=1):  
            for j, value in enumerate(row, start=1):
                safe_write_to_excel(sheet, i, j, value)