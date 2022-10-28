from csv import excel_tab
import win32com.client
xl = win32com.client.Dispatch("Excel.Application")  #instantiate excel app
import time
from importlib.machinery import PathFinder


excel_path=r"....\VBA_Macro_4_sapR3.xlsm"

#Run SAP_REPPORT_1

wb = xl.Workbooks.Open(excel_path)

xl.Application.Run('VBA_Macro_4_sapR3.xlsm!Module1.SAP_REPPORT_1')
wb.Save()
xl.Application.Quit()
time.sleep(150)

#Run SAP_REPPORT_2

wb = xl.Workbooks.Open(excel_path)
xl.Application.Run('VBA_Macro_4_sapR3.xlsm!Module1.SAP_REPPORT_2')
wb.Save()
xl.Application.Quit() 

time.sleep(150)

#Run SAP_REPPORT_3

wb = xl.Workbooks.Open(excel_path)
xl.Application.Run('VBA_Macro_4_sapR3.xlsm!Module1.SAP_REPPORT_3')
wb.Save()
xl.Application.Quit()

time.sleep(150)

#Run SAP_REPPORT_4

wb = xl.Workbooks.Open(excel_path)
xl.Application.Run('VBA_Macro_4_sapR3.xlsm!Module1.SAP_REPPORT_4')
wb.Save()
xl.Application.Quit()

time.sleep(150)