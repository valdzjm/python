import xlwings as xw
import pandas as pd
from tkinter import filedialog
import shutil
from datetime import datetime
import pyautogui as pg

#Working With the Master Excel File

fpmaster = r"C:\Users\jvaldez21\OneDrive - DXC Production\Tracker\DXC ChgM Change Review Tracker.xlsx"
#fptest =  r"C:\Users\jvaldez21\OneDrive - DXC Production\Tracker\Testing area\Tracker Test.xlsx"
fpnew = filedialog.askopenfilename()

master_wb = xw.Book(fpmaster) #Master excel file

master_sheets = master_wb.sheets


#Reading new data
newdata_wb = xw.Book(fpnew)
newdata_sheets = newdata_wb.sheets

#Converting datatime into string
for cell in master_sheets[0].used_range, master_sheets[1].used_range, master_sheets[2].used_range, master_sheets[3].used_range:
    if isinstance(cell.value, datetime):
        cell.value = cell.value.strftime('%m-%d-%Y %H:%M:%S')  

#Getting new data
new_data_raw = newdata_wb.sheets[0].range('A2').expand().value
new_data_raw

#Lists for Requested for Approval for each sheet
decomDate = master_sheets[3].range('AB:AB').expand('down').value 
icabDate = master_sheets[0].range('AB:AB').expand('down').value 
acabDate = master_sheets[1].range('AB:AB').expand('down').value 
tpfDate = master_sheets[2].range('AB:AB').expand('down').value 

for D in master_sheets: #For Decom
    decom = [i[0:20] + i[21::21] for i in new_data_raw if i[19] == 'Decommission' and i[21] not in decomDate]
    dnewrow = master_wb.sheets[3].range('A1').end('down').row + 1
    master_wb.sheets[3].range(dnewrow, 8).value = decom

master_wb.sheets[3].range('H2').expand('right').copy()
master_wb.sheets[3].range('H2').expand().paste(paste='formats')
# master_wb.sheets[3].range('AF59').expand('down').number_format = 'mm-dd-yyyy hh:mm:ss'

for I in master_sheets: #For ICAB
    icab = [i[0:20] + i[21::21] for i in new_data_raw if i[19] != 'Decommission' and i[3] == 'DXC' and i[21] not in icabDate]
    inewrow = master_wb.sheets[0].range('A1').end('down').row + 1
    master_wb.sheets[0].range(inewrow, 8).value = icab
    
  
master_wb.sheets[0].range('H2').expand('right').copy()
master_wb.sheets[0].range('H2').expand().paste(paste='formats')
# master_wb.sheets[0].range('AE76').expand('down').number_format = 'mm-dd-yyyy hh:mm:ss'

for T in master_sheets: #For TPF
    tpf = [i[0:20] + i[21::21]  for i in new_data_raw if i[19] != 'Decommission' and i[3] == 'TPF/PSS' and i[21] not in tpfDate]
    tnewrow = master_wb.sheets[2].range('A1').end('down').row + 1
    master_wb.sheets[2].range(tnewrow, 8).value = tpf
 
master_wb.sheets[2].range('H2').expand('right').copy()
master_wb.sheets[2].range('H2').expand().paste(paste='formats')
# master_wb.sheets[3].range('AE27').expand('down').number_format = 'mm-dd-yyyy hh:mm:ss'

for A in master_sheets: #For ACAB
    acab = [i[0:20] + i[21::21]  for i in new_data_raw if i[19] != 'Decommission' and i[3] != 'DXC' and i[3] !='TPF/PSS' and i[21] not in acabDate ] 
    anewrow = master_wb.sheets[1].range('A1').end('down').row + 1
    master_wb.sheets[1].range(anewrow, 8).value = acab

master_wb.sheets[1].range('H2').expand('right').copy()
master_wb.sheets[1].range('H2').expand().paste(paste='formats')
# master_wb.sheets[1].range('AE23').expand('down').number_format = 'mm-dd-yyyy hh:mm:ss'

#convert every datetime to string
for cell in master_sheets[0].used_range, master_sheets[1].used_range, master_sheets[2].used_range, master_sheets[3].used_range:
    if isinstance(cell.value, datetime):
        cell.value = cell.value.strftime('%m-%d-%Y %H:%M:%S')  

master_wb.save()

pg.alert('Success updating the Tracker', 'Good Morning Handsome')