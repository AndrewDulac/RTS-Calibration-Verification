#  pip install pyinstaller
#  pyinstaller yourprogram.py
#  pyinstaller -F yourprogram.py

from easygui import fileopenbox
from openpyxl import load_workbook, Workbook
import matplotlib.pyplot as plt
from datetime import datetime, timedelta, time
import numpy as np

colIds = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L',
'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL']

def CellIdForValueInRow(val, row):
    i = 0
    colId = colIds[i]
    cellId = colId + str(row)
    focusedCell = ws1[cellId]
    while(focusedCell.value != None and val not in str(focusedCell.value)):
        i = i + 1
        colId = colIds[i]
        cellId = colId + str(row)
        focusedCell = ws1[cellId]
    return colId,row

def CellIdForValueInColumn(val, col):
    rowId = 1
    cellId = col + str(rowId)
    focusedCell = ws1[cellId]
    while(focusedCell.value != None and val not in str(focusedCell.value)):
        rowId = rowId + 1
        cellId = col+ str(rowId)
        focusedCell = ws1[cellId]
    return col,rowId

def PopulateArrayForHeader(header, col, row):
    newArray = []
    #newArray.append(header)
    rowId = row
    cellId = col + str(rowId + 1)
    focusedCell = ws1[cellId]
    while(focusedCell.value != None):
        if(type(focusedCell.value) == str):
            focusedCell.value = focusedCell.value.replace(" ", "")
        rowId = rowId + 1
        cellId = col+ str(rowId)
        newArray.append(focusedCell.value)
        focusedCell = ws1[cellId]
        
    return newArray

wb = load_workbook(filename = fileopenbox())

ws1 = wb.worksheets[0]

rowId = 1
cellId = 'A'+ str(rowId)
focusedCell = ws1[cellId]

col,row = CellIdForValueInColumn("Record", 'A')
TimeCol,TimeRow = CellIdForValueInRow("Time", row)
VaCol,VARow = CellIdForValueInRow("VA", row)
VbCol,VBRow = CellIdForValueInRow("VB", row)
VcCol,VCRow = CellIdForValueInRow("VC", row)
IaCol,IARow = CellIdForValueInRow("IA", row)
IbCol,IBRow = CellIdForValueInRow("IB", row)
IcCol,ICRow = CellIdForValueInRow("IC", row)
InCol,INRow = CellIdForValueInRow("IN", row)
ThdVaCol,ThdVaRow = CellIdForValueInRow("THDVA", row)
ThdVbCol,ThdVbRow = CellIdForValueInRow("THDVB", row)
ThdVcCol,ThdVcRow = CellIdForValueInRow("THDVC", row)
ThdIaCol,ThdIaRow = CellIdForValueInRow("THDIA", row)
ThdIbCol,ThdIbRow = CellIdForValueInRow("THDIB", row)
ThdIcCol,ThdIcRow = CellIdForValueInRow("THDIC", row)
ThdInCol,ThdInRow = CellIdForValueInRow("THDIN", row)

print(
"Headers Located: \n",
"\nTime: ",TimeCol,TimeRow,
"\nVoltage A: ",VaCol,VARow,
"\nVoltage B: ",VbCol,VBRow,
"\nVoltage C: ",VcCol,VCRow,
"\nCurrent A: ",IaCol,IARow,
"\nCurrent B: ",IbCol,IBRow,
"\nCurrent C: ",IcCol,ICRow,
"\nCurrent N: ",InCol,INRow,
"\nTHD Va: ",ThdVaCol,ThdVaRow,
"\nTHD Vb: ",ThdVbCol,ThdVbRow,
"\nTHD Vc: ",ThdVcCol,ThdVcRow,
"\nTHD Ia: ",ThdIaCol,ThdIaRow,
"\nTHD Ib: ",ThdIbCol,ThdIbRow,
"\nTHD Ic: ",ThdIcCol,ThdIcRow,
"\nTHD In: ",ThdInCol,ThdInRow,
"\nFilling arrays..."
)

Time = PopulateArrayForHeader("Time", TimeCol,TimeRow)
Va = PopulateArrayForHeader("Voltage A", VaCol,VARow)
Vb = PopulateArrayForHeader("Voltage B", VbCol,VBRow)
Vc = PopulateArrayForHeader("Voltage C", VcCol,VCRow)
Ia = PopulateArrayForHeader("Current A", IaCol,IARow)
Ib = PopulateArrayForHeader("Current B", IbCol,IBRow)
Ic = PopulateArrayForHeader("Current C", IcCol,ICRow)
In = PopulateArrayForHeader("Current N", InCol,INRow)
ThdVa = PopulateArrayForHeader("THD Va", ThdVaCol,ThdVaRow)
ThdVb = PopulateArrayForHeader("THD Vb", ThdVbCol,ThdVbRow)
ThdVc = PopulateArrayForHeader("THD Vc", ThdVcCol,ThdVcRow)
ThdIa = PopulateArrayForHeader("THD Ia", ThdIaCol,ThdIaRow)
ThdIb = PopulateArrayForHeader("THD Ib", ThdIbCol,ThdIbRow)
ThdIc = PopulateArrayForHeader("THD Ic", ThdIcCol,ThdIcRow)
ThdIn = PopulateArrayForHeader("THD In", ThdInCol,ThdInRow)

allData = [Time, Va, Vb, Vc, Ia, Ib, Ic, In, ThdVa, ThdVb, ThdVc, ThdIa, ThdIb, ThdIc, ThdIn]
print("\nArrays filled...")


plt.plot(Time, Va, Time, Ia)
plt.yticks(np.arange(min(Va), max(Va)+1,5.0))
timegapdt64 = np.arange(datetime.strptime(min(Time),'%H:%M:%S'), datetime.strptime(max(Time),'%H:%M:%S'), timedelta(seconds = 30))
timegapped = []
for rectime in timegapdt64:
    temp = str(rectime).split('T')[1].split('.')[0]
    timegapped.append(temp)
plt.xticks(timegapped, rotation = "vertical")
plt.show()
print("plotted")
# Save the file
# wb.save("RESULTS.xlsx")
