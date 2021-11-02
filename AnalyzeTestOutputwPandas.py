# Install c++ runtime environment if not present -- latest version here: https://docs.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist?view=msvc-160
# pip install numpy
# pip install pandas
# pip install easygui
# pip install openpyxl


#  pip install pyinstaller
#  pyinstaller yourprogram.py
#  pyinstaller -F yourprogram.py
import pandas as pd
from easygui import fileopenbox
from openpyxl import load_workbook, Workbook
from easygui import fileopenbox
from datetime import datetime, timedelta, time
import numpy as np

def CellIdForValueInColumn(val, col):
    rowId = 1
    cellId = col + str(rowId)
    focusedCell = ws1[cellId]
    while(focusedCell.value != None and val not in str(focusedCell.value)):
        rowId = rowId + 1
        cellId = col+ str(rowId)
        focusedCell = ws1[cellId]
    return col,rowId

#fn = fileopenbox()
fn = 'C:\\Users\\andre\\source\\repos\\ECE591\\TESTFILEV1.csv'
wb = load_workbook(fn)
ws1 = wb.worksheets[0]

rowId = 1
cellId = 'A'+ str(rowId)
focusedCell = ws1[cellId]

col,row = CellIdForValueInColumn("Record", 'A')


df = pd.read_excel(fn, skiprows=(row-1))
df.columns = df.columns.str.replace("\"","")
df.columns = df.columns.str.replace(" ","")
cols = df.columns

expectedVals = {
    "VA": [25,50,75,100,125,150,175,200,225,250],
    "VB": [25,50,75,100,125,150,175,200,225,250],
    "VC": [25,50,75,100,125,150,175,200,225,250],
    "IA": [1,2,3,4,5],
    "IB": [1,2,3,4,5],
    "IC": [1,2,3,4,5],
}

doubleExpectedVals = {
    "V1": [25,50,75,100,125,150],
    "V2": [25,50,75,100,125,150],
    "V3": [25,50,75,100,125,150],
    "V4": [25,50,75,100,125,150],
    "V5": [25,50,75,100,125,150],
    "V6": [25,50,75,100,125,150],
    "V1P": [0, 60, 120, 180, -120, -60],
    "V2P": [0, 60, 120, 180, -120, -60],
    "V3P": [0, 60, 120, 180, -120, -60],
    "V4P": [0, 60, 120, 180, -120, -60],
    "V5P": [0, 60, 120, 180, -120, -60],
    "V6P": [0, 60, 120, 180, -120, -60],

# delta phase = 30deg
    "I1": [1,2,3,4,5,6],
    "I2": [1,2,3,4,5,6],
    "I3": [1,2,3,4,5,6],
    "I4": [1,2,3,4,5,6],
    "I5": [1,2,3,4,5,6],
    "I6": [1,2,3,4,5,6],
    "I1P": [0, 60, 120, 180, -120, -60],
    "I2P": [0, 60, 120, 180, -120, -60],
    "I3P": [0, 60, 120, 180, -120, -60],
    "I4P": [0, 60, 120, 180, -120, -60],
    "I5P": [0, 60, 120, 180, -120, -60],
    "I6P": [0, 60, 120, 180, -120, -60],
}

doubleTolerances = {
    "V1": 0.3, 
    "V2": 0.3,
    "V3": 0.3,
    "V4": 0.3,
    "V5": 0.3,
    "V6": 0.3,
    "I1": 0.3,
    "I2": 0.3,
    "I3": 0.3,
    "I4": 0.3,
    "I5": 0.3,
    "I6": 0.3,
}

results = {
    "VA_3SEC": [None] * len(expectedVals["VA_3SEC"]),
    "VB_3SEC": [None] * len(expectedVals["VB_3SEC"]),
    "VC_3SEC": [None] * len(expectedVals["VC_3SEC"]),
    "IA_3SEC": [None] * len(expectedVals["IA_3SEC"]),
    "IB_3SEC": [None] * len(expectedVals["IB_3SEC"]),
    "IC_3SEC": [None] * len(expectedVals["IC_3SEC"]),
}

dobleResults = {
    "V1": [None] * len(doubleExpectedVals["V1"]),
    "V2": [None] * len(doubleExpectedVals["V2"]),
    "V3": [None] * len(doubleExpectedVals["V3"]),
    "V4": [None] * len(doubleExpectedVals["V4"]),
    "V5": [None] * len(doubleExpectedVals["V5"]),
    "V6": [None] * len(doubleExpectedVals["V6"]),
    "V1P": [None] * len(doubleExpectedVals["V1P"]),
    "V2P": [None] * len(doubleExpectedVals["V2P"]),
    "V3P": [None] * len(doubleExpectedVals["V3P"]),
    "V4P": [None] * len(doubleExpectedVals["V4P"]),
    "V5P": [None] * len(doubleExpectedVals["V5P"]),
    "V6P": [None] * len(doubleExpectedVals["V6P"]),

    "I1": [None] * len(doubleExpectedVals["I1"]),
    "I2": [None] * len(doubleExpectedVals["I2"]),
    "I3": [None] * len(doubleExpectedVals["I3"]),
    "I4": [None] * len(doubleExpectedVals["I4"]),
    "I5": [None] * len(doubleExpectedVals["I5"]),
    "I6": [None] * len(doubleExpectedVals["I6"]),
    "I1P": [None] * len(doubleExpectedVals["I1P"]),
    "I2P": [None] * len(doubleExpectedVals["I2P"]),
    "I3P": [None] * len(doubleExpectedVals["I3P"]),
    "I4P": [None] * len(doubleExpectedVals["I4P"]),
    "I5P": [None] * len(doubleExpectedVals["I5P"]),
    "I6P": [None] * len(doubleExpectedVals["I6P"]),
}

def CompareWithTolerance(expectedVal, val, tolerance):
    if(expectedVal )
    decimalPercent = tolerance / 200.0
    highRange = expectedVal * (1.0 + decimalPercent)
    lowRange = expectedVal * (1.0 - decimalPercent)
    if(lowRange <= val and val <= highRange):
        return True
    else:
        return False    

def FilterAndIdentify(row):
    for column in doubleExpectedVals:
        if( CompareWithTolerance(doubleExpectedVals["VA"], row["VA"], 1) and
            CompareWithTolerance(doubleExpectedVals["VB"], row["VB"], 1) and
            CompareWithTolerance(doubleExpectedVals["VC"], row["VC"], 1)
    return row


df['Time'] = df['Time'].map(lambda x: x.replace(" ", ""))
pd.to_datetime(df['Time'], format='%H:%M:%S')
df.set_index('Time', inplace=True)

df.apply(lambda x: TestCheckSingle(x), axis = 1)

print(results)
df.head(5)

