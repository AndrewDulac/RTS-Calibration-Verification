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
v1to3pop = False

expectedVals = {
    "VA": [25,50,75,100,125,150,175,200,225,250],
    "VB": [25,50,75,100,125,150,175,200,225,250],
    "VC": [25,50,75,100,125,150,175,200,225,250],
    "IA": [1,2,3,4,5,6],
    "IB": [1,2,3,4,5,6],
    "IC": [1,2,3,4,5,6],
    "PFDA": [0, 60, 120, 180, -120, -60],
    "PFDB": [0, 60, 120, 180, -120, -60],
    "PFDC": [0, 60, 120, 180, -120, -60],

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

tests = {
    "TYPE": ["START"],
    "V1": [0],
    "V2": [0],
    "V3": [0],
    "V4": [0],
    "V5": [0],
    "V6": [0],    
    "I1": [0],
    "I2": [0],
    "I3": [0],
    "I4": [0],
    "I5": [0],
    "I6": [0],
    "PV1": [0],
    "PV2": [0],
    "PV3": [0],
    "PV4": [0],
    "PV5": [0],
    "PV6": [0],
    "PI1": [0],
    "PI2": [0],
    "PI3": [0],
    "PI4": [0],
    "PI5": [0],
    "PI6": [0],

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
    decimalPercent = tolerance / 200.0
    highRange = expectedVal * (1.0 + decimalPercent)
    lowRange = expectedVal * (1.0 - decimalPercent)
    if(lowRange <= val and val <= highRange):
        return True
    else:
        return False    

def FilterAndIdentify(row):
    for expectedVal in expectedVals["VA"]:
        if( CompareWithTolerance(expectedVal, row["VA"], 1) and
            CompareWithTolerance(expectedVal, row["VB"], 1) and
            CompareWithTolerance(expectedVal, row["VB"], 1) and
            CompareWithTolerance(0, row["IA"], 1) and
            CompareWithTolerance(0, row["IB"], 1) and
            CompareWithTolerance(0, row["IC"], 1) and
            CompareWithTolerance(1, row["PFDA"], .1) and
            CompareWithTolerance(1, row["PFDB"], .1) and
            CompareWithTolerance(1, row["PFDC"], .1)
            ):
            if(v1to3pop):
                tests["TYPE"].append("VOLTAGE4-6")
            else:
                tests["TYPE"].append("VOLTAGE1-3")
            tests["VA"].append(row["VA"])
            tests["VB"].append(row["VB"])
            tests["VC"].append(row["VC"])
            tests["IA"].append(0)
            tests["IB"].append(0)
            tests["IC"].append(0)
            tests["PFDA"].append(0)
            tests["PFDB"].append(0)
            tests["PFDC"].append(0)


    for expectedVal in expectedVals["PFDA"]:
        if(CompareWithTolerance(5, row["VA"], 1) and
            CompareWithTolerance(5, row["VB"], 1) and
            CompareWithTolerance(5, row["VB"], 1) and
            CompareWithTolerance(1, row["IA"], 1) and
            CompareWithTolerance(1, row["IB"], 1) and
            CompareWithTolerance(1, row["IC"], 1) and
            CompareWithTolerance(expectedVal, np.arccos(row["PFDA"]), 1) and
            CompareWithTolerance(expectedVal, np.arccos(row["PFDB"]), 1) and
            CompareWithTolerance(expectedVal, np.arccos(row["PFDC"]), 1)
            ):
            if(v1to3pop):
                tests["TYPE"].append("PHASE4-6at5V")
            else:
                tests["TYPE"].append("PHASE1-3at5V")
            
            tests["VA"].append(0)
            tests["VB"].append(0)
            tests["VC"].append(0)
            tests["IA"].append(0)
            tests["IB"].append(0)
            tests["IC"].append(0)
            tests["PFDA"].append(np.arccos(row["PFDA"]))
            tests["PFDB"].append(np.arccos(row["PFDB"]))
            tests["PFDC"].append(np.arccos(row["PFDC"]))

        elif(CompareWithTolerance(50, row["VA"], 1) and
            CompareWithTolerance(50, row["VB"], 1) and
            CompareWithTolerance(50, row["VB"], 1) and
            CompareWithTolerance(1, row["IA"], 1) and
            CompareWithTolerance(1, row["IB"], 1) and
            CompareWithTolerance(1, row["IC"], 1) and
            CompareWithTolerance(expectedVal, np.arccos(row["PFDA"]), 1) and
            CompareWithTolerance(expectedVal, np.arccos(row["PFDB"]), 1) and
            CompareWithTolerance(expectedVal, np.arccos(row["PFDC"]), 1)
            ):
            tests["TYPE"].append("PHASE50V")
            tests["VA"].append(0)
            tests["VB"].append(0)
            tests["VC"].append(0)
            tests["IA"].append(0)
            tests["IB"].append(0)
            tests["IC"].append(0)
            tests["PFDA"].append(np.arccos(row["PFDA"]))
            tests["PFDB"].append(np.arccos(row["PFDB"]))
            tests["PFDC"].append(np.arccos(row["PFDC"]))

        elif(CompareWithTolerance(100, row["VA"], 1) and
            CompareWithTolerance(100, row["VB"], 1) and
            CompareWithTolerance(100, row["VB"], 1) and
            CompareWithTolerance(1, row["IA"], 1) and
            CompareWithTolerance(1, row["IB"], 1) and
            CompareWithTolerance(1, row["IC"], 1) and
            CompareWithTolerance(expectedVal, np.arccos(row["PFDA"]), 1) and
            CompareWithTolerance(expectedVal, np.arccos(row["PFDB"]), 1) and
            CompareWithTolerance(expectedVal, np.arccos(row["PFDC"]), 1)
            ):
            tests["TYPE"].append("PHASE100V")
            tests["VA"].append(0)
            tests["VB"].append(0)
            tests["VC"].append(0)
            tests["IA"].append(0)
            tests["IB"].append(0)
            tests["IC"].append(0)
            tests["PFDA"].append(np.arccos(row["PFDA"]))
            tests["PFDB"].append(np.arccos(row["PFDB"]))
            tests["PFDC"].append(np.arccos(row["PFDC"]))

        elif(CompareWithTolerance(150, row["VA"], 1) and
            CompareWithTolerance(150, row["VB"], 1) and
            CompareWithTolerance(150, row["VB"], 1) and
            CompareWithTolerance(1, row["IA"], 1) and
            CompareWithTolerance(1, row["IB"], 1) and
            CompareWithTolerance(1, row["IC"], 1) and
            CompareWithTolerance(expectedVal, np.arccos(row["PFDA"]), 1) and
            CompareWithTolerance(expectedVal, np.arccos(row["PFDB"]), 1) and
            CompareWithTolerance(expectedVal, np.arccos(row["PFDC"]), 1)
            ):
            tests["TYPE"].append("PHASE150V")
            tests["VA"].append(0)
            tests["VB"].append(0)
            tests["VC"].append(0)
            tests["IA"].append(0)
            tests["IB"].append(0)
            tests["IC"].append(0)
            tests["PFDA"].append(np.arccos(row["PFDA"]))
            tests["PFDB"].append(np.arccos(row["PFDB"]))
            tests["PFDC"].append(np.arccos(row["PFDC"]))
            if(CompareWithTolerance(-60, np.arccos(row["PFDA"]), 1)):
                v1to3pop = True
            
            
    return row


df['Time'] = df['Time'].map(lambda x: x.replace(" ", ""))
pd.to_datetime(df['Time'], format='%H:%M:%S')
df.set_index('Time', inplace=True)

df.apply(lambda x: TestCheckSingle(x), axis = 1)

print(results)
df.head(5)

