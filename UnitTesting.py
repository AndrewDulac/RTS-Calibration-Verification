#  pip install pyinstaller
#  pyinstaller yourprogram.py
#  pyinstaller -F yourprogram.py

import pandas as pd
from easygui import fileopenbox
import openpyxl
from openpyxl import load_workbook, Workbook
from easygui import fileopenbox
from datetime import datetime, timedelta, time
import numpy as np
import csv
import os

wb = load_workbook(os.path.dirname(os.path.realpath(__file__)) + "\\data\\DobleF6150.xlsx")

ws1 = wb.worksheets[0]
testdf = pd.read_excel(os.path.dirname(os.path.realpath(__file__)) + "\\DATAFILE.xlsx")

templatedf = pd.read_excel(os.path.dirname(os.path.realpath(__file__)) + "\\data\\DobleF6150.xlsx", )

print(testdf.head(5))
print(templatedf.head(5))

print("ok")