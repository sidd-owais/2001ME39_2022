import pandas as pd
from openpyxl import load_workbook , Workbook
import math
df = pd.read_excel("input_octant_transition_identify.xlsx")
wb = load_workbook("input_octant_transition_identify.xlsx")
sheet = wb.active
# Creating column for different velocity
sheet.cell(row = 1,column = 5).value = "U_avg"
sheet.cell(row = 1,column = 6).value = "V_avg"
sheet.cell(row = 1,column = 7).value = "W_avg"
sheet.cell(row = 1,column = 8).value = "U-U_avg"
sheet.cell(row = 1,column = 9).value = "V-V_avg"
sheet.cell(row = 1,column = 10).value = "W-W_avg"
sheet.cell(row = 1,column = 11).value = "Octant"
# Average of different coordinate velocity
U_avg = df["U"].mean()
V_avg = df["V"].mean()
W_avg = df["W"].mean()
# Appending means values to sheet
sheet.cell(row = 2,column=5,value = U_avg)
sheet.cell(row = 2,column=6,value = V_avg)
sheet.cell(row = 2,column=7,value = W_avg)
# Calculating difference of absolute and mean
for i in range(2,len(df)+2):
    c2 = sheet.cell(i,2).value
    sheet.cell(i,8,c2-U_avg)
for i in range(2,len(df)+2):
    c2 = sheet.cell(i,3).value
    sheet.cell(i,9,c2-V_avg)
for i in range(2,len(df)+2):
    c2 = sheet.cell(i,4).value
    sheet.cell(i,10,c2-W_avg)