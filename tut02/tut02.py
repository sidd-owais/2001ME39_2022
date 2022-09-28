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