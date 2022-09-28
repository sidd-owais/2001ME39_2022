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
# Deciding octant of different coordinate
octant = [0,0,0,0,0,0,0,0] #List for overall count of octant
for i in range(2,len(df)+2):
    x = sheet.cell(i,8).value
    y = sheet.cell(i,9).value
    z = sheet.cell(i,10).value
    if(x > 0 and y > 0 and z > 0):
        sheet.cell(i,11,"+1")
        octant[0]+=1
    if(x < 0 and y > 0 and z > 0):
        sheet.cell(i,11,"+2")
        octant[2]+=1
    if(x < 0 and y < 0 and z > 0):
        sheet.cell(i,11,"+3")
        octant[4]+=1
    if(x > 0 and y < 0 and z > 0):
        sheet.cell(i,11,"+4")
        octant[6]+=1
    if(x > 0 and y > 0 and z < 0):
        sheet.cell(i,11,"-1")
        octant[1]+=1
    if(x < 0 and y > 0 and z < 0):
        sheet.cell(i,11,"-2")
        octant[3]+=1
    if(x < 0 and y < 0 and z < 0):
        sheet.cell(i,11,"-3")
        octant[5]+=1
    if(x > 0 and y < 0 and z < 0):
        sheet.cell(i,11,"-4")
        octant[7]+=1
wb.save("output.xlsx")
# Appending octant list 
temp = [[0 for i in range(8)] for i in range(1)]
oct = pd.DataFrame(temp, columns = ['+1','-1','+2','-2','+3','-3','+4','-4'])
oct.iloc[0] = octant
print(oct,"\n")
writer = pd.ExcelWriter('output.xlsx', mode = 'a', if_sheet_exists = 'overlay')
oct.to_excel(writer, startcol = 11 , startrow = 0, index=False)