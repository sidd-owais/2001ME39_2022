from platform import python_version
from datetime import datetime
start_time = datetime.now()


def octant_longest_subsequence_count_with_range():
    import pandas as pd
    from openpyxl import load_workbook
    import math
    df = pd.read_excel("input_octant_longest_subsequence.xlsx")
    # Using try and except
    try:
        wb = load_workbook("input_octant_longest_subsequence.xlsx")
    except:
        print("Input file missing")

    sheet = wb.active

    # Creating column for different velocity
    sheet.cell(row=1, column=5).value = "U_avg"
    sheet.cell(row=1, column=6).value = "V_avg"
    sheet.cell(row=1, column=7).value = "W_avg"
    sheet.cell(row=1, column=8).value = "U-U_avg"
    sheet.cell(row=1, column=9).value = "V-V_avg"
    sheet.cell(row=1, column=10).value = "W-W_avg"
    sheet.cell(row=1, column=11).value = "Octant"

    # Average of different coordinate velocity
    U_avg = df["U"].mean()
    V_avg = df["V"].mean()
    W_avg = df["W"].mean()

    # Appending means values to sheet
    sheet.cell(row=2, column=5, value=U_avg)
    sheet.cell(row=2, column=6, value=V_avg)
    sheet.cell(row=2, column=7, value=W_avg)

    # Calculating difference of absolute and mean
    for i in range(2, len(df)+2):
        c2 = sheet.cell(i, 2).value
        sheet.cell(i, 8, c2-U_avg)
    for i in range(2, len(df)+2):
        c2 = sheet.cell(i, 3).value
        sheet.cell(i, 9, c2-V_avg)
    for i in range(2, len(df)+2):
        c2 = sheet.cell(i, 4).value
        sheet.cell(i, 10, c2-W_avg)

    # Deciding octant of different coordinate
    for i in range(2, len(df)+2):
        x = sheet.cell(i, 8).value
        y = sheet.cell(i, 9).value
        z = sheet.cell(i, 10).value
        if (x > 0 and y > 0 and z > 0):
            sheet.cell(i, 11, "+1")
        if (x < 0 and y > 0 and z > 0):
            sheet.cell(i, 11, "+2")
        if (x < 0 and y < 0 and z > 0):
            sheet.cell(i, 11, "+3")
        if (x > 0 and y < 0 and z > 0):
            sheet.cell(i, 11, "+4")
        if (x > 0 and y > 0 and z < 0):
            sheet.cell(i, 11, "-1")
        if (x < 0 and y > 0 and z < 0):
            sheet.cell(i, 11, "-2")
        if (x < 0 and y < 0 and z < 0):
            sheet.cell(i, 11, "-3")
        if (x > 0 and y < 0 and z < 0):
            sheet.cell(i, 11, "-4")
    wb.save("output.xlsx")
