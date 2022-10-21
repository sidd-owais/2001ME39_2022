from platform import python_version
from datetime import datetime
start_time = datetime.now()

# Help https://youtu.be/N6PBd4XdnEw


def octant_range_names(mod=5000):

    octant_name_id_mapping = {"1": "Internal outward interaction", "-1": "External outward interaction", "2": "External Ejection",
                              "-2": "Internal Ejection", "3": "External inward interaction", "-3": "Internal inward interaction", "4": "Internal sweep", "-4": "External sweep"}
    import pandas as pd
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import PatternFill
    import math
    from scipy.stats import rankdata
    df = pd.read_excel("octant_input.xlsx")

    # Using try and except
    try:
        wb = load_workbook("octant_input.xlsx")
    except:
        print("Input file missing")
        exit()
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

    # List for overall count of octant
    octant = [[0 for i in range(8)] for j in range(2)]
    octant[0] = ['+1', '-1', '+2', '-2', '+3', '-3', '+4', '-4']
    # Deciding octant of different coordinate
    for i in range(2, len(df)+2):
        x = sheet.cell(i, 8).value
        y = sheet.cell(i, 9).value
        z = sheet.cell(i, 10).value
        if (x > 0 and y > 0 and z > 0):
            sheet.cell(i, 11, "+1")
            octant[1][0] += 1
        if (x < 0 and y > 0 and z > 0):
            sheet.cell(i, 11, "+2")
            octant[1][2] += 1
        if (x < 0 and y < 0 and z > 0):
            sheet.cell(i, 11, "+3")
            octant[1][4] += 1
        if (x > 0 and y < 0 and z > 0):
            sheet.cell(i, 11, "+4")
            octant[1][6] += 1
        if (x > 0 and y > 0 and z < 0):
            sheet.cell(i, 11, "-1")
            octant[1][1] += 1
        if (x < 0 and y > 0 and z < 0):
            sheet.cell(i, 11, "-2")
            octant[1][3] += 1
        if (x < 0 and y < 0 and z < 0):
            sheet.cell(i, 11, "-3")
            octant[1][5] += 1
        if (x > 0 and y < 0 and z < 0):
            sheet.cell(i, 11, "-4")
            octant[1][7] += 1
    wb.save("octant_output_ranking_excel.xlsx")


ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")


mod = 5000
octant_range_names(mod)


# This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
