from platform import python_version
from datetime import datetime
start_time = datetime.now()

# Help


def scorecard():
    import os
    import re
    import pandas as pd
    import numpy as np

    file_name_list = []
    for file in os.listdir():
        if file.endswith(".txt"):
            if (file != "teams.txt" and file != "scorecard.txt"):
                file_name_list.append(file.split(".")[0])

    first_team = ""
    Second_team = ""

    for i in file_name_list:
        if (i[-1] == "1"):
            first_team = i
        else:
            Second_team = i
    batting_order = []
    batting_order.append(first_team)
    batting_order.append(Second_team)

    count = 0
    for i in batting_order:
        Batting = {"Batsman": [" ", "R", "B", "4s", "6s", "SR"]}
        Bowling = {"BOWLER": ["O", "M", "R", "W", "NB", "WD", "ECO"]}
        Bowling_total_ball = {}

        Total_Run = 0
        Total_wicket = 0
        leg_byes = 0
        wides = 0
        byes = 0
        no_ball = 0
        Total_over = 0
        Fall_of_wicket = ""
        Current_run = 0
        wicket_number = 0
    # Batting score card

        f = open(i+".txt")
        l = [x for x in f.readlines() if x != "\n"]
        f.close()

        for i in range(len(l)):
            x = re.split(r',|\!', l[i])
            info = x[0].split(' ', 1)
            Bats_man_Name = re.split(r'to', info[1])[1]
            if i == len(l)-1:
                Total_over = info[0]
            Bowler_Name = re.split(r'to', info[1])[0]
            Event = x[1]  # SIX , FOUR , out , wides , leg byes ,run
            if Bats_man_Name not in Batting.keys():
                # (Run(0),(out or not and event),Ball(1),Four(2),Six(3),SR(4))
                Batting[Bats_man_Name] = [0 for i in range(6)]
                Batting[Bats_man_Name][0] = "not out"

            if Event.strip() == "SIX":
                Batting[Bats_man_Name][1] += 6
                Batting[Bats_man_Name][2] += 1
                Batting[Bats_man_Name][4] += 1
                Batting[Bats_man_Name][5] = round(
                    (Batting[Bats_man_Name][1]/Batting[Bats_man_Name][2])*100, 2)

            if Event.strip() == "FOUR":
                Batting[Bats_man_Name][1] += 4
                Batting[Bats_man_Name][2] += 1
                Batting[Bats_man_Name][3] += 1
                Batting[Bats_man_Name][5] = round(
                    (Batting[Bats_man_Name][1]/Batting[Bats_man_Name][2])*100, 2)
            if re.split(r' ', Event)[-1].strip() == "run" or re.split(r' ', Event)[-1].strip() == "runs":
                if re.split(r' ', Event)[1].strip() == "1":
                    Batting[Bats_man_Name][1] += 1
                if re.split(r' ', Event)[1].strip() == "2":
                    Batting[Bats_man_Name][1] += 2
                if re.split(r' ', Event)[1].strip() == "3":
                    Batting[Bats_man_Name][1] += 3
                Batting[Bats_man_Name][2] += 1
                Batting[Bats_man_Name][5] = round(
                    (Batting[Bats_man_Name][1]/Batting[Bats_man_Name][2])*100, 2)
            if re.split(r' ', Event)[1].strip() == "out" or re.split(r' ', Event)[-1].strip() == "byes":
                Batting[Bats_man_Name][2] += 1
                Batting[Bats_man_Name][5] = round(
                    (Batting[Bats_man_Name][1]/Batting[Bats_man_Name][2])*100, 2)
                if re.split(r' ', Event)[1].strip() == "out":
                    if re.split(r' ', Event)[-1].strip() == "Lbw":
                        Batting[Bats_man_Name][0] = "lbw b " + Bowler_Name
                    if re.split(r' ', Event)[-1].strip() == "Bowled":
                        Batting[Bats_man_Name][0] = "b " + Bowler_Name
                    if re.split(r' ', Event)[2].strip() == "Caught":
                        caughter = re.search(
                            '[\w\s]+by\s+([\w\s]+)', Event).group(1)
                        Batting[Bats_man_Name][0] = "c " + \
                            caughter + " b " + Bowler_Name
        Batting["Extras"] = ["" for i in range(6)]
        Batting["Totals"] = ["" for i in range(6)]


# Code
ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")


scorecard()


# This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
