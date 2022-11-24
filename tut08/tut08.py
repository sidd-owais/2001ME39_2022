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
    # Bowling Card

        Current_baller = ""
        count_over_run = 0
        count_over_ball = 0

        for i in range(len(l)):
            x = re.split(r',|\!', l[i])
            info = x[0].split(' ', 1)
            Bats_man_Name = re.split(r'to', info[1])[1]
            Bowler_Name = re.split(r'to', info[1])[0]
            # SIX,FOUR,out,wides,leg byes,run,No ball(later),runout(later))
            Event = x[1]
            if Current_baller != Bowler_Name:
                count_over_run = 0
                count_over_ball = 0
                Current_baller = Bowler_Name
            if Bowler_Name not in Bowling.keys():
                # (over(0),maiden(1),run(2),wicket(3),no_ball(4),wide_ball(5),economy(6))
                Bowling[Bowler_Name] = [0 for i in range(7)]
                Bowling_total_ball[Bowler_Name] = 0
        # Events--------------------------------------------------------------------------------------
            if Event.strip() == "SIX":
                count_over_run += 1
                count_over_ball += 1
                Bowling[Bowler_Name][2] += 6
                Bowling_total_ball[Bowler_Name] += 1
                Total_Run += 6
                Current_run += 6
            if Event.strip() == "FOUR":
                count_over_run += 1
                count_over_ball += 1
                Bowling[Bowler_Name][2] += 4
                Bowling_total_ball[Bowler_Name] += 1
                Total_Run += 4
                Current_run += 4
            if re.split(r' ', Event)[-1].strip() == "run" or re.split(r' ', Event)[-1].strip() == "runs":
                if re.split(r' ', Event)[1].strip() == "1":
                    Bowling[Bowler_Name][2] += 1
                    count_over_run += 1
                    Total_Run += 1
                    Current_run += 1
                if re.split(r' ', Event)[1].strip() == "2":
                    Bowling[Bowler_Name][2] += 2
                    count_over_run += 2
                    Total_Run += 2
                    Current_run += 2
                if re.split(r' ', Event)[1].strip() == "3":
                    Bowling[Bowler_Name][2] += 3
                    count_over_run += 3
                    Total_Run += 3
                    Current_run += 3
                count_over_ball += 1
                Bowling_total_ball[Bowler_Name] += 1
            if re.split(r' ', Event)[-1].strip() == "wide" or re.split(r' ', Event)[-1].strip() == "wides":
                if re.split(r' ', Event)[-1].strip() == "wide":
                    count_over_run += 1
                    Bowling[Bowler_Name][2] += 1
                    Total_Run += 1
                    wides += 1
                    Bowling[Bowler_Name][5] += 1
                    Current_run += 1
                if re.split(r' ', Event)[-1].strip() == "wides":
                    x = int(re.split(r' ', Event)[1].strip())
                    count_over_run += x
                    Bowling[Bowler_Name][2] += x
                    Total_Run += x
                    wides += x
                    Bowling[Bowler_Name][5] += x
                    Current_run += x
            if re.split(r' ', Event)[1].strip() == "out":
                Bowling[Bowler_Name][3] += 1
                count_over_ball += 1
                Bowling_total_ball[Bowler_Name] += 1
                Total_wicket += 1
                wicket_number += 1
                Fall_of_wicket += f'{Current_run}-{wicket_number} ({Bats_man_Name},{info[0]}),'
            if re.split(r' ', Event)[-1].strip() == "byes":
                Events = x[2]
                if Events.strip() == "FOUR":
                    count_over_ball += 1
                    Total_Run += 4
                    Current_run += 4
                    if re.split(r' ', Event)[1].strip() == "leg":
                        leg_byes += 4
                    else:
                        byes += 4
                if re.split(r' ', Events)[-1].strip() == "run" or re.split(r' ', Events)[-1].strip() == "runs":
                    if re.split(r' ', Events)[1].strip() == "1":
                        count_over_ball += 1
                        Total_Run += 1
                        Current_run += 1
                        if re.split(r' ', Event)[1].strip() == "leg":
                            leg_byes += 1
                        else:
                            byes += 1
                    if re.split(r' ', Events)[1].strip() == "2":
                        count_over_ball += 1
                        Total_Run += 2
                        Current_run += 2
                        if re.split(r' ', Event)[1].strip() == "leg":
                            leg_byes += 2
                        else:
                            byes += 2
                    if re.split(r' ', Events)[1].strip() == "3":
                        count_over_ball += 1
                        Total_Run += 3
                        Current_run += 3
                        if re.split(r' ', Event)[1].strip() == "leg":
                            leg_byes += 3
                        else:
                            byes += 3
                Bowling_total_ball[Bowler_Name] += 1

            if count_over_ball == 6 and count_over_run == 0:
                Bowling[Bowler_Name][1] += 1
        # X---------------------------------------------------------X---------------------------X

        # For Calculating Over number and Economy

        for keys in Bowling.keys():
            if keys == "BOWLER":
                continue
            if Bowling_total_ball[keys] % 6 == 0:
                Bowling[keys][0] = Bowling_total_ball[keys]//6
            else:
                After_dot = str(Bowling_total_ball[keys] % 6)
                Before_dot = str(Bowling_total_ball[keys]//6)
                Bowling[keys][0] = Before_dot+"."+After_dot
            Bowling[keys][6] = round(
                (6*Bowling[keys][2])/Bowling_total_ball[keys], 1)

        Extras = leg_byes + wides + byes
        Extra = f'{Extras}:(b:{byes} lb:{leg_byes} w:{wides} nb:{no_ball})'
        Batting["Extras"][1] = Extra
        Total = f'{Total_Run} ({Total_wicket} wkts , {Total_over} Ov)'
        Batting["Totals"][1] = Total

        file_name = "scorecard.txt"
        if count == 0:
            if file_name not in os.listdir():
                with open(file_name, "w") as f:
                    f.close()
    # Batting
        df_1 = pd.DataFrame(Batting).T
        df_1 = df_1.to_markdown(tablefmt='plain')
        list_1 = df_1.split('\n')
        if count == 0:
            with open(file_name, 'w') as fp:
                for item in list_1[1:]:
                    fp.write("%s\n" % item)
                fp.write("\n")
                fp.close()
        else:
            with open(file_name, 'a') as fp:
                for item in list_1[1:]:
                    fp.write("%s\n" % item)
                fp.write("\n")
                fp.close()

        # Fall of wicket
        text_file = open(r"scorecard.txt", "a")
        text_file.write(Fall_of_wicket)
        text_file.write("\n")
        text_file.close()

        # Bowling
        df_2 = pd.DataFrame(Bowling, dtype=str).T
        df_2 = df_2.to_markdown(tablefmt='plain')
        list_2 = df_2.split('\n')
        with open(file_name, 'a') as fp:
            for item in list_2[1:]:
                fp.write("%s\n" % item)
            fp.write("\n")
            fp.close()
        count += 1


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
