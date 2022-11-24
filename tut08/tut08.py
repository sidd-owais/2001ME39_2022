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
