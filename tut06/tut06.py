from platform import python_version
from datetime import datetime
start_time = datetime.now()


def attendance_report():
    import pandas as pd
    import datetime
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    import os

    try:
        df = pd.read_csv("input_attendance.csv")
        df_1 = pd.read_csv("input_registered_students.csv")
    except:
        print("Input file is wrong")
        exit()
    h_1 = ['Roll', 'Name']
    s_1 = set()
    dat_1 = {}
    count = 2

    # Creating output folder if it don't exist and saving output in output folder
    current_directory = os.getcwd()
    final_directory = os.path.join(current_directory, r'output')
    if not os.path.exists(final_directory):
        os.makedirs(final_directory)
    os.chdir(final_directory)


ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")


attendance_report()


# This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
