from platform import python_version
from datetime import datetime
start_time = datetime.now()


def attendance_report():
    import pandas as pd
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders

    try:
        df = pd.read_csv("input_attendance.csv")
        df_1 = pd.read_csv("input_registered_students.csv")
    except:
        print("Input file is wrong")
        exit()

    h_1 = ['Roll', 'Name', 'total_lecture_taken',
           'attendance_count_actual', 'attendance_count_fake', 'attendance_count_absent', 'Percentage (attendance_count_actual/total_lecture_taken) 2 digit decimal']
    df_2 = pd.DataFrame(h_1).T
    df_2.to_csv("attendance_report_consolidated.csv",
                index=False, header=False)


ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")


attendance_report()


# This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
