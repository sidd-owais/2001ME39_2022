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

    # Calculation of total lecture
    dt_1 = set()
    for ind in df.index:
        d = df["Timestamp"][ind].split(" ")[0]
        dt = pd.to_datetime(d, format="%d/%m/%Y")
        if ((dt.day_name() == "Monday") or (dt.day_name() == "Thursday")):
            dt_1.add(dt)
    total_lec_taken = len(dt_1)
    # creating attendance Report
    for ind in df_1.index:
        temp = df[df["Attendance"].str.match(
            str(df_1['Roll No'][ind])+'.*', na=False)].reset_index()
        list_1 = [0 for i in range(7)]
        if len(temp) == 0:
            list_1[0] = df_1["Roll No"][ind]
            list_1[1] = df_1["Name"][ind]
            list_1[2] = total_lec_taken
            list_1[5] = total_lec_taken
        else:
            list_1[0] = temp["Attendance"][0].split(" ")[0]
            list_1[1] = temp["Attendance"][0].split(" ", 1)[1]
            for i in range(len(temp)):
                string_1 = temp["Timestamp"][i].split(" ")
                date = pd.to_datetime(string_1[0], format="%d/%m/%Y")
                if ((date.day_name() != "Monday") and (date.day_name() != "Thursday")):
                    list_1[4] += 1
                else:
                    string_2 = string_1[1].split(":")
                    if (string_2[0] == "14" or ((string_2[0] == "15") and (string_2[1] == "00") and (string_2[2] == "00"))):
                        list_1[3] += 1
                    else:
                        list_1[4] += 1
            list_1[2] = total_lec_taken
            list_1[5] = total_lec_taken - list_1[3]
            list_1[6] = (list_1[3]/list_1[2])
            list_1[6] = round(list_1[6], 2)
        l = pd.DataFrame(list_1).T
        l.to_csv("attendance_report_consolidated.csv",
                 mode='a', index=False, header=False)
    # Creating report for every single student
        df_2.to_csv(list_1[0]+".csv", index=False, header=False)
        l.to_csv(list_1[0]+".csv", mode='a', index=False, header=False)
    # Creating Duplicate attendence report
    h_2 = ["Date", "Roll", "Name", "Total count of attendance on that day"]
    df_3 = pd.DataFrame(h_2).T
    df_3.to_csv("attendance_report_duplicate.csv", index=False, header=False)

    for ind in df_1.index:
        temp = df[df["Attendance"].str.match(
            str(df_1['Roll No'][ind])+'.*', na=False)].reset_index()
        if len(temp) > 0:
            dict_1 = {}
            for i in range(len(temp)):
                string_1 = temp["Timestamp"][i].split(" ")
                date = pd.to_datetime(string_1[0], format="%d/%m/%Y")
                if ((date.day_name() == "Monday") or (date.day_name() == "Thursday")):
                    string_2 = string_1[1].split(":")
                    if (string_2[0] == "14" or ((string_2[0] == "15") and (string_2[1] == "00") and (string_2[2] == "00"))):
                        if (dict_1.get(string_1[0])) == None:
                            dict_1[string_1[0]] = 1
                        else:
                            dict_1[string_1[0]] += 1
            for key in dict_1:
                if (dict_1[key] > 1):
                    list_1 = [0 for i in range(4)]
                    list_1[0] = key
                    list_1[1] = temp["Attendance"][0].split(" ")[0]
                    list_1[2] = temp["Attendance"][0].split(" ", 1)[1]
                    list_1[3] = dict_1[key]
                    l = pd.DataFrame(list_1).T
                    l.to_csv("attendance_report_duplicate.csv",
                             mode='a', index=False, header=False)
    # Emailing
    try:
        from_addr = "md2001ME39@gmail.com"
        # (Important)I am using app password beacuse in gmail there no other option to mail due to security reason.
        pass_of_sender = "bodzwwdevzuxevkr"
        to_addr = "owais786siddiquei@gmail.com"

        # instance of MIMEMultipart
        msg = MIMEMultipart()
        msg['From'] = from_addr
        msg['To'] = to_addr
        msg['Subject'] = "2001ME39 attendance_report_duplicate_file"
        body = "Name - Md 0wais Siddiquei and Roll No. - 2001ME39"
        msg.attach(MIMEText(body, 'plain'))

        # open the file to be sent
        filename = "attendance_report_duplicate.csv"
        attachment = open("attendance_report_duplicate.csv", "rb")

        p = MIMEBase('application', 'octet-stream')
        p.set_payload((attachment).read())
        encoders.encode_base64(p)
        p.add_header('Content-Disposition',
                     "attachment; filename= %s" % filename)
        msg.attach(p)

        # creating SMTP session
        connection = smtplib.SMTP('smtp.gmail.com', 587)
        connection.starttls()
        connection.login(from_addr, pass_of_sender)
        text = msg.as_string()
        connection.sendmail(from_addr, to_addr, text)
        connection.quit()
    except:
        print("Error in sending mail")


ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")


attendance_report()


# This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
