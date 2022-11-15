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
    # Creating consolidate report
    for ind in df.index:
        d = df["Timestamp"][ind].split(" ")[0]
        dt = pd.to_datetime(d, format="%d/%m/%Y")
        if ((dt.day_name() == "Monday") or (dt.day_name() == "Thursday")):
            dt = dt.date()
            dt = dt.strftime("%d/%m/%Y")
            if dt in s_1:
                continue
            else:
                s_1.add(dt)
                h_1.append(dt)
                dat_1[dt] = count
                count += 1
    h_1.append("Actual Lecture Taken")
    h_1.append("Total Real")
    h_1.append("% Attendance")
    df_2 = pd.DataFrame(h_1).T
    df_2.to_excel("attendance_report_consolidated.xlsx",
                  index=False, header=False)
    l = len(h_1)
    total_lec_taken = len(s_1)
    count_1 = 0
    for ind in df_1.index:
        temp = df[df["Attendance"].str.match(
            str(df_1['Roll No'][ind])+'.*', na=False)].reset_index()
        list_1 = [0 for i in range(l)]
        list_1[2+len(s_1)] = total_lec_taken
        for i in range(2, 2+len(s_1)):
            list_1[i] = 'A'
        if len(temp) == 0:
            list_1[0] = df_1["Roll No"][ind]
            list_1[1] = df_1["Name"][ind]
        else:
            list_1[0] = temp["Attendance"][0].split(" ")[0]
            list_1[1] = temp["Attendance"][0].split(" ", 1)[1]
            for i in range(len(temp)):
                string_1 = temp["Timestamp"][i].split(" ")
                date = pd.to_datetime(string_1[0], format="%d/%m/%Y")
                if ((date.day_name() == "Monday") or (date.day_name() == "Thursday")):
                    string_2 = string_1[1].split(":")
                    if (string_2[0] == "14" or ((string_2[0] == "15") and (string_2[1] == "00") and (string_2[2] == "00"))):
                        date = date.date()
                        date = date.strftime("%d/%m/%Y")
                        if (list_1[dat_1[date]] != "P"):
                            list_1[dat_1[date]] = "P"
                            list_1[3+len(s_1)] += 1
            list_1[4+len(s_1)] = (list_1[3+len(s_1)]/list_1[2+len(s_1)])*100
            list_1[4+len(s_1)] = round(list_1[4+len(s_1)], 2)
        l_1 = pd.DataFrame(list_1).T
        writer = pd.ExcelWriter('attendance_report_consolidated.xlsx',
                                mode='a', if_sheet_exists='overlay')
        l_1.to_excel(writer, startcol=0, startrow=1 +
                     count_1, index=False, header=False)
        writer.close()
        count_1 += 1
    # Creating Report for every single student

    list_0 = ['Date', 'Roll', 'Name',
              'Total Attendance Count', 'Real', 'Duplicate', 'Invalid', 'Absent']
    for ind in df_1.index:
        temp = df[df["Attendance"].str.match(
            str(df_1['Roll No'][ind])+'.*', na=False)].reset_index()
        list_2 = [['', '', '', 0, 0, 0, 0, 0] for j in range(len(s_1)+2)]
        list_2[0] = list_0
        for keys in dat_1:
            list_2[dat_1[keys]][0] = keys
        if len(temp) == 0:
            list_2[1][1] = df_1["Roll No"][ind]
            list_2[1][2] = df_1["Name"][ind]
            for i in range(2, 2+len(s_1)):
                list_2[i][7] = 1
        else:
            list_2[1][1] = temp["Attendance"][0].split(" ")[0]
            list_2[1][2] = temp["Attendance"][0].split(" ", 1)[1]
            for i in range(len(temp)):
                string_1 = temp["Timestamp"][i].split(" ")
                date = pd.to_datetime(string_1[0], format="%d/%m/%Y")
                if ((date.day_name() == "Monday") or (date.day_name() == "Thursday")):
                    string_2 = string_1[1].split(":")
                    date = date.date()
                    date = date.strftime("%d/%m/%Y")
                    if (string_2[0] == "14" or ((string_2[0] == "15") and (string_2[1] == "00") and (string_2[2] == "00"))):
                        if (list_2[dat_1[date]][4]) == 0:
                            list_2[dat_1[date]][4] = 1
                            list_2[dat_1[date]][3] += 1
                        else:
                            list_2[dat_1[date]][5] += 1
                            list_2[dat_1[date]][3] += 1
                    else:
                        list_2[dat_1[date]][6] += 1
                        list_2[dat_1[date]][3] += 1
            for keys in dat_1:
                if list_2[dat_1[keys]][4] == 0:
                    list_2[dat_1[keys]][7] = 1
        df_4 = pd.DataFrame(list_2)
        df_4.to_excel(list_2[1][1]+".xlsx",
                      index=False, header=False)

    # Emailing
    try:
        from_addr = "md2001ME39@gmail.com"  # sender address
        # (Important)I am using app password beacuse in gmail there no other option to mail due to security reason.
        pass_of_sender = "bodzwwdevzuxevkr"
        to_addr = "owais786siddiquei@gmail.com"  # Receiver address

        # instance of MIMEMultipart
        msg = MIMEMultipart()
        msg['From'] = from_addr
        msg['To'] = to_addr
        msg['Subject'] = "2001ME39 attendance_report_duplicate_file"
        body = "Name - Md 0wais Siddiquei and Roll No. - 2001ME39"
        msg.attach(MIMEText(body, 'plain'))

        # open the file to be sent
        filename = "attendance_report_consolidated.xlsx"
        try:
            attachment = open("attendance_report_consolidated.xlsx", "rb")
        except:
            print("file is not available in output folder")

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
