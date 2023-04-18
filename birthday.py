import pandas as pd
import datetime
import smtplib  # helps in sending email
import os
os.chdir(r"C:\\Users\\kirthan kumar\\Desktop\\VS code\\Python\\python projects\\birthday wisher")
os.mkdir("testing")
# Enter your authentication details
GMAIL_ID = 'kirthan.cs21@bmsce.ac.in'
GMAIL_PSWD = 'YouWillNotGetIt'


def sendEmail(to, sub, msg):
    print(f"Email to {to} send with subject {sub} and message {msg}")
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PSWD)

    s.sendmail(GMAIL_ID, to, f"Subject: {sub}\n\n{msg}")
    s.quit()


if __name__ == "__main__":
    df = pd.read_excel("data.xlsx")
    # print(df)
    today = datetime.datetime.now().strftime("%d-%m")
    yearNow = datetime.datetime.now().strftime("%y")
    # print(today)
    writeInd = []
    for index, item in df.iterrows():
        # print(index, item["Birthday"])
        bday = item["Birthday"].strftime("%d-%m")
        # print(bday)
        if(today == bday):
            sendEmail(item["e-mail"], "HAPPY BIRTHDAY", item["Dialogue"])
            writeInd.append(index)

    # print(writeInd)
    for i in writeInd:
        yr = df.loc[i, 'year']
        df.loc[i, 'year'] = str(yr) + ',' + str(yearNow)
        # print(df.loc[i, 'year'])

    # print(df)
    df.to_excel('data.xlsx', index=False)
    # print(df)

# search "turn on less secure apps" in google and turn them on if they are off
