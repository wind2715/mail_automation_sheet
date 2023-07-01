from get_api import get_datas
import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time

# Lấy user password, host, port
user = ""
pass_work = ""
host = "smtp.gmail.com"
port = 465
context = ssl.create_default_context()

# Lấy data từ sheet
KEY_SHEET = "1uG-lkvhblM244hxKixw6KAkPQGwYiY2j360idHpfsgo"
RANGE_SHEET = "Test!A1:A4"
sheet = get_datas()

# Đăng nhập mail
server = smtplib.SMTP_SSL(host, port, context=context)
server.login(user, pass_work)

# Content
msg = MIMEMultipart()
msg['From'] = user
msg['Subject'] = "CHÓ TÙNG"
with open('test.html', 'r', encoding='utf-8') as file:
    test_html = file.read()
msg.attach(MIMEText(test_html, 'html'))
i = 1
while True:
    time.sleep(1)
    try:
        i += 1
        receiver_mail = sheet.values().get(spreadsheetId=KEY_SHEET, range=f"Test!B{i}").execute()["values"][0][0]
        sheet.values().update(spreadsheetId=KEY_SHEET, range=f"Test!C{i}",
                              valueInputOption="USER_ENTERED", body={'values': [["NotDone"]]}).execute()
        if sheet.values().get(spreadsheetId=KEY_SHEET, range=f"Test!C{i}").execute()["values"][0][0] == "NotDone":
            server.sendmail(user, receiver_mail, msg.as_string())
            sheet.values().update(spreadsheetId=KEY_SHEET, range=f"Test!C{i}",
                                  valueInputOption="USER_ENTERED", body={'values': [["Done"]]}).execute()
    except KeyError:
        i -= 1
print("ok")
