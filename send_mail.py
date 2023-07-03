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
KEY_SHEET = ""
sheet = get_datas()

# Đăng nhập mail
server = smtplib.SMTP_SSL(host, port, context=context)
server.login(user, pass_work)


# Content
def send(receiver_mail, received_name):
    msg = MIMEMultipart()
    msg['From'] = user
    msg['Subject'] = "THƯ MỜI THAM DỰ"
    with open('test.html', 'r', encoding='utf-8') as file:
        test_html = file.read()
    test_html = test_html.replace("$NAME", received_name)
    msg.attach(MIMEText(test_html, 'html'))
    server.sendmail(user, receiver_mail, msg.as_string())


i = 1
while True:
    i += 1
    time.sleep(1)
    try:
        receiver_mail = sheet.values().get(spreadsheetId=KEY_SHEET, range=f"Test!B{i}").execute()["values"][0][0]
        received_name = sheet.values().get(spreadsheetId=KEY_SHEET, range=f"Test!C{i}").execute()["values"][0][0]
        try:
            status = sheet.values().get(spreadsheetId=KEY_SHEET, range=f"Test!D{i}").execute()["values"][0][0]
            if status == "Done":
                continue
        except KeyError:
            send(receiver_mail, received_name)
            sheet.values().update(spreadsheetId=KEY_SHEET, range=f"Test!D{i}", valueInputOption="USER_ENTERED",
                                  body={"values": [["Done"]]}).execute()
    except KeyError:
        i -= 1
print("ok")
# try:
#     i += 1
#     receiver_mail = sheet.values().get(spreadsheetId=KEY_SHEET, range=f"Test!B{i}").execute()["values"][0][0]
#     sheet.values().update(spreadsheetId=KEY_SHEET, range=f"Test!C{i}",
#                           valueInputOption="USER_ENTERED", body={'values': [["NotDone"]]}).execute()
#     if sheet.values().get(spreadsheetId=KEY_SHEET, range=f"Test!C{i}").execute()["values"][0][0] == "NotDone":
#         server.sendmail(user, receiver_mail, msg.as_string())
#         sheet.values().update(spreadsheetId=KEY_SHEET, range=f"Test!C{i}",
#                               valueInputOption="USER_ENTERED", body={'values': [["Done"]]}).execute()
# except KeyError:
#     i -= 1
