import smtplib
import time
import win32com.client
import base64
from cryptor import *

smtp_server = 'smtp.163.com'
smtp_port = 25
smtp_acct = 'q941307914@163.com'
smtp_password = 'PSAUHWGDWMFKVHKZ'
tgt_accts = ['hzhnan@qq.com']


def plain_email(subject,contents):
    encodecontents = encrypt(contents)
    message = f'Subject:{subject}\nFrom {smtp_acct}\n'
    message += f'To:{tgt_accts}\n\n{encodecontents}'
    server = smtplib.SMTP(smtp_server,smtp_port)
    server.starttls()
    server.login(smtp_acct,smtp_password)
    # server.set_debuglevel(1)
    server.sendmail(smtp_acct,tgt_accts,message)
    time.sleep(1)
    server.quit()


def outlook(subject,contents):
    outlook = win32com.client.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.DeleteAfterSubmit = True
    message.Subject = subject
    message.Body = contents.decode()
    message.To = tgt_accts[0]
    message.Send()

if __name__ == '__main__':
    plain_email('hello','this is the test')