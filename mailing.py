import smtplib
from email.message import EmailMessage


def send_mail(SEND_MAIL_FROM, SEND_MAIL_TO, subject, text, files=None, server=["smtp.gmail.com", 465]):
    msg = EmailMessage()

    APP_PW = "ltguzmtffkaaighp"

    msg['subject'] = subject
    msg['from'] = SEND_MAIL_FROM
    msg['to'] = SEND_MAIL_TO

    msg.set_content(text)

    if files:
        with open(files, 'rb') as content_file:
            content = content_file.read()
            msg.add_attachment(content, maintype='application', subtype='pdf', filename=files.split("/")[-1])

    smtp_server = smtplib.SMTP_SSL(server[0], server[1])
    smtp_server.login(SEND_MAIL_TO, APP_PW)
    smtp_server.send_message(msg)
    smtp_server.quit()
