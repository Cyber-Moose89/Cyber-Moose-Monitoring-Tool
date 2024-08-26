import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def send_email(subject, body, config):
    smtp_server = config['EMAIL']['SMTP_Server']
    smtp_port = config['EMAIL']['SMTP_Port']
    email_from = config['EMAIL']['From']
    email_password = config['EMAIL']['Password']
    email_to_list = config['EMAIL']['To'].split(',')

    msg = MIMEMultipart("alternative")
    msg['From'] = email_from
    msg['To'] = ', '.join(email_to_list)
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'html'))

    try:
        server = smtplib.SMTP(smtp_server, int(smtp_port))
        server.starttls()
        server.login(email_from, email_password)
        server.sendmail(email_from, email_to_list, msg.as_string())
        server.quit()
    except Exception as e:
        print(f"Failed to send email: {e}")
