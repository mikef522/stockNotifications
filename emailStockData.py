from datetime import datetime
from pytz import timezone

import smtplib
from pathlib import Path

from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate

COMMASPACE = ', '
from email import encoders





def send_mail(send_from, send_to, subject, message, files=[],
              server="localhost", port=587, username='', password='',
              use_tls=True):
    """Compose and send email with provided info and attachments.

    Args:
        send_from (str): from name
        send_to (list[str]): to name(s)
        subject (str): message title
        message (str): message body
        files (list[str]): list of file paths to be attached to email
        server (str): mail server host name
        port (int): port number
        username (str): server auth username
        password (str): server auth password
        use_tls (bool): use TLS mode
    """
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(message))

    for path in files:
        part = MIMEBase('application', "octet-stream")
        with open(path, 'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        'attachment; filename="{}"'.format(Path(path).name))
        msg.attach(part)

    smtp = smtplib.SMTP(server, port)
    if use_tls:
        smtp.starttls()
    smtp.login(username, password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.quit()


def emailStockData(stockDataFilesToEmail,emailRecipients, datesForStockDataBeingEmailed, userOptions):
    '''
    stockDataFilesToEmail (list[str]): list of filenames
    '''


    #compose the email message
    tz = timezone('US/Eastern')
    currentTimeAndDate = datetime.now(tz).strftime('%I:%M%p EST on %m-%d-%Y')

    header = 'This is an automated end of day stock report for the following dates: \n'
    for date in datesForStockDataBeingEmailed:
        header = header + date.strftime('%m-%d-%Y') + '\n'

    footer = 'Data generated at ' + currentTimeAndDate + ' using data from Yahoo! Finance. Thank you for choosing Mike Ferguson automated stock tracking and reports. Good day'

    emailBody = header + '\nPlease see the attached files for the stock data for each day listed above. The file name tells you the day that that data is for.\n\n\n\n' + footer





    #send the email (also compose the subject line)
    send_from = userOptions['notificationEmailAddress']
    send_to = emailRecipients

    subject = ''
    #check if there is only data for one date being emailed. Adjust the subject title accordingly
    subjectString = 'Automated Stock Notification v10'
    #subjectString = 'Most shorted stocks'
    if (datesForStockDataBeingEmailed[0] == datesForStockDataBeingEmailed[-1]):
        subject = datesForStockDataBeingEmailed[0].strftime('%m-%d-%Y') + ' ' + subjectString
    else:
        subject = datesForStockDataBeingEmailed[0].strftime('%m-%d-%Y') + ' to ' + datesForStockDataBeingEmailed[-1].strftime('%m-%d-%Y') + ' ' + subjectString

    message = emailBody

    files = []
    for file in stockDataFilesToEmail:
        files.append(Path.cwd().joinpath(file))

    server = 'smtp.gmail.com'
    username = userOptions['notificationEmailAddress']
    password = userOptions['notificationEmailPassword']

    send_mail(send_from,send_to,subject,message,files,server=server,username=username,password=password)