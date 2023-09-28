import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import email.utils
from pickle import load
import pandas as pd
import logging
from datetime import datetime
import traceback


def createMsg(sender, senderName, recipient, Alum):
    # Create message container - the correct MIME type is multipart/alternative.
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "WIT Solapur wishes you a very Happy Birthday!"
    msg['From'] = email.utils.formataddr((senderName, sender))
    msg['To'] = recipient
    msg['Reply-To'] = "alumniconnect@witsolapur.org"

    # Create the body of the message (a plain-text and an HTML version).
    text = """Dear """+Alum+""",\nHere's wishing you tons of happiness on your special day.
    May all your dreams come true and infinite prosperity come your way.\n\nWarm Regards,
    \nWalchand Institute of Technology"""

    html = """\
    <html>
        <body style="background: #008888; text-align: center;">

            <img style="padding: 1em;" src="cid:image1" width="550" height="400">
            <h3 style="color: #FFFFFF; padding-left: 1em; padding-right:1em">Dear """+Alum+""",<br>
                Here's wishing you tons of happiness on your special day.<br>
                May all your dreams come true and infinite prosperity come your way.<br><br>
                Warm Regards,<br>
                Walchand Institute of Technology
            </h3><br>
        </body>
    </html>
    """

    # Record the MIME types of both parts - text/plain and text/html.
    part1 = MIMEText(text, 'plain')
    part2 = MIMEText(html, 'html')

    # Attach parts into message container.
    # According to RFC 2046, the last part of a multipart message, in this case
    # the HTML message, is best and preferred.
    with open('birthday.jpg', 'rb') as fp:
        msgImage = MIMEImage(fp.read())

    # Define the image's ID as referenced above
    msgImage.add_header('Content-ID', '<image1>')
    msg.attach(msgImage)

    msg.attach(part1)
    msg.attach(part2)

    return msg.as_string()


# credentials
sender = ''
senderName = ''
usernameSMTP = ""
passwordSMTP = ""

# get list of recipients
#  load dataframe from cache
with open('db.pkl', 'rb') as f:
    df = load(f)

#  get today's date
today = pd.Timestamp.now().strftime('%m-%d')

# match records
today_df = df.loc[df['Date of Birth'].dt.strftime('%m-%d').eq(today)]

# drop dob column
today_df.drop(['Date of Birth'], axis=1, inplace=True)

# convert into dictionary
recipients = today_df.to_dict('records')

# log recipients
logging.basicConfig(level=logging.INFO, filename='logs\\logs.txt')
logging.info(
    f'{datetime.now().strftime("%d/%m/%Y %H:%M:%S")} : {recipients}')


try:
    # send mails to recipients
    server = smtplib.SMTP('email-smtp.ap-south-1.amazonaws.com', 587)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(usernameSMTP, passwordSMTP)

    for item in recipients:
        server.sendmail(sender, item['Email Address'], createMsg(
            sender, senderName, item['Email Address'], item['First Name']+item['Last Name']))

    server.quit()

except:
    # log exceptions
    logging.exception(traceback.format_exc)
