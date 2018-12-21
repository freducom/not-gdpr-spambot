import smtplib
import os
import csv
import time
import random
from dotenv import load_dotenv
from email.mime.text import MIMEText


def establish_connection(from_, pass_, server_, port_):
    try:
        print("Setting up connection to your SMTP Server...")
        smtpserver = smtplib.SMTP(server_, port_)
        print("Connecting to SMTP server...")
        smtpserver.ehlo()
        print("Starting private session...")
        smtpserver.starttls()
        smtpserver.ehlo()
        print("Logging into SMTP server...")
        smtpserver.login(from_, pass_)
        return smtpserver
    except:
        print("Error establishing connection to SMTP server")


def send_email(smtpserver, from_address, to_address, bcc_address, subject, body):
    msg = MIMEText(body)
    msg['From'] = from_address
    msg['To'] = to_address
    msg['Subject'] = subject
    msg['Bcc'] = bcc_address

    try:
        smtpserver.sendmail(from_address, to_address, msg.as_string())
        print("Email sent to " + to_address)
    except:
        print("Error sending email to " + to_address)


def main():
    # Read env variables
    load_dotenv()

    login = os.getenv("SENDER")
    password = os.getenv("EMAILPASSWORD")
    smtp_server = os.getenv("SMTPSERVER")
    smtp_port = os.getenv("SMTPPORT")
    bcc_address = os.getenv("BCCADDRESS")

    # Connect to server
    smtp_connection = establish_connection(login, password, smtp_server, smtp_port)

    # Define message & subject
    sender = login
    subject = os.getenv("EMAILSUBJECT")
    with open('email_content.txt', 'r') as myfile:
        data = myfile.read()

    # Loop through all recipients and send email
    first = True
    with open("recipients.csv") as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            if first:  # only wait between sends, not before first or after last
                first = False
            else:
                rndwait = 2 + random.randint(0, 9)  # No higher science behind this delay, I just liked these numbers.
                print("Waiting", rndwait, "seconds...")
                time.sleep(rndwait)
            to = row['email']
            body = data.replace('%fname', row['first_name'].capitalize())
            send_email(smtp_connection, sender, to, bcc_address, subject, body)

    # Finally close the SMTP connection
    smtp_connection.close()


if __name__ == '__main__':
    main()
