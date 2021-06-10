# This python script is written to help sending emails with a gmail account to a list of person with customs words
# It's aimed for sending confirmation HTML styles emails
# The code is created following this tutorial : https://realpython.com/python-send-email/#option-1-setting-up-a-gmail-account-for-development
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import csv

csv_contact_file = "contact_test.csv"
html_message = ""
plain_text_message = ""

with open("plain_text_email.txt", "r", encoding="utf-8") as plain_text_file:
    plain_text_message = plain_text_file.read()

with open("html_email.txt", "r", encoding="utf-8") as html_file:
    html_message = html_file.read()

port = 465
sender_email = input("Type mail here : ")

while not sender_email.endswith("@gmail.com"):
    print(f"\"{sender_email}\" is not a valid email, it must be a gmail account")
    sender_email = input("Type again your mail here : ")

password = input("Type password here : ")

# Create a secure SSL context
context = ssl.create_default_context()

with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
    server.login(sender_email, password)
    with open(csv_contact_file, encoding="utf-8") as file:
        reader = csv.reader(file)
        next(reader)  # Skip header row
        counter = 0
        for args in reader:
            counter += 1
            try:
                first_name = args[2]
                name = args[3]
                receiver_email = args[4]

                # Create a MIME Object
                message = MIMEMultipart("alternative")
                message["Subject"] = "Test Mail from Grabuge"
                message["From"] = sender_email
                message["To"] = receiver_email

                html_message_formatted = html_message.format(First_Name=first_name)

                # Turn these into plain/html MIMEText objects
                plain_part = MIMEText(plain_text_message, "plain")
                html_part = MIMEText(html_message_formatted, "html")

                # Add HTML/plain-text parts to MIMEMultipart message
                # The email client will try to render the last part first
                message.attach(plain_part)
                message.attach(html_part)

                print(f"sending email to {first_name}")
                server.sendmail(sender_email, receiver_email, message.as_string())

            except IndexError as error:
                print(error)
                print(f"The row number {counter} didn't provide enough information, skipping it...")
                pass
