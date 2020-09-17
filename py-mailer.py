# Adapted from this example: https://www.dev2qa.com/python-send-email-to-multiple-contact-in-csv-file-with-personalized-content-example/

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import ssl
import os
import csv
import codecs

from dotenv import load_dotenv
load_dotenv()

def read_contacts_from_csv(csv_file_path):

    csv_rows_list = []

    try:
        # open the csv file.
        file_object = open(csv_file_path, 'r')

        # create a csv.DictReader object with above file object.
        csv_file_dict_reader = csv.DictReader(file_object)
        # get column names list in the csv file.
        column_names = csv_file_dict_reader.fieldnames
        # loop in csv rows.
        for row in csv_file_dict_reader:
            # create a dictionary object to store the row column name value pair.
            tmp_row_dict = {}
            # loop in the row column names list.
            for column_name in column_names:
                # get column value in this row. Convert to string to avoid type convert error.
                column_value = str(row[column_name])
                tmp_row_dict[column_name] = column_value

            csv_rows_list.append(tmp_row_dict)

    except FileNotFoundError:
        print(csv_file_path + " not found.")
    finally:
        print("csv_rows_list = " + csv_rows_list.__repr__())
        return csv_rows_list

# this method mainly focus on send the email.    
def send_email( to_addr ):

    # Email content
    message = MIMEMultipart("alternative")
    message["Subject"] = os.getenv("EMAIL_SUBJECT")
    message["From"] = os.getenv("FROM")
    message["To"] = to_addr

    # Import text and html messages
    text = open("data/message.txt", 'r').read()
    html = codecs.open("data/message.html",'r').read()

    # Turn these into plain/html MIMEText objects
    part1 = MIMEText(text, "plain")
    part2 = MIMEText(html, "html")

    # Add HTML/plain-text parts to MIMEMultipart message
    # The email client will try to render the last part first
    message.attach(part1)
    message.attach(part2)

    # Create a secure SSL context
    context = ssl.create_default_context()

    # Send some mail!
    smtp_server = smtplib.SMTP_SSL(os.getenv("MAIL_SERVER"), os.getenv("PORT"), context=context)
    smtp_server.login(os.getenv("EMAIL_ADDRESS"), os.getenv("EMAIL_PASSWORD"))
    smtp_server.sendmail(os.getenv("EMAIL_ADDRESS"), to_addr, message.as_string())
    smtp_server.quit()


contacts = read_contacts_from_csv('data/members_test.csv')

for row_dict in contacts:

    # parse out required column values from the dictionary object.
    to_addr = ''

    # loop in row keys(csv column names).
    for csv_column_name in row_dict.keys():
        # csv column value.
        column_value = row_dict[csv_column_name]

        # assign column value to related variable.
        if csv_column_name.lower() == 'email address':
            to_addr = column_value

        send_email(to_addr)