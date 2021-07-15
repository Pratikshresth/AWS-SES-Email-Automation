import boto3
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from botocore.exceptions import ClientError


def email_send(file, attachment, subject, body, sender):
    """
    :param file: File containing multiple emails (txt or xlsx)
    :param attachment: Email Attachments (One or Multiple)
    :param subject: Subject of the Email
    :param body: Body of the Email
    :param sender: Sender of the Email
    """

    # For Missing Subject
    if subject == None:
        print("!!! ALERT !!! \nSubject Missing")

    # For Missing body
    elif body == None:
        print("!!! ALERT !!! \nBody Missing")

    # For Missing Sender
    elif sender == None:
        print("!!! ALERT !!! \nPlease Mention the Sender")


    else:
        # FILE EXTENSION VALIDATION
        df = pd.read_excel(file, index_col=0)
        df = df.where(pd.notnull(df), None)
        # Email is the Key and Name is its value
        main_dict = df.to_dict()["Name"]  # Key Value pair

        # SENDER
        SENDER = sender

        # Create a new SES resource and specify a region.
        client = boto3.client('ses', region_name="us-west-2")
        try:
            counter = 0
            # Loop For Handling Multiple Emails
            for email in main_dict:
                msg = MIMEMultipart()
                msg['From'] = SENDER
                msg['To'] = email
                msg['subject'] = subject

                # Body
                body_file = open(body)
                html = body_file.read()
                html = html.replace('{$NAME}',main_dict[email])
                part2 = MIMEText(html, 'html')
                msg.attach(part2)

                # File Attachment
                # Loop For Handling Multiple Attachments
                if attachment==None:
                    pass
                else:
                    for i in attachment:
                        filename = i
                        attach = open(i, "rb")
                        p = MIMEBase('application', 'octet-stream')
                        p.set_payload((attach).read())
                        # encode into base64
                        encoders.encode_base64(p)
                        p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
                        msg.attach(p)

                # Initiate Email Sending
                client.send_raw_email(Source=SENDER,
                                      Destinations=[email],
                                      RawMessage={"Data": msg.as_string()})
                print(f"Email Successfully Sent to {email}")

        except ClientError as e:
            print(e.response['Error']['Message'])


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description='SEND EMAIL')

    parser.add_argument('-e', '--emaillist', metavar="EMAILLIST", type=str,
                        help='File Containing Emails')

    parser.add_argument('-a', '--attach', metavar="ATTACHMENT", action="append",
                        help='Attachment')

    parser.add_argument('-s', '--subject', metavar="SUBJECT", type=str,
                        help='Subject of Message')

    parser.add_argument('-b', '--body', metavar="BODY", type=str,
                        help='Body of Message')

    parser.add_argument('-se', '--sender', metavar="Sender", type=str,
                        help='Sender')

    args = parser.parse_args()

    email_send(args.emaillist, args.attach, args.subject, args.body, args.sender)

