import csv
import os
import time
from settings import SENDER_EMAIL, PASSWORD, DISPLAY_NAME, MAIL_COMPOSE, SUBJECT

from smtplib import SMTP
import smtplib
import markdown
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import ssl

def capitalize_name(name):
    # Capitalize the first letter of the name
    return name.capitalize()

def extract_name_from_email(email):
    # Extract the part before the '@'
    local_part = email.split('@')[0]
    # Remove numbers and split by '.'
    name_part = ''.join([char for char in local_part if not char.isdigit()])
    name = name_part.split('.')[0]
    # Capitalize the first letter
    return capitalize_name(name)

def get_msg(csv_file_path, template):
    with open(csv_file_path, "r") as file:
        headers = file.readline().strip().split(",")
        # headers[len(headers) - 1] = headers[len(headers) - 1][:-1]
    # i am opening the csv file two times above and below INTENTIONALLY, changing will cause error
    with open(csv_file_path, "r") as file:
        data = csv.DictReader(file)
        for row in data:
            email = row["EMAIL"]
            if "NAME" in headers:
                name = capitalize_name(row["NAME"])  # Ensure the name is capitalized
            else:
                name = extract_name_from_email(email)
            required_string = template.replace("$NAME", name)
            yield email, required_string
            #uncomment below code if you wish to have NAME column in CSV
            # required_string = template
            # for header in headers:
            #     value = row[header]
            #     required_string = required_string.replace(f"${header}", value)
            # yield row["EMAIL"], required_string


def confirm_attachments():
    file_contents = []
    file_names = []
    try:
        for filename in os.listdir("ATTACH"):

            # entry = input(
            #     f"""TYPE IN 'Y' AND PRESS ENTER IF YOU CONFIRM T0 ATTACH {filename}
            #                         TO SKIP PRESS ENTER: """
            # )
            # confirmed = True if entry == "Y" else False
            # if confirmed:
                file_names.append(filename)
                with open(f"{os.getcwd()}/ATTACH/{filename}", "rb") as f:
                    content = f.read()
                file_contents.append(content)

        return {"names": file_names, "contents": file_contents}
    except FileNotFoundError:
        print("No ATTACH directory found...")
        return {"names": [], "contents": []}


# def send_emails(server: SMTP, template, is_html):

#     attachments = confirm_attachments()
#     sent_count = 0

#     for receiver, message in get_msg("data.csv", template):

#         multipart_msg = MIMEMultipart("alternative")

#         if SUBJECT:
#             multipart_msg["Subject"] = SUBJECT
#         else:
#             multipart_msg["Subject"] = message.splitlines()[0]
#         multipart_msg["From"] = DISPLAY_NAME + f" <{SENDER_EMAIL}>"
#         multipart_msg["To"] = receiver

#         text = message
#         if not is_html:
#             html = markdown.markdown(text)
#         else:
#             html = text

#         part1 = MIMEText(text, "plain")
#         part2 = MIMEText(html, "html")

#         multipart_msg.attach(part1)
#         multipart_msg.attach(part2)

#         if attachments:
#             for content, name in zip(attachments["contents"], attachments["names"]):
#                 attach_part = MIMEBase("application", "octet-stream")
#                 attach_part.set_payload(content)
#                 encoders.encode_base64(attach_part)
#                 attach_part.add_header(
#                     "Content-Disposition", f"attachment; filename={name}"
#                 )
#                 multipart_msg.attach(attach_part)

#         try:
#             server.sendmail(SENDER_EMAIL, receiver, multipart_msg.as_string())
#         except Exception as err:
#             print(f"Problem occurend while sending to {receiver} ")
#             print(err)
#             input("PRESS ENTER TO CONTINUE")
#         else:
#             sent_count += 1

#     print(f"Sent {sent_count} emails")

def send_emails(server: SMTP, template, is_html):
    attachments = confirm_attachments()
    sent_count = 0

    # Split the template into subject and body
    lines = template.splitlines()
    subject = lines[0]  # First line is the subject
    body = "\n".join(lines[1:])  # Rest is the body

    for receiver, message in get_msg("data.csv", body):  # Pass only the body to get_msg
        multipart_msg = MIMEMultipart("alternative")

        # Set the subject
        multipart_msg["Subject"] = subject
        multipart_msg["From"] = DISPLAY_NAME + f" <{SENDER_EMAIL}>"
        multipart_msg["To"] = receiver

        text = message
        if not is_html:
            html = markdown.markdown(text)
        else:
            html = text

        part1 = MIMEText(text, "plain")
        part2 = MIMEText(html, "html")

        multipart_msg.attach(part1)
        multipart_msg.attach(part2)

        if attachments:
            for content, name in zip(attachments["contents"], attachments["names"]):
                attach_part = MIMEBase("application", "octet-stream")
                attach_part.set_payload(content)
                encoders.encode_base64(attach_part)
                attach_part.add_header(
                    "Content-Disposition", f"attachment; filename={name}"
                )
                multipart_msg.attach(attach_part)

        try:
            server.sendmail(SENDER_EMAIL, receiver, multipart_msg.as_string())
        except Exception as err:
            print(f"Problem occurred while sending to {receiver}")
            print(err)
            input("PRESS ENTER TO CONTINUE")
        else:
            sent_count += 1
            print(f"Email sent to {receiver} ({sent_count})")

        time.sleep(20)

    print(f"Sent {sent_count} emails")

# if __name__ == "__main__":
#     host = "smtp.gmail.com"
#     port = 587  # TLS replaced SSL in 1999

#     is_html = MAIL_COMPOSE.endswith("html")

#     with open(MAIL_COMPOSE, "r", encoding="utf-8") as f:
#         template = f.read()

#     context = ssl.create_default_context()

#     server = SMTP(host=host, port=port)
#     server.connect(host=host, port=port)
#     server.ehlo()
#     # server.starttls(context=context)
#     server.starttls()
#     server.ehlo()
#     server.login(SENDER_EMAIL, PASSWORD)
#     print(SENDER_EMAIL, PASSWORD)

#     # with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
#     #     server.login(SENDER_EMAIL, PASSWORD)
#     send_emails(server, template, is_html)
if __name__ == "__main__":
    host = "smtp.gmail.com"  # Gmail SMTP server
    port = 587  # TLS port

    is_html = MAIL_COMPOSE.endswith("html")

    with open(MAIL_COMPOSE, "r", encoding="utf-8") as f:
        template = f.read()

    context = ssl.create_default_context()

    try:
        server = SMTP(host=host, port=port)
        server.connect(host=host, port=port)
        server.ehlo()
        server.starttls(context=context)
        server.ehlo()
        print("Connected to SMTP server.")
        server.login(SENDER_EMAIL, PASSWORD)
        print("Login successful.")
    except Exception as e:
        print(f"Error: {e}")
        exit(1)

    send_emails(server, template, is_html)

# AAHNIK 2023
