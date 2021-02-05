import argparse
import email
import getpass
import imaplib
import json
import os
import re
import smtplib
import ssl
import webbrowser


def clean(text):
    # Clean text for creating a folder.
    return re.sub("[^0-9a-zA-Z]+", "_", text)


def get_email():
    user_email = input("Email: ")
    # Validate the email entered.
    while not re.match(r"[^@]+@[^@]+\.[^@]+", user_email):
        print("Error: Email provided is invalid.")
        user_email = input("Email: ")
    return user_email


def get_credentials():
    user_email = get_email()
    # Mask the password entered.
    password = getpass.getpass(prompt="Password: ")
    return user_email, password


def new_email():
    email, password = get_credentials()
    smtp_server = ""
    with open("settings.json", "r") as settings_file:
        settings = json.load(settings_file)
    try:
        smtp_server = settings[email[email.find("@") + 1 : email.rfind(".")]]["smtp_server"]
    except KeyError:
        print("Error: Email provider not supported.")
        return
    recipient = get_email()
    subject = input("Subject: ")
    content = input("Message: ")
    message = "Subject:" + subject + "\n\n" + content
    # Set up SMTP server.
    port = 587 
    context = ssl.create_default_context()
    # Try to log in and send the email.
    with smtplib.SMTP(smtp_server,port) as server:
        try:
            server.ehlo()
            server.starttls(context=context)
            server.ehlo()
            server.login(email, password)
            server.sendmail(email, recipient, message)
        except Exception as exception:
            # Print any error messages to stdout.
            print(exception)


def email_multipart(msg, email_subject):
    # Invoked when the email is multipart.
    # Iterate over email parts.
    for part in msg.walk():
        # Extract content type of the email.
        content_type = part.get_content_type()
        content_disposition = str(part.get("Content-Disposition"))
        try:
            # Get the email body.
            body = part.get_payload(decode=True).decode()
        except:
            pass
        if content_type == "text/plain" and "attachment" not in content_disposition:
            # Print text/plain emails and skip attachments.
            return body
        elif "attachment" in content_disposition:
            # Download attachment.
            filename = part.get_filename()
            if filename:
                folder_name = clean(email_subject)
                if not os.path.isdir(folder_name):
                    # Make a folder for this email (named after the subject).
                    os.mkdir(folder_name)
                filepath = os.path.join(folder_name, filename)
                # Download attachment and save it.
                open(filepath, "wb").write(part.get_payload(decode=True))


def open_html_email(body, email_subject):
    # If it is HTML, create a new HTML file and open it in browser if the user wishes to do so.
    folder_name = clean(email_subject)
    if not os.path.isdir(folder_name):
        # Make a folder for this email (named after the subject).
        os.mkdir(folder_name)
    filename = "index.html"
    filepath = os.path.join(folder_name, filename)
    # Write the file.
    open(filepath, "w").write(body)
    # Open in the default browser.
    webbrowser.open(filepath)


def view_email(data):
    for response_part in data:
        arr = response_part[0]
        if isinstance(arr, tuple):
            msg = email.message_from_string(str(arr[1],"utf-8"))
            email_subject = msg["subject"] if msg["subject"] else "NO_SUBJECT"
            email_from = msg["from"]
            print("From : " + email_from + "\n")
            print("Subject : " + email_subject + "\n")
            # If the message is multipart.
            if msg.is_multipart():
                print(email_multipart(msg, email_subject))
            else:
                # Extract content type of email.
                content_type = msg.get_content_type()
                # Get the email body.
                body = msg.get_payload(decode=True).decode()
                if content_type == "text/plain":
                    # Print only text email parts.
                    print(body)
            content_type = msg.get_content_type()
            if content_type == "text/html":
                answer = input("The following email contains HTML. Would you like to open it (y/n)?: ")
                while answer != "y" and answer != "n":
                    print("Error: Invalid answer.")
                    answer = input("The following email contains HTML. Would you like to download and open it (y/n)?: ")
                if answer == "y":
                    open_html_email(body, email_subject)
            print("=" * 100)


def view_inbox():
    # Repeat prompting for credentials until correct ones are provided.
    logged_in = False
    while not logged_in:
        user_email, password = get_credentials()
        # Try to load the IMAP server name from settings.json file.
        imap_server_name = ""
        with open("settings.json", "r") as settings_file:
            settings = json.load(settings_file)
        try:
            imap_server_name = settings[user_email[user_email.find("@") + 1 : user_email.rfind(".")]]["imap_server"]
        except KeyError:
            print("Error: Email provider not supported.")
            continue
        # Initialise the IMAP server.
        imap_server =imaplib.IMAP4_SSL(imap_server_name)
        try:
            imap_server.login(user_email, password)
        except:
            print("Error: Incorrect credentials.")
        else:
            logged_in = True
    messages = imap_server.select("INBOX")[1]
    while True:
        try:
            emails_to_load = int(input("Select a number of emails to load: "))
        except ValueError:
            print("Error: Invalid value.")
            continue
        else:
            break
    messages = int(messages[0])
    for i in range(messages, messages - emails_to_load, -1):
        data = imap_server.fetch(str(i), "(RFC822)")
        view_email(data)
    imap_server.close()
    imap_server.logout()


def add_server():
    # Add a server to the settings.
    server_type = input("Type of server (SMTP / IMAP): ")
    while server_type.casefold() != "smtp" and server_type.casefold() != "imap":
        print("Error: Type must be SMTP or IMAP.")
        server_type = input("Type of server (SMTP / IMAP): ")
    
    name = input("Email provider: ")
    while not name:
        print("Error: Email provider can't be empty.")
        name = input("Email provider: ")
    with open("settings.json", "r+") as settings_file:
        settings = json.load(settings_file) 
    if name in settings and server_type.casefold() + "_server" in settings[name]:
        print("Error: This provider already exists. You can try editing it.")
    else:
        server = input("Email server: ")
        if name in settings:
            settings[name][server_type.casefold() + "_server"] = server
            settings_file.seek(0)
            json.dump(settings, settings_file, indent=4)
        else:
            settings[name] = {
                server_type.casefold() + "_server": server
            }
            settings_file.seek(0)
            json.dump(settings, settings_file, indent=4)


def edit_server():
    # Edit server in the settings.
    server_type = input("Type of server (SMTP / IMAP): ")
    while type.casefold() != "smtp" and server_type.casefold() != "imap":
        print("Error: Type must be SMTP or IMAP.")
        server_type = input("Type of server (SMTP / IMAP): ")

    name = input("Email provider: ")
    while not name:
        print("Error: Email provider can't be empty.")
        name = input("Email provider: ")

    with open("settings.json", "r+") as settings_file:
        settings = json.load(settings_file)
    if name not in settings or server_type.casefold() + "_server" not in settings[name]:
        print("Error: This provider doesn't exist. You can try adding it.")
    else:
        server = input("Email server: ")
        while not server:
            print("Error: Email server can't be empty. You can try removing it.")
            server = input("Email server: ")
        settings[name][server_type.casefold() + "_server"] = server
        settings_file.seek(0)
        json.dump(settings, settings_file, indent=4)


def remove_server():
    name = input("Email provider: ")
    while not name:
        print("Error: Email provider can't be empty.")
        name = input("Email provider: ")
    with open("settings.json", "r") as settings_file:
        settings = json.load(settings_file)
    while name not in settings:
        print("Error: This provider doesn't exist.")
        name = input("Email provider: ")
    else:
        answer = input("Are you sure (y/n)?: ")
        while answer != "y" and answer != "n":
            print("Error: Invalid answer.")
            answer = input("Are you sure (y/n)?: ")
        if answer == "y":
            del settings[name]
            with open("settings.json", "w") as settings_file:
                json.dump(settings, settings_file, indent=4)


if __name__ == "__main__":
    arg_parser = argparse.ArgumentParser()
    arg_parser.add_argument("action", help="view/send/add/edit/remove")
    args = arg_parser.parse_args()
    # Dictionary that matches the value of action argument to a function that is supposed to be called.
    try:
        action_to_function = {"view": view_inbox, "send": new_email, "add": add_server, "edit": edit_server, "remove": remove_server}
        action_to_function[args.action]()
    except KeyError:
        arg_parser.print_help()


