import imaplib
import email
from email.header import decode_header
import csv

# Gmail IMAP server details
IMAP_SERVER = "imap.gmail.com"
IMAP_PORT = 993

# Login credentials
EMAIL_ACCOUNT = "rdaha526@rku.ac.in"
PASSWORD = "myck zjdi hxwh zrhb"

try:
    # Connect to the server
    print("Connecting to server...")
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)

    # Login to your email account
    print("Logging in...")
    mail.login(EMAIL_ACCOUNT, PASSWORD)
    print("Logged in successfully.")

    # Select the mailbox you want to use (in this case, the inbox)
    mail.select("inbox")

    # Search for all emails
    status, messages = mail.search(None, "ALL")
    email_ids = messages[0].split()
    print(f"Number of emails: {len(email_ids)}")

    # Prepare the CSV file to write the extracted data
    with open("emails.csv", "w", newline="") as csvfile:
        fieldnames = ["Sender", "Subject"]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()

        # Loop through each email
        for email_id in email_ids:
            status, msg_data = mail.fetch(email_id, "(RFC822)")
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding if encoding else "utf-8")
                    from_ = msg.get("From")
                    if from_:
                        sender = from_.split("<")[-1].strip(">")
                    print(f"Writing: {sender} - {subject}")
                    writer.writerow({"Sender": sender, "Subject": subject})

    print("Emails extracted successfully.")

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    # Logout and close the connection
    mail.logout()
