import os
import imaplib
import email

import gspread
from dotenv import load_dotenv

load_dotenv()
SERVICE_ACCOUNT_PATH = os.getenv('SERVICE_ACCOUNT_PATH')
GSHEET_ID = os.getenv('GSHEET_ID')
EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS')
APP_PASSWORD = os.getenv('APP_PASSWORD')

gc = gspread.service_account(filename=SERVICE_ACCOUNT_PATH)
spreadsheetID = GSHEET_ID
sh = gc.open_by_key(spreadsheetID)
worksheet = sh.sheet1

# Access imap using SSL
with imaplib.IMAP4_SSL(host="imap.mail.yahoo.com", port=imaplib.IMAP4_SSL_PORT) as imap_ssl:

    # Login to mailbox
    print("Logging into mailbox...")
    resp_code, response = imap_ssl.login(EMAIL_ADDRESS, APP_PASSWORD)
    print(f"Response Code : {resp_code}")
    print(f"Response      : {response[0].decode()}")

    # Set mailbox
    resp_code, mail_count = imap_ssl.select(mailbox="INBOX", readonly=False)

    # Retrieve mail IDs for emails from the no-reply LinkedIn job inbox
    resp_code, mail_ids = imap_ssl.search(None, 'FROM "jobs-noreply@linkedin.com"')

    # Iterate through last n mail_ids and find the ones we're looking for
    for mail_id in mail_ids[0].decode().split()[-50:]:
        resp_code, mail_data = imap_ssl.fetch(mail_id, '(RFC822)') ## Fetch mail data.
        message = email.message_from_bytes(mail_data[0][1]) ## Construct message from mail data
        print(f"Evaluating mail_id {mail_id}")
        if message.get("Subject").startswith("You applied for "):
            emailSubject = message.get("Subject")
            print("Email found - adding to spreadsheet")
            dateApplied = message.get("Date")
            companyAndTitle = emailSubject.split("for ")[1]
            positionAppliedFor,companyAppliedTo = companyAndTitle.split(" at ")
            value_list = [positionAppliedFor,companyAppliedTo,dateApplied]
            worksheet.append_row(value_list,value_input_option='USER_ENTERED')
            # Delete message (moves to Trash folder, per current email settings)
            print(f"Deleting {emailSubject}")
            imap_ssl.store(mail_id,"+FLAGS", "\\Deleted")
    print("End of process.")
    # Manually close the connection and logout
    imap_ssl.close()
    imap_ssl.logout()