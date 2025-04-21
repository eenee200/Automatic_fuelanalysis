from email.header import decode_header
import os
import re
import email
import imaplib
from datetime import datetime, timedelta
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from smtplib import SMTP
import os

EMAIL_ADDRESS = os.environ.get("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD")
EMAIL_SEND = os.environ.get("EMAIL_SEND")
# --- Helper functions ---

def sanitize_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', '_', filename)

def extract_date_range(filename):
    match = re.search(r'(\d{4}-\d{2}-\d{2} \d{2}_\d{2}_\d{2})_(\d{4}-\d{2}-\d{2} \d{2}_\d{2}_\d{2})', filename)
    if match:
        start = match.group(1).replace("_", ":")
        end = match.group(2).replace("_", ":")
        try:
            start_dt = datetime.strptime(start, "%Y-%m-%d %H:%M:%S")
            end_dt = datetime.strptime(end, "%Y-%m-%d %H:%M:%S")
            if (end_dt - start_dt).days == 7:
                return start_dt, end_dt
        except Exception:
            pass
    return None, None

# --- Gmail Attachment Downloader ---

def save_attachments_from_gmail(save_directory):
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    mail.select('inbox')

    # Get last 7 days in IMAP format
    date_since = (datetime.now() - timedelta(days=7)).strftime('%d-%b-%Y')

    # Search for emails since that date
    status, email_ids = mail.search(None, f'(SINCE {date_since})')
    email_ids = email_ids[0].split()

    if not os.path.exists(save_directory):
        os.makedirs(save_directory)

    attachment_files = []
    print(f"Attachments from emails SINCE {date_since}:\n")

    for e_id in email_ids:
        status, data = mail.fetch(e_id, '(RFC822)')
        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)

        subject = msg.get("Subject", "")
        decoded_subject = decode_header(subject)
        full_subject = ''.join(
            part.decode(enc or "utf-8", errors="ignore") if isinstance(part, bytes) else part
            for part, enc in decoded_subject
        )

        has_attachment = False

        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            file_name = part.get_filename()
            if file_name:
                has_attachment = True
                safe_file_name = sanitize_filename(file_name)
                file_path = os.path.join(save_directory, safe_file_name)
                with open(file_path, 'wb') as f:
                    f.write(part.get_payload(decode=True))
                attachment_files.append(file_path)

        if has_attachment:
            print(f"- {full_subject}")

    mail.logout()
    return attachment_files

# --- Email Sender ---

def send_email_with_attachment(to_email, subject, body, file_path):
    from_email = EMAIL_ADDRESS
    password = EMAIL_PASSWORD

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    with open(file_path, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(file_path)}')
        msg.attach(part)

    with SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(from_email, password)
        server.send_message(msg)

# --- Main Script ---

if __name__ == "__main__":
    save_dir = "./gmail_attachments"
    extracted_files = save_attachments_from_gmail(save_dir)

    # Organize by date range
    gps_pairs = {}
    for f in extracted_files:
        base = os.path.basename(f)
        start, end = extract_date_range(base)
        if start and end:
            key = (start, end)
            if key not in gps_pairs:
                gps_pairs[key] = {}
            if "gps.vechicle" in base:
                gps_pairs[key]['vechicle'] = f
            elif "gps.vech" in base:
                gps_pairs[key]['vech'] = f

    # Filter only complete pairs with 7-day difference
    valid_pairs = [(k, v) for k, v in gps_pairs.items() if 'vechicle' in v and 'vech' in v]
    valid_pairs.sort(key=lambda x: x[0][0], reverse=True)

    if valid_pairs:
        latest_key, latest_files = valid_pairs[0]
        file_path1 = latest_files['vechicle']
        file_path2 = latest_files['vech']

        # Run your analysis function
        from fuel_analysis import main
        excel_file, num_datasets = main(file_path2, file_path1)

        # Force the output file name
        custom_excel_name = "test1.xlsx"
        if excel_file:
            new_excel_path = os.path.join(os.path.dirname(excel_file), custom_excel_name)
            os.rename(excel_file, new_excel_path)

            print(f"Analysis of {num_datasets} datasets exported to {new_excel_path}")
            send_email_with_attachment(
                EMAIL_SEND,
                "Fuel Analysis Report",
                "Please find the attached fuel analysis report.",
                new_excel_path
            )
        else:
            print("Failed to export analysis to Excel")
    else:
        print("No valid 7-day pairs found for analysis.")