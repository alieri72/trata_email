import os
import win32com.client as win32
import tempfile
import email
from email.utils import formatdate
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

output_folder = "Output_PDFs"
html_output_folder = "Output_PDFs"
html_filename = "Email_Bodies.html"

def count_emails_in_inbox():
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # Inbox folder
    num_emails = len(inbox.Items)
    return num_emails

def save_emails_as_pdf():
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # Inbox folder

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    if not os.path.exists(html_output_folder):
        os.makedirs(html_output_folder)

    for email_item in inbox.Items:
        if email_item.Subject == "XX-some-subject--XX":
            subject_parts = email_item.Body.split('|')
            if len(subject_parts) >= 3:
                custom_name = subject_parts[2].strip()
                custom_name = custom_name.replace(":", "")  # Remove colons from the file name
                save_email_as_pdf(email_item, custom_name)
                append_email_body_to_html(email_item)

def save_email_as_pdf(email_item, custom_name):
    subject = email_item.Subject
    sender = email_item.SenderEmailAddress
    recipients = ', '.join([recipient.Address for recipient in email_item.Recipients])
    email_date = formatdate(email_item.ReceivedTime.timestamp(), localtime=True)

    # Parse the email message
    msg = email.message_from_string(email_item.Body)

    # Create a temporary PDF file
    temp_pdf_file = tempfile.NamedTemporaryFile(delete=False)
    temp_pdf_file.close()

    # Generate the PDF
    doc = SimpleDocTemplate(temp_pdf_file.name, pagesize=A4)
    styles = getSampleStyleSheet()

    story = []
    story.append(Paragraph(f"Subject: {subject}", styles['Title']))
    story.append(Paragraph(f"From: {sender}", styles['Normal']))
    story.append(Paragraph(f"To: {recipients}", styles['Normal']))
    story.append(Paragraph(f"Date: {email_date}", styles['Normal']))
    story.append(Spacer(1, 12))

    # Process the email body parts
    for part in msg.walk():
        if part.get_content_type() == 'text/plain':
            body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
            body_lines = body.split("\n| ")
            for line in body_lines:
                story.append(Paragraph(line, styles['Normal']))
                story.append(Spacer(1, 6))  # Line break

    doc.build(story)

    # Rename the temporary PDF to the desired output filename
    output_filename = os.path.join(output_folder, f"{custom_name}.pdf")
    os.rename(temp_pdf_file.name, output_filename)

def append_email_body_to_html(email_item):
    html_file_path = os.path.join(html_output_folder, html_filename)

    with open(html_file_path, 'a', encoding='utf-8') as html_file:
        body = email_item.Body
        lines = body.splitlines()

        for line in lines:
            line = line.replace("/nT|", "<li>")
            if "\n" in line:
                bold_red_text = f'<font color="red"><strong>{line.replace("nT|", "<br/>")}</strong></font>'
                html_file.write(bold_red_text + "<hr/>")
            else:
                html_file.write(line + "<br/>")
        html_file.write("<hr/>")

def main():
    num_emails = count_emails_in_inbox()
    print(f"Number of emails in the inbox: {num_emails}")
    save_emails_as_pdf()
    print("Emails with subject 'XX-some-subject--XX' saved as PDF files in the 'Output_PDFs' folder.")
    print("Email body text appended to the HTML file.")

if __name__ == "__main__":
    main()
