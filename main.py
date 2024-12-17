import win32com.client as client
import time
import os
from datetime import datetime
import hashlib
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.colors import red
import io
import shutil


def create_watermark(watermark_text):
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)

    can.setFont("Helvetica", 100)
    can.setFillColor(red)
    can.setFillAlpha(0.3)

    can.translate(100, 200)
    can.rotate(45)

    can.drawString(0, 0, watermark_text)
    can.save()

    packet.seek(0)
    return packet


def add_watermark(input_pdf_path, output_pdf_path, watermark_text):
    watermark_pdf = PdfReader(create_watermark(watermark_text))
    watermark_page = watermark_pdf.pages[0]

    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()

    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        page.merge_page(watermark_page)
        writer.add_page(page)

    with open(output_pdf_path, "wb") as output_pdf:
        writer.write(output_pdf)


def generate_email_id(message):
    unique_string = f"{message.EntryID}_{message.ReceivedTime.strftime('%Y%m%d%H%M%S')}"
    email_hash = hashlib.md5(unique_string.encode()).hexdigest()[:8]
    return email_hash


def send_processed_files(output_folder, target_email, unique_email_id):
    try:
        # Get the MAPI namespace first
        outlook = client.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace('MAPI')

        # Create a new mail item
        mail = outlook.CreateItem(0)

        # Get the default profile's email address
        accounts = namespace.Accounts
        default_account = None
        for account in accounts:
            print(f"Found account: {account.DisplayName}")
            default_account = account
            break

        if default_account is None:
            raise Exception("No email account found in Outlook")

        # Configure the email
        mail.SendUsingAccount = default_account
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, default_account))  # Force the account to be used

        mail.To = target_email
        mail.Subject = f"Processed files for email {unique_email_id}"
        mail.Body = f"Please find attached the processed files from folder: {output_folder}"

        # Add attachments
        for filename in os.listdir(output_folder):
            file_path = os.path.abspath(os.path.join(output_folder, filename))
            if os.path.isfile(file_path):
                try:
                    print(f"Attaching file: {file_path}")
                    mail.Attachments.Add(Source=file_path)
                except Exception as attach_err:
                    print(f"Error attaching {filename}: {attach_err}")

        # Display email before sending (for debugging)
        mail.Display()

        # Save to drafts first
        mail.Save()
        print("Email saved to drafts")

        # Try to send
        try:
            mail.Send()
            print(f"Email sent successfully to {target_email}")
        except Exception as send_err:
            print(f"Failed to send email automatically: {send_err}")
            print("Email is saved in drafts folder - please send manually")

    except Exception as e:
        print(f"Error setting up email: {e}")
        print(f"Detailed error: {str(e)}")
        try:
            # Fallback: Save to drafts if everything else fails
            if 'mail' in locals():
                mail.Save()
                print("Email saved to drafts as fallback")
        except:
            print("Could not even save to drafts")


def process_attachments(input_folder, output_folder, watermark_text):
    # Process all files in the input folder
    for filename in os.listdir(input_folder):
        input_file_path = os.path.join(input_folder, filename)
        output_file_path = os.path.join(output_folder, filename)

        # If it's a PDF, add watermark
        if filename.lower().endswith('.pdf'):
            print(f"Watermarking PDF: {filename}")
            add_watermark(input_file_path, output_file_path, watermark_text)
        # For non-PDF files, just copy them
        else:
            print(f"Copying file: {filename}")
            shutil.copy2(input_file_path, output_file_path)


def monitor_outlook_inbox():
    # Create project directories
    base_attachments_dir = os.path.join(os.getcwd(), 'attachments')
    output_attachments_dir = os.path.join(os.getcwd(), 'output_attachments')
    os.makedirs(base_attachments_dir, exist_ok=True)
    os.makedirs(output_attachments_dir, exist_ok=True)

    # Create Outlook application object
    outlook = client.Dispatch('Outlook.Application')
    namespace = outlook.GetNamespace('MAPI')
    inbox = namespace.GetDefaultFolder(6)

    processed_emails = set()

    print("Starting inbox monitoring...")
    print(f"Watching for emails from: noaamaman325158@gmail.com")

    try:
        while True:
            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)

            for message in messages:
                email_id = f"{message.EntryID}"

                if email_id in processed_emails:
                    continue

                if message.SenderEmailAddress.lower() == "noaamaman325158@gmail.com":
                    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    unique_email_id = generate_email_id(message)

                    print("\nNew email detected!")
                    print(f"Email ID: {unique_email_id}")
                    print(f"Time detected: {current_time}")
                    print(f"Subject: {message.Subject}")

                    if message.Attachments.Count > 0:
                        # Create input and output folders for this email
                        email_folder_name = f"email_{unique_email_id}"
                        input_attachments_dir = os.path.join(base_attachments_dir, email_folder_name)
                        output_attachments_dir_email = os.path.join(output_attachments_dir, email_folder_name)

                        os.makedirs(input_attachments_dir, exist_ok=True)
                        os.makedirs(output_attachments_dir_email, exist_ok=True)

                        # Save original attachments
                        for attachment in message.Attachments:
                            attachment_path = os.path.join(input_attachments_dir, attachment.FileName)
                            print(f"Saving attachment: {attachment.FileName}")
                            attachment.SaveAsFile(attachment_path)

                        # Process PDFs with watermark
                        process_attachments(
                            input_attachments_dir,
                            output_attachments_dir_email,
                            f"Processed {unique_email_id}"
                        )

                        print(f"Original attachments saved to: {input_attachments_dir}")
                        print(f"Processed attachments saved to: {output_attachments_dir_email}")

                        # Send processed files to target email
                        send_processed_files(
                            output_attachments_dir_email,
                            "noaamaman2@gmail.com",
                            unique_email_id
                        )

                    print("-" * 50)
                    processed_emails.add(email_id)

            time.sleep(10)

    except KeyboardInterrupt:
        print("\nMonitoring stopped by user")
    except Exception as e:
        print(f"\nAn error occurred: {e}")


if __name__ == "__main__":
    monitor_outlook_inbox()