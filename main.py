import win32com.client as client
import win32gui
import os
import time
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.colors import red
import io
import hashlib


class EmailTracker:
    def __init__(self):
        self.processed_signatures = set()

    def generate_signature(self, email_item):
        try:
            parts = [
                email_item.Subject,
                str(email_item.Attachments.Count)
            ]
            for attachment in email_item.Attachments:
                parts.append(attachment.FileName)
            signature = "||".join(parts)
            return hashlib.md5(signature.encode()).hexdigest()
        except Exception as e:
            print(f"Error generating email signature: {e}")
            return datetime.now().strftime("%Y%m%d%H%M%S")

    def is_processed(self, email_item):
        signature = self.generate_signature(email_item)
        return signature in self.processed_signatures

    def mark_processed(self, email_item):
        signature = self.generate_signature(email_item)
        self.processed_signatures.add(signature)
        return signature


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


def identify_new_email_tab(email_tracker):
    try:
        outlook = client.Dispatch("Outlook.Application")
        for inspector in outlook.Inspectors:
            current_item = inspector.CurrentItem
            if current_item and current_item.Class == 43:
                if not current_item.Sent and not email_tracker.is_processed(current_item):
                    if current_item.Attachments.Count > 0:
                        # Skip if the subject starts with "Processed:"
                        if current_item.Subject.startswith("Processed:"):
                            continue
                        return current_item, inspector
        return None, None
    except Exception as e:
        print(f"Error accessing Outlook: {e}")
        return None, None


def process_and_create_new_email(message, input_folder, output_folder, watermark_text, original_inspector):
    try:
        processed_files = {}
        for attachment in message.Attachments:
            if not attachment.FileName.lower().endswith('.pdf'):
                print(f"Skipping non-PDF file: {attachment.FileName}")
                continue
            input_path = os.path.join(input_folder, attachment.FileName)
            output_path = os.path.join(output_folder, attachment.FileName)
            attachment.SaveAsFile(input_path)
            add_watermark(input_path, output_path, watermark_text)
            processed_files[attachment.FileName] = output_path

        if not processed_files:
            return False

        outlook = client.Dispatch("Outlook.Application")
        new_mail = outlook.CreateItem(0)
        new_mail.Subject = f"Processed: {message.Subject}"
        new_mail.Body = "Please find the processed attachments."

        for filename, filepath in processed_files.items():
            try:
                new_mail.Attachments.Add(Source=filepath)
            except Exception as e:
                print(f"Error adding attachment {filename}: {e}")

        new_mail.Display()

        try:
            if original_inspector:
                original_inspector.Close(0)
        except Exception as e:
            print(f"Warning: Could not close original email window: {e}")

        try:
            for file in os.listdir(input_folder):
                os.remove(os.path.join(input_folder, file))
            for file in os.listdir(output_folder):
                os.remove(os.path.join(output_folder, file))
        except Exception as e:
            print(f"Warning: Could not clean up some temporary files: {e}")

        return True

    except Exception as e:
        print(f"Error creating new email: {e}")
        return False


# For future usage
def is_valid_filename_pattern(filename):
    # Check if filename matches pattern: PQ24000001 - הדפסת הצעת מחיר-פלדום פינגולד.pdf
    if not filename.endswith('.pdf'):
        return False

    parts = filename.split(' - ')
    if len(parts) != 2:
        return False

    id_part = parts[0]
    if not (id_part.startswith('PQ') and len(id_part) >= 8 and id_part[2:].isdigit()):
        return False

    if parts[1] != 'הדפסת הצעת מחיר-פלדום פינגולד.pdf':
        return False

    return True


def monitor_outlook():
    base_attachments_dir = os.path.join(os.getcwd(), 'attachments')
    output_attachments_dir = os.path.join(os.getcwd(), 'output_attachments')
    os.makedirs(base_attachments_dir, exist_ok=True)
    os.makedirs(output_attachments_dir, exist_ok=True)

    email_tracker = EmailTracker()

    print("Starting Outlook monitoring...")
    print("Looking for new email composition windows with attachments...")
    print("Press Ctrl+C to stop monitoring")

    try:
        while True:
            new_email, inspector = identify_new_email_tab(email_tracker)

            if new_email:
                print("\nFound new email composition window:")
                print(f"Subject: {new_email.Subject}")
                print("\nAttachments found:")
                for attachment in new_email.Attachments:
                    print(f" - {attachment.FileName}")

                timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                email_folder_name = f"email_{timestamp}"
                input_attachments_dir = os.path.join(base_attachments_dir, email_folder_name)
                output_attachments_dir_email = os.path.join(output_attachments_dir, email_folder_name)

                os.makedirs(input_attachments_dir, exist_ok=True)
                os.makedirs(output_attachments_dir_email, exist_ok=True)

                success = process_and_create_new_email(
                    new_email,
                    input_attachments_dir,
                    output_attachments_dir_email,
                    f"Processed {timestamp}",
                    inspector
                )

                if success:
                    email_tracker.mark_processed(new_email)
                    print(f"Successfully processed email. Continuing to monitor for new emails...")
                    print("-" * 50)

            time.sleep(2)

    except KeyboardInterrupt:
        print("\nMonitoring stopped by user")
    except Exception as e:
        print(f"\nAn error occurred: {e}")
    finally:
        print("\nMonitoring ended")


if __name__ == "__main__":
    monitor_outlook()
