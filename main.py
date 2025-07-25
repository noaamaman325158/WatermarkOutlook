import shutil

import win32com.client as client
import os
import time
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.colors import red
import io
import hashlib
from typing import Tuple, Optional
import logging
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont



logging.basicConfig(
    filename='outlook_monitor.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Register the font - assuming you have the font file in a 'fonts' folder in your project
font_path = os.path.join(os.path.dirname(__file__), 'fonts', 'Arial.ttf')
pdfmetrics.registerFont(TTFont('HebrewFont', font_path))

logging.info("Starting Outlook monitoring...")


class EmailTracker:
    """Tracks processed emails to avoid duplicates"""

    def __init__(self):
        self.processed_signatures = set()

    def generate_signature(self, email_item) -> str:
        """Generate a unique signature for an email based on its content"""
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

    def is_processed(self, email_item) -> bool:
        """Check if an email has already been processed"""
        signature = self.generate_signature(email_item)
        return signature in self.processed_signatures

    def mark_processed(self, email_item) -> str:
        """Mark an email as processed"""
        signature = self.generate_signature(email_item)
        self.processed_signatures.add(signature)
        return signature


class PDFProcessor:
    """Handles PDF processing operations"""

    @staticmethod
    def create_watermark(watermark_text: str) -> io.BytesIO:
        """Create a watermark PDF with proper Hebrew text support"""
        # Reverse the Hebrew text for proper RTL rendering
        reversed_text = watermark_text[::-1]

        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)

        # Use appropriate font size
        can.setFont("HebrewFont", 100)
        can.setFillColor(red)
        can.setFillAlpha(0.3)

        # Get page dimensions
        page_width, page_height = letter

        # Calculate center of the page
        center_x = page_width / 2
        center_y = page_height / 2

        # Translate to the center of the page
        can.translate(center_x, center_y)

        # Rotate around the center
        can.rotate(45)

        # Get text width to center it properly
        text_width = can.stringWidth(reversed_text, "HebrewFont", 100)

        # Draw the Hebrew text centered
        can.drawString(-text_width / 2, 0, reversed_text)

        can.save()
        packet.seek(0)
        return packet

    @staticmethod
    def add_watermark(input_pdf_path: str, output_pdf_path: str, watermark_text: str):
        """Add watermark to a PDF file"""
        watermark_pdf = PdfReader(PDFProcessor.create_watermark(watermark_text))
        watermark_page = watermark_pdf.pages[0]

        reader = PdfReader(input_pdf_path)
        writer = PdfWriter()

        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            page.merge_page(watermark_page)
            writer.add_page(page)

        with open(output_pdf_path, "wb") as output_pdf:
            writer.write(output_pdf)


class OutlookMonitor:
    """Main class for monitoring and processing Outlook emails"""

    def __init__(self):
        self.email_tracker = EmailTracker()
        self.base_attachments_dir = os.path.join(os.getcwd(), 'attachments')
        self.output_attachments_dir = os.path.join(os.getcwd(), 'output_attachments')

        # Create necessary directories
        os.makedirs(self.base_attachments_dir, exist_ok=True)
        os.makedirs(self.output_attachments_dir, exist_ok=True)

    def cleanup_files(self, input_folder: str, output_folder: str):
        """Clean up temporary files and directories"""
        for folder in [input_folder, output_folder]:
            try:
                shutil.rmtree(folder)
                print(f"Deleted folder: {folder}")
            except Exception as e:
                print(f"Warning: Could not delete folder {folder}: {e}")

    def identify_new_email_tab(self) -> Tuple[Optional[object], Optional[object]]:
        """Identify new email composition windows that need processing"""
        outlook = client.Dispatch("Outlook.Application")

        for inspector in outlook.Inspectors:
            current_item = inspector.CurrentItem
            if current_item and current_item.Class == 43:
                # Check if the email is created and sent by you
                if current_item.SentOnBehalfOfName != outlook.Session.CurrentUser.Name:
                    continue

                if "Processed" in current_item.Subject:
                    continue

                if ("הדפסת הצעת מחיר" in current_item.Subject and
                        not current_item.Sent and
                        not self.email_tracker.is_processed(current_item) and
                        current_item.Attachments.Count > 0):
                    return current_item, inspector
        return None, None

    def process_attachments(self, message, input_folder: str, output_folder: str,
                            watermark_text: str) -> dict:
        """Process email attachments"""
        processed_files = {}
        for attachment in message.Attachments:
            if not attachment.FileName.lower().endswith('.pdf'):
                print(f"Skipping non-PDF file: {attachment.FileName}")
                continue

            if "הצעת מחיר" in attachment.FileName:
                print(f"Skipping empty file: {attachment.FileName}")
                input_path = os.path.join(input_folder, attachment.FileName)
                output_path = os.path.join(output_folder, attachment.FileName)
                attachment.SaveAsFile(input_path)
                shutil.copy(input_path, output_path)
                processed_files[attachment.FileName] = output_path
                continue

            # Generate new filename with watermark suffix
            filename_base, file_ext = os.path.splitext(attachment.FileName)
            new_filename = f"{filename_base}_watermark{file_ext}"

            input_path = os.path.join(input_folder, attachment.FileName)
            output_path = os.path.join(output_folder, new_filename)

            attachment.SaveAsFile(input_path)
            PDFProcessor.add_watermark(input_path, output_path, watermark_text)
            processed_files[new_filename] = output_path

        return processed_files

    def create_new_email(self, original_message, processed_files: dict) -> object:
        """Create new email with processed attachments"""
        outlook = client.Dispatch("Outlook.Application")
        new_mail = outlook.CreateItem(0)
        new_mail.Subject = f"Processed: {original_message.Subject}"
        new_mail.Body = "Please find the processed attachments."

        for filename, filepath in processed_files.items():
            try:
                new_mail.Attachments.Add(Source=filepath)
            except Exception as e:
                print(f"Error adding attachment {filename}: {e}")

        return new_mail

    def process_email(self, message, inspector) -> bool:
        """Process a single email"""
        current_year = datetime.now().year
        current_timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        email_folder_name = f"email_{current_timestamp}"

        input_folder = os.path.join(self.base_attachments_dir, email_folder_name)
        output_folder = os.path.join(self.output_attachments_dir, email_folder_name)

        os.makedirs(input_folder, exist_ok=True)
        os.makedirs(output_folder, exist_ok=True)

        try:
            # Process attachments
            processed_files = self.process_attachments(
                message, input_folder, output_folder, "דוגמא לאישור"
            )

            if not processed_files:
                return False

            # Create and display new email
            new_mail = self.create_new_email(message, processed_files)
            new_mail.Display()

            # Close original email window
            if inspector:
                try:
                    inspector.Close(0)
                except Exception as e:
                    print(f"Warning: Could not close original email window: {e}")

            # Cleanup
            self.cleanup_files(input_folder, output_folder)
            self.email_tracker.mark_processed(message)

            return True

        except Exception as e:
            print(f"Error processing email: {e}")
            return False
    def start_monitoring(self):
        """Start monitoring Outlook for new emails"""
        print("Starting Outlook monitoring...")
        print("Looking for new email composition windows with attachments...")
        print("Press Ctrl+C to stop monitoring")

        try:
            while True:
                new_email, inspector = self.identify_new_email_tab()

                if new_email:
                    print("\nFound new email composition window:")
                    print(f"Subject: {new_email.Subject}")
                    print("\nAttachments found:")
                    for attachment in new_email.Attachments:
                        print(f" - {attachment.FileName}")

                    success = self.process_email(new_email, inspector)
                    if success:
                        print(f"Successfully processed email. Continuing to monitor...")
                        print("-" * 50)

                time.sleep(2)

        except KeyboardInterrupt:
            print("\nMonitoring stopped by user")
        except Exception as e:
            print(f"\nAn error occurred: {e}")
        finally:
            print("\nMonitoring ended")


def main():
    """Main entry point"""
    monitor = OutlookMonitor()
    monitor.start_monitoring()


if __name__ == "__main__":
    main()