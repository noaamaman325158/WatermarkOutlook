import ctypes
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


class EmailTracker:
    """Tracks processed emails to avoid duplicates"""

    def __init__(self):
        self.processed_signatures = set()
        self.declined_signatures = set()

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
    
    def is_declined(self, email_item) -> bool:
        """Check if an email has been declined by the user"""
        signature = self.generate_signature(email_item)
        return signature in self.declined_signatures

    def mark_processed(self, email_item) -> str:
        """Mark an email as processed"""
        signature = self.generate_signature(email_item)
        self.processed_signatures.add(signature)
        return signature
    
    def mark_declined(self, email_item) -> str:
        """Mark an email as declined by the user"""
        signature = self.generate_signature(email_item)
        self.declined_signatures.add(signature)
        return signature


class PDFProcessor:
    """Handles PDF processing operations"""

    @staticmethod
    def create_watermark(watermark_text: str) -> io.BytesIO:
        """Create a watermark PDF as one continuous diagonal line"""
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        can.setFont("Helvetica", 45)
        can.setFillColor(red)
        can.setFillAlpha(0.2)

        # Get page dimensions
        page_width, page_height = letter

        # Start from top-left corner
        start_x = 0
        start_y = page_height

        can.saveState()
        can.translate(start_x, start_y)
        can.rotate(-45)  # Negative rotation for the diagonal direction

        # Draw all text on the same rotated line with spacing
        text = "Customer-view   -   Customer-view   -   Customer-view   -   Customer-view"
        can.drawString(0, 0, text)

        can.restoreState()

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
        self.outlook = client.Dispatch("Outlook.Application")

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
        for inspector in self.outlook.Inspectors:
            current_item = inspector.CurrentItem
            if current_item and current_item.Class == 43:
                if "Processed" in current_item.Subject:
                    continue

                if ("הדפסת הצעת מחיר" in current_item.Subject and
                        not current_item.Sent and
                        not self.email_tracker.is_processed(current_item) and
                        not self.email_tracker.is_declined(current_item) and
                        current_item.Attachments.Count > 0):
                    return current_item, inspector
        return None, None

    def confirm_processing(self, email_subject: str) -> bool:
        """Display a confirmation dialog to the user."""
        message = (
            f"Found email with subject: '{email_subject}'.\n\n"
            "This script will add a 'Customer-view' watermark to all PDF attachments "
            "and create a new email with the modified files.\n\n"
            "Do you want to proceed?"
        )
        title = "Confirm Attachment Processing"
        # MB_YESNO = 0x00000004
        # IDYES = 6
        result = ctypes.windll.user32.MessageBoxW(0, message, title, 4)
        return result == 6

    def process_attachments(self, message, input_folder: str, output_folder: str,
                            watermark_text: str) -> dict:
        """Process email attachments"""
        # Edge case: If only 2 attachments and they are PDF + XLS, remove XLS
        if message.Attachments.Count == 2:
            attachment_files = [att.FileName.lower() for att in message.Attachments]
            has_pdf = any(f.endswith('.pdf') for f in attachment_files)
            has_xls = any(f.endswith('.xls') or f.endswith('.xlsx') for f in attachment_files)
            
            if has_pdf and has_xls:
                print("Edge case detected: Found only PDF and XLS attachments. Removing XLS attachment.")
                # Remove XLS attachment(s)
                attachments_to_remove = []
                for i, attachment in enumerate(message.Attachments):
                    if (attachment.FileName.lower().endswith('.xls') or 
                        attachment.FileName.lower().endswith('.xlsx')):
                        attachments_to_remove.append(i + 1)  # COM uses 1-based indexing
                
                # Remove in reverse order to maintain correct indices
                for index in reversed(attachments_to_remove):
                    message.Attachments.Remove(index)
                    print(f"Removed XLS attachment at index {index}")

        processed_files = {}
        for i, attachment in enumerate(message.Attachments):
            # Include first attachment without manipulation
            if i == 0:
                print(f"Including first attachment without changes: {attachment.FileName}")
                input_path = os.path.join(input_folder, attachment.FileName)
                output_path = os.path.join(output_folder, attachment.FileName)
                attachment.SaveAsFile(input_path)
                import shutil
                shutil.copy2(input_path, output_path)
                processed_files[attachment.FileName] = output_path
                continue
                
            if not attachment.FileName.lower().endswith('.pdf'):
                print(f"Skipping non-PDF file: {attachment.FileName}")
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
        new_mail = self.outlook.CreateItem(0)
        new_mail.Subject = f"Processed: {original_message.Subject}"
        new_mail.Body = ""

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
                message, input_folder, output_folder, "Customer-view"
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
                try:
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
                except Exception as e:
                    if "RPC server is unavailable" in str(e) or "disconnected from its clients" in str(e):
                        print("Connection to Outlook lost. Reconnecting...")
                        self.outlook = client.Dispatch("Outlook.Application")
                    else:
                        print(f"Error occurred: {e}")
                        print("Continuing monitoring...")

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