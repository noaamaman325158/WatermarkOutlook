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


def create_watermark(watermark_text):
    print(f"Creating watermark with text: {watermark_text}")
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
    print("Watermark created successfully")
    return packet


def add_watermark(input_pdf_path, output_pdf_path, watermark_text):
    print(f"Adding watermark to PDF: {input_pdf_path}")
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
    print(f"Watermarked PDF saved to: {output_pdf_path}")


def generate_email_id(message):
    unique_string = f"{message.EntryID}_{message.ReceivedTime.strftime('%Y%m%d%H%M%S')}"
    email_hash = hashlib.md5(unique_string.encode()).hexdigest()[:8]
    return email_hash


def process_and_create_new_email(message, input_folder, output_folder, watermark_text):
    try:
        # Process attachments first
        processed_files = {}

        # Save and process each attachment
        for attachment in message.Attachments:
            if not attachment.FileName.lower().endswith('.pdf'):
                print(f"Skipping non-PDF file: {attachment.FileName}")
                continue

            input_path = os.path.join(input_folder, attachment.FileName)
            output_path = os.path.join(output_folder, attachment.FileName)

            # Save original attachment
            attachment.SaveAsFile(input_path)
            print(f"Attachment saved: {input_path}")

            # Watermark the PDF
            print(f"Watermarking PDF: {attachment.FileName}")
            add_watermark(input_path, output_path, watermark_text)

            # Store information for later use
            processed_files[attachment.FileName] = output_path

        if not processed_files:
            print("No PDF files to process.")
            return

        # Create a new email item
        outlook = client.Dispatch("Outlook.Application")
        new_mail = outlook.CreateItem(0)  # 0: olMailItem

        # Set the subject and body of the new email
        new_mail.Subject = f"Processed: {message.Subject}"
        new_mail.Body = "Please find the processed attachments."

        # Add processed attachments to the new email
        for filename, filepath in processed_files.items():
            print(f"Adding processed file: {filename}")
            try:
                new_mail.Attachments.Add(Source=filepath)
                print(f"Processed file added: {filepath}")
            except Exception as e:
                print(f"Error adding attachment {filename}: {e}")

        # Display the new email
        new_mail.Display()
        print("New email created successfully with processed attachments")

        # Clean up temporary files
        try:
            for file in os.listdir(input_folder):
                os.remove(os.path.join(input_folder, file))
            for file in os.listdir(output_folder):
                os.remove(os.path.join(output_folder, file))
            print("Temporary files cleaned up")
        except Exception as e:
            print(f"Warning: Could not clean up some temporary files: {e}")

    except Exception as e:
        print(f"Error creating new email: {e}")
        raise


def get_outlook_dialog():
    def enum_windows_callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            window_text = win32gui.GetWindowText(hwnd)
            if "Outlook" in window_text:
                windows.append((hwnd, window_text))
        return True

    windows = []
    win32gui.EnumWindows(enum_windows_callback, windows)
    return windows


def identify_current_outlook_message():
    try:
        outlook = client.Dispatch("Outlook.Application")
        inspector = outlook.ActiveInspector()
        if inspector:
            current_item = inspector.CurrentItem
            if current_item:
                print(f"Subject: {current_item.Subject}")
                print(f"Sender: {current_item.SenderName}")
                print(f"Received Time: {current_item.ReceivedTime}")
                return current_item
            else:
                print("No current item in the active inspector.")
        else:
            print("No active inspector found.")
    except Exception as e:
        print(f"Error accessing Outlook: {e}")
    return None


def identify_new_email_tab():
    try:
        outlook = client.Dispatch("Outlook.Application")
        inspector = outlook.ActiveInspector()
        if inspector:
            current_item = inspector.CurrentItem
            if current_item and current_item.Class == 43:  # 43 corresponds to olMailItem
                print("A new email tab is open.")
                print(f"Subject: {current_item.Subject}")

                # Identify attachments
                if current_item.Attachments.Count > 0:
                    print("Attachments:")
                    for attachment in current_item.Attachments:
                        print(f" - {attachment.FileName}")
                else:
                    print("No attachments found.")

                return current_item
            else:
                print("No new email tab is open.")
        else:
            print("No active inspector found.")
    except Exception as e:
        print(f"Error accessing Outlook: {e}")
    return None


def identify_new_email_tab():
    try:
        outlook = client.Dispatch("Outlook.Application")

        # Get all open inspectors
        for inspector in outlook.Inspectors:
            current_item = inspector.CurrentItem

            # Check if it's a mail item and if it's in compose mode
            if current_item and current_item.Class == 43:  # olMailItem = 43
                # Check if the item is unsent (draft/new email)
                if not current_item.Sent:
                    print("\nFound new email composition window:")
                    print(f"Subject: {current_item.Subject}")

                    # Check for attachments
                    if current_item.Attachments.Count > 0:
                        print("\nAttachments found:")
                        for attachment in current_item.Attachments:
                            print(f" - {attachment.FileName}")
                        return current_item

        return None

    except Exception as e:
        print(f"Error accessing Outlook: {e}")
        return None


def monitor_outlook():
    # Create project directories
    base_attachments_dir = os.path.join(os.getcwd(), 'attachments')
    output_attachments_dir = os.path.join(os.getcwd(), 'output_attachments')
    os.makedirs(base_attachments_dir, exist_ok=True)
    os.makedirs(output_attachments_dir, exist_ok=True)

    processed_items = set()  # Track processed items by their EntryID

    print("Starting Outlook monitoring...")
    print("Looking for new email composition windows with attachments...")
    print("Press Ctrl+C to stop monitoring")

    try:
        while True:
            # Check for new email composition windows
            new_email = identify_new_email_tab()

            if new_email and new_email.EntryID not in processed_items:
                print("\nNew email composition detected!")
                print(f"Subject: {new_email.Subject}")

                if new_email.Attachments.Count > 0:
                    print("Attachments found:")
                    for attachment in new_email.Attachments:
                        print(f" - {attachment.FileName}")

                    # Create directories for this email
                    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                    email_folder_name = f"email_{timestamp}"
                    input_attachments_dir = os.path.join(base_attachments_dir, email_folder_name)
                    output_attachments_dir_email = os.path.join(output_attachments_dir, email_folder_name)

                    os.makedirs(input_attachments_dir, exist_ok=True)
                    os.makedirs(output_attachments_dir_email, exist_ok=True)

                    # Process the email
                    process_and_create_new_email(
                        new_email,
                        input_attachments_dir,
                        output_attachments_dir_email,
                        f"Processed {timestamp}"
                    )

                    processed_items.add(new_email.EntryID)
                    print(f"Email processed successfully")
                    print("-" * 50)

            time.sleep(2)  # Check every 2 seconds

    except KeyboardInterrupt:
        print("\nMonitoring stopped by user")
    except Exception as e:
        print(f"\nAn error occurred: {e}")
        raise


if __name__ == "__main__":
    monitor_outlook()
