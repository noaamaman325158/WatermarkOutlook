import win32com.client as client
import time
import os
from datetime import datetime
import hashlib


def generate_email_id(message):
    # Create a unique identifier using email properties
    unique_string = f"{message.EntryID}_{message.ReceivedTime.strftime('%Y%m%d%H%M%S')}"
    # Create a short hash (first 8 characters) to use as an ID
    email_hash = hashlib.md5(unique_string.encode()).hexdigest()[:8]
    return email_hash


def monitor_outlook_inbox():
    # Create project directories if they don't exist
    base_attachments_dir = os.path.join(os.getcwd(), 'attachments')
    os.makedirs(base_attachments_dir, exist_ok=True)

    # Create Outlook application object
    outlook = client.Dispatch('Outlook.Application')
    namespace = outlook.GetNamespace('MAPI')

    # Access the inbox
    inbox = namespace.GetDefaultFolder(6)

    # Keep track of processed emails to avoid duplicates
    processed_emails = set()

    print("Starting inbox monitoring...")
    print(f"Watching for emails from: noaamaman325158@gmail.com")

    try:
        while True:
            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)  # Sort by newest first

            # Check the most recent emails
            for message in messages:
                # Create unique identifier for the email
                email_id = f"{message.EntryID}"

                # Skip if we've already processed this email
                if email_id in processed_emails:
                    continue

                # Check if the email is from the target address
                if message.SenderEmailAddress.lower() == "noaamaman325158@gmail.com":
                    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    # Generate unique email ID
                    unique_email_id = generate_email_id(message)

                    print("\nNew email detected!")
                    print(f"Email ID: {unique_email_id}")
                    print(f"Time detected: {current_time}")
                    print(f"Subject: {message.Subject}")
                    print(f"Received: {message.ReceivedTime}")
                    print(f"Body: {message.Body[:200]}...")  # First 200 characters

                    # Handle attachments
                    if message.Attachments.Count > 0:
                        # Create a folder with the unique email ID
                        email_folder_name = f"email_{unique_email_id}"
                        email_attachments_dir = os.path.join(base_attachments_dir, email_folder_name)
                        os.makedirs(email_attachments_dir, exist_ok=True)

                        # Save each attachment
                        for attachment in message.Attachments:
                            attachment_path = os.path.join(email_attachments_dir, attachment.FileName)

                            print(f"Saving attachment: {attachment.FileName}")
                            attachment.SaveAsFile(attachment_path)

                        print(f"Attachments saved to: {email_attachments_dir}")

                    print("-" * 50)

                    # Add to processed emails
                    processed_emails.add(email_id)

            # Wait before next check
            time.sleep(10)  # Check every 10 seconds

    except KeyboardInterrupt:
        print("\nMonitoring stopped by user")
    except Exception as e:
        print(f"\nAn error occurred: {e}")


if __name__ == "__main__":
    monitor_outlook_inbox()