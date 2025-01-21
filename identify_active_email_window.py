import win32com.client as client


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


if __name__ == "__main__":
    identify_new_email_tab()