import win32com.client as client
import time
from win32gui import FindWindow, GetWindowText, EnumWindows, GetClassName, GetWindowRect
import win32gui
import win32con
import os
from datetime import datetime


def close_outlook_dialogs():
    """Find and close any open Outlook dialog boxes."""

    def enum_windows_callback(hwnd, dialog_windows):
        class_name = GetClassName(hwnd)
        window_text = GetWindowText(hwnd)

        outlook_dialog_classes = [
            "#32770",
            "DIMDialog",
            "_WwG",
            "bosa_sdm_Microsoft Office Outlook"
        ]

        if class_name in outlook_dialog_classes and window_text:
            dialog_windows.append(hwnd)
        return True

    dialog_windows = []
    EnumWindows(enum_windows_callback, dialog_windows)

    for hwnd in dialog_windows:
        try:
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            print(f"Closing dialog: {GetWindowText(hwnd)}")
            time.sleep(0.5)
        except Exception as e:
            print(f"Failed to close dialog: {e}")


def get_watermark_files(timestamp):
    """Get the watermarked files from the watermark folder."""
    watermark_folder = os.path.join(os.getcwd(), "attachments_watermark", f"message_{timestamp}")
    start_time = time.time()
    timeout = 30  # 30 seconds timeout

    print(f"\nממתין ליצירת תיקיית סימני המים: message_{timestamp}")

    while (time.time() - start_time) < timeout:
        if os.path.exists(watermark_folder):
            print(f"\nנוצרה תיקיית סימני המים: message_{timestamp}")
            watermarked_files = []
            for file in os.listdir(watermark_folder):
                watermarked_files.append(os.path.join(watermark_folder, file))
            return watermarked_files
        time.sleep(1)

    print("\nתם זמן ההמתנה - לא נוצרה תיקיית סימני מים")
    return None


def replace_attachments(message, timestamp):
    """Replace original attachments with watermarked versions."""
    try:
        watermarked_files = get_watermark_files(timestamp)
        if not watermarked_files:
            print("לא נמצאו קבצים עם סימני מים")
            return False

        # Remove existing attachments
        while message.Attachments.Count > 0:
            message.Attachments.Item(1).Delete()

        # Add watermarked attachments
        for file_path in watermarked_files:
            message.Attachments.Add(file_path)
            print(f"הוחלף נספח עם גרסה מוטבעת: {os.path.basename(file_path)}")

        return True

    except Exception as e:
        print(f"שגיאה בהחלפת הנספחים: {str(e)}")
        close_outlook_dialogs()
        return False


def save_attachments(message, save_folder):
    try:
        if message.Attachments.Count > 0:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            message_folder = os.path.join(save_folder, f"message_{timestamp}")
            os.makedirs(message_folder, exist_ok=True)
            filename_auxiliary = None
            for attachment in message.Attachments:
                try:
                    filename_auxiliary = attachment.FileName
                    file_path = os.path.join(message_folder, filename_auxiliary)
                    attachment.SaveAsFile(file_path)
                    print(f"נשמר נספח: {attachment.FileName}")
                except Exception as e:
                    print(f"שגיאה בשמירת נספח {attachment.FileName}: {str(e)}")
                    close_outlook_dialogs()
                    time.sleep(1)
                    attachment.SaveAsFile(file_path)

            output_dir_path = "C:\Users\Admin\PycharmProjects\outlookIntegration\attachments_watermark"
            # After saving original attachments, replace them with watermarked versions
            if replace_attachments(message, timestamp):
                print("הנספחים הוחלפו בהצלחה עם גרסאות מוטבעות")
            else:
                print("שגיאה בהחלפת הנספחים עם גרסאות מוטבעות")

            return True
    except Exception as e:
        print(f"שגיאה בשמירת הנספחים: {str(e)}")
        close_outlook_dialogs()
    return False


def callback(hwnd, windows):
    window_text = GetWindowText(hwnd)
    class_name = win32gui.GetClassName(hwnd)

    if class_name == "rctrl_renwnd32" and window_text:
        if window_text.startswith("Untitled") or "Message" in window_text:
            windows.append({
                'handle': hwnd,
                'title': window_text,
                'class': class_name
            })
    return True


def monitor_outlook_new_message():
    print("מתחיל לנטר חלונות הודעה חדשה ב-Outlook...")
    previous_windows = set()

    try:
        outlook = client.Dispatch("Outlook.Application")
        attachments_folder = os.path.join(os.getcwd(), "attachments")
        os.makedirs(attachments_folder, exist_ok=True)

        while True:
            try:
                outlook_windows = []
                EnumWindows(callback, outlook_windows)

                current_windows = set(window['handle'] for window in outlook_windows)
                new_windows = current_windows - previous_windows

                if new_windows:
                    for window in outlook_windows:
                        if window['handle'] in new_windows:
                            print("\nנפתח חלון הודעה חדש!")
                            print(f"כותרת: {window['title']}")

                            try:
                                close_outlook_dialogs()
                                inspector = outlook.ActiveInspector()
                                if inspector:
                                    current_item = inspector.CurrentItem
                                    if current_item:
                                        if save_attachments(current_item, attachments_folder):
                                            print("הנספחים נשמרו והוחלפו בהצלחה")
                                        else:
                                            print("אין נספחים בהודעה או שהייתה שגיאה")
                            except Exception as e:
                                print(f"שגיאה בגישה להודעה: {str(e)}")
                                close_outlook_dialogs()

                previous_windows = current_windows
                time.sleep(0.5)

            except Exception as e:
                print(f"שגיאה במהלך הניטור: {str(e)}")
                close_outlook_dialogs()
                time.sleep(1)

    except KeyboardInterrupt:
        print("הניטור הופסק")
    finally:
        outlook = None


if __name__ == "__main__":
    try:
        monitor_outlook_new_message()
    except Exception as e:
        print(f"שגיאה קריטית: {str(e)}")
        input("Press Enter to exit...")