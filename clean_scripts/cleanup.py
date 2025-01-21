import os
import time
from datetime import datetime, timedelta


def cleanup_old_files(max_age_hours=24):
    current_dir = os.getcwd()
    folders_to_clean = [
        os.path.join(current_dir, '..', 'attachments'),
        os.path.join(current_dir, '..', 'output_attachments')
    ]

    current_time = time.time()
    max_age_seconds = max_age_hours * 3600

    print(f"Starting cleanup process for files older than {max_age_hours} hours...")
    print(f"Current time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    for folder in folders_to_clean:
        if not os.path.exists(folder):
            print(f"Folder not found: {folder}")
            os.makedirs(folder)
            print(f"Created folder: {folder}")
            continue

        print(f"\nChecking folder: {folder}")
        for root, dirs, files in os.walk(folder, topdown=False):
            # First remove empty subdirectories
            for name in dirs:
                try:
                    dir_path = os.path.join(root, name)
                    if not os.listdir(dir_path):  # if directory is empty
                        os.rmdir(dir_path)
                        print(f"Removed empty directory: {dir_path}")
                except Exception as e:
                    print(f"Error removing directory {name}: {e}")

            # Then remove old files
            for name in files:
                try:
                    file_path = os.path.join(root, name)
                    file_age = current_time - os.path.getmtime(file_path)

                    if file_age > max_age_seconds:
                        os.remove(file_path)
                        print(f"Removed old file: {name}")
                except Exception as e:
                    print(f"Error removing file {name}: {e}")

    print("\nCleanup completed!")


if __name__ == "__main__":
    # Ask user for max age of files to keep
    try:
        hours = int(input("Enter maximum age of files to keep (in hours, default 24): ") or "24")
    except ValueError:
        print("Invalid input, using default 24 hours")
        hours = 24

    cleanup_old_files(hours)
    input("\nPress Enter to exit...")
