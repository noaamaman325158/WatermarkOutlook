from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.colors import red
import io
import os
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler


def create_watermark(watermark_text):
    packet = io.BytesIO()
    width, height = letter
    can = canvas.Canvas(packet, pagesize=letter)

    # Setting font and size
    can.setFont("Helvetica-Bold", 150)
    can.setFillColor(red)
    can.setFillAlpha(0.2)

    # Centering and rotating
    can.translate(width/2, height/2)
    can.rotate(45)

    # Draw "Sample" and "Paldom" on separate lines - single appearance
    can.drawCentredString(0, 50, "Sample")
    can.drawCentredString(0, -100, "Paldom")

    can.save()
    packet.seek(0)
    return packet


def add_watermark(input_pdf_path, output_pdf_path, watermark_text):
    try:
        watermark_pdf = PdfReader(create_watermark(watermark_text))
        watermark_page = watermark_pdf.pages[0]

        reader = PdfReader(input_pdf_path)
        writer = PdfWriter()

        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            page.merge_page(watermark_page)
            writer.add_page(page)

        os.makedirs(os.path.dirname(output_pdf_path), exist_ok=True)

        with open(output_pdf_path, "wb") as output_pdf:
            writer.write(output_pdf)
        return True
    except Exception as e:
        print(f"Error adding watermark to {input_pdf_path}: {str(e)}")
        return False


def process_existing_pdfs(input_folder, output_folder, watermark_text):
    """Process all existing PDF files in the input folder and its subfolders."""
    for root, dirs, files in os.walk(input_folder):
        for file in files:
            if file.lower().endswith('.pdf'):
                input_pdf_path = os.path.join(root, file)
                # Get relative path from input folder
                rel_path = os.path.relpath(root, input_folder)
                output_dir = os.path.join(output_folder, rel_path)

                # Create output directory if it doesn't exist
                os.makedirs(output_dir, exist_ok=True)

                output_pdf_path = os.path.join(output_dir, f"watermarked_{file}")

                print(f"Processing existing PDF: {input_pdf_path}")
                if add_watermark(input_pdf_path, output_pdf_path, watermark_text):
                    print(f"Successfully created watermarked file at: {output_pdf_path}")
                else:
                    print(f"Failed to process: {input_pdf_path}")


class PDFHandler(FileSystemEventHandler):
    def __init__(self, watermark_text, output_folder, input_folder):
        self.watermark_text = watermark_text
        self.output_folder = output_folder
        self.input_folder = input_folder

    def on_created(self, event):
        if event.is_directory:
            rel_path = os.path.relpath(event.src_path, self.input_folder)
            new_output_dir = os.path.join(self.output_folder, rel_path)
            os.makedirs(new_output_dir, exist_ok=True)
            return

        if event.src_path.endswith('.pdf'):
            try:
                input_pdf_path = event.src_path
                rel_path = os.path.relpath(os.path.dirname(input_pdf_path), self.input_folder)
                output_dir = os.path.join(self.output_folder, rel_path)
                os.makedirs(output_dir, exist_ok=True)

                filename = os.path.basename(input_pdf_path)
                output_pdf_path = os.path.join(output_dir, f"watermarked_{filename}")

                print(f"Processing new PDF: {input_pdf_path}")
                if add_watermark(input_pdf_path, output_pdf_path, self.watermark_text):
                    print(f"Successfully created watermarked file at: {output_pdf_path}")
                else:
                    print(f"Failed to process: {input_pdf_path}")
            except Exception as e:
                print(f"Error processing {event.src_path}: {str(e)}")

    def on_moved(self, event):
        if not event.is_directory and event.dest_path.endswith('.pdf'):
            self.on_created(event)


def monitor_folder(input_folder, output_folder, watermark_text):
    input_folder = os.path.abspath(input_folder)
    output_folder = os.path.abspath(output_folder)

    os.makedirs(output_folder, exist_ok=True)

    print(f"Starting PDF monitor...")
    print(f"Input folder: {input_folder}")
    print(f"Output folder: {output_folder}")
    print(f"Watermark text: {watermark_text}")

    # Process existing files first
    print("\nProcessing existing PDF files...")
    process_existing_pdfs(input_folder, output_folder, watermark_text)
    print("Finished processing existing files.\n")

    # Start monitoring for new files
    print("Starting to monitor for new files...")
    event_handler = PDFHandler(watermark_text, output_folder, input_folder)
    observer = Observer()
    observer.schedule(event_handler, path=input_folder, recursive=True)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nStopping monitor...")
        observer.stop()
    observer.join()


def set_up_config_watermark():
    base_dir = os.getcwd()
    input_folder = os.path.join(base_dir, "attachments")
    output_folder = os.path.join(base_dir, "attachments_watermark")
    watermark_text = "Sample"

    os.makedirs(input_folder, exist_ok=True)

    monitor_folder(input_folder, output_folder, watermark_text)


if __name__ == "__main__":
    try:
        set_up_config_watermark()
    except Exception as e:
        print(f"Critical error: {str(e)}")
        input("Press Enter to exit...")