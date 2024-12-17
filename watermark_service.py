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


class PDFHandler(FileSystemEventHandler):
    def __init__(self, watermark_text, output_folder):
        self.watermark_text = watermark_text
        self.output_folder = output_folder

    def on_created(self, event):
        if event.is_directory:
            return None

        if event.src_path.endswith(".pdf"):
            input_pdf_path = event.src_path
            filename = os.path.basename(input_pdf_path)
            output_pdf_path = os.path.join(self.output_folder, f"watermarked_{filename}")
            print(f"Watermarking: {input_pdf_path}")
            add_watermark(input_pdf_path, output_pdf_path, self.watermark_text)
            print(f"Watermarked file saved at: {output_pdf_path}")


def monitor_folder(input_folder, output_folder, watermark_text):
    event_handler = PDFHandler(watermark_text, output_folder)
    observer = Observer()
    observer.schedule(event_handler, path=input_folder, recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


if __name__ == "__main__":
    input_folder = "/home/noaa/Desktop/test_folder"
    output_folder = "/home/noaa/Desktop/test_folder/output_folder"
    watermark_text = "Sample"

    monitor_folder(input_folder, output_folder, watermark_text)