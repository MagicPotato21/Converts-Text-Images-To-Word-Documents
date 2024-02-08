import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pytesseract
import os
import requests
import subprocess
from docx import Document

class OCRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OCR Area Selection Tool")

        self.zoom_level = 1.0

        self.scrollbar_x = tk.Scrollbar(self.root, orient="horizontal")
        self.scrollbar_x.pack(side="bottom", fill="x")

        self.scrollbar_y = tk.Scrollbar(self.root, orient="vertical")
        self.scrollbar_y.pack(side="right", fill="y")

        self.canvas = tk.Canvas(root, cursor="cross", xscrollcommand=self.scrollbar_x.set, yscrollcommand=self.scrollbar_y.set)
        self.canvas.pack(side="left", fill="both", expand=True)

        self.scrollbar_x.config(command=self.canvas.xview)
        self.scrollbar_y.config(command=self.canvas.yview)

        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_move_press)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)
        self.canvas.bind_all("<MouseWheel>", self.on_mousewheel)

        self.load_image_button = tk.Button(root, text="Load Image", command=self.load_image)
        self.load_image_button.pack(side="bottom", fill="x")

        self.rect = None
        self.start_x = None
        self.start_y = None
        self.image = None

    def load_image(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.image = Image.open(file_path)
            self.show_image()

    def show_image(self, event=None):
        if self.image:
            self.zoom_image = self.image.resize([int(self.zoom_level * s) for s in self.image.size], Image.Resampling.LANCZOS)
            self.tkimage = ImageTk.PhotoImage(self.zoom_image)
            self.canvas.create_image(0, 0, anchor="nw", image=self.tkimage)
            self.canvas.config(scrollregion=self.canvas.bbox(tk.ALL))
            self.rect = None

    def on_button_press(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x + 1, self.start_y + 1, outline='red')

    def on_move_press(self, event):
        curX = self.canvas.canvasx(event.x)
        curY = self.canvas.canvasy(event.y)
        self.canvas.coords(self.rect, self.start_x, self.start_y, curX, curY)

    def on_button_release(self, event):
        if self.start_x and self.start_y and self.image:
            end_x = self.canvas.canvasx(event.x)
            end_y = self.canvas.canvasy(event.y)
            box = [int(coord / self.zoom_level) for coord in (self.start_x, self.start_y, end_x, end_y)]
            cropped_image = self.image.crop(box)
            text = pytesseract.image_to_string(cropped_image)
            self.save_text_to_docx(text)

    def save_text_to_docx(self, text):
        file_path = filedialog.asksaveasfilename(defaultextension=".docx",
                                                 filetypes=[("Word documents", "*.docx"), ("All files", "*.*")])
        if file_path:
            doc = Document()
            doc.add_paragraph(text)
            doc.save(file_path)
            messagebox.showinfo("Info", "Text saved successfully in Word document.")

    def on_mousewheel(self, event):
        if event.delta > 0:
            self.zoom_level *= 1.1
        else:
            self.zoom_level /= 1.1
        self.zoom_level = max(min(self.zoom_level, 3), 0.5)
        self.show_image()

def download_tesseract_installer():
    tesseract_url = "https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-v5.0.0.20210503.exe"
    local_filename = "tesseract_installer.exe"
    print("Downloading Tesseract OCR installer...")
    try:
        with requests.get(tesseract_url, stream=True) as r:
            with open(local_filename, 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
        messagebox.showinfo("Download Complete", f"Installer downloaded as {local_filename}. Please run this file to install Tesseract OCR.")
    except Exception as e:
        messagebox.showerror("Download Error", f"Error downloading Tesseract: {e}")

def is_tesseract_installed_at_path(path):
    tesseract_executable = os.path.join(path, 'tesseract.exe')
    return os.path.isfile(tesseract_executable)

if __name__ == "__main__":
    tesseract_path = 'C:\\Program Files\\Tesseract-OCR'
    if not is_tesseract_installed_at_path(tesseract_path):
        if sys.platform.startswith('win'):
            response = messagebox.askyesno("Tesseract Not Found", "Tesseract OCR is not installed. Would you like to download it?")
            if response:
                download_tesseract_installer()
            sys.exit(1)
        else:
            messagebox.showerror("Tesseract Not Found", "Tesseract is not installed. Please install it manually.")
            sys.exit(1)
    else:
        pytesseract.pytesseract.tesseract_cmd = os.path.join(tesseract_path, 'tesseract.exe')
        root = tk.Tk()
        app = OCRApp(root)
        root.mainloop()
