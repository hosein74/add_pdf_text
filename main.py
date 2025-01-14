import sys
from PyQt5 import QtCore, QtGui, QtWidgets
import fitz  # PyMuPDF برای نمایش PDF
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from bidi.algorithm import get_display  # برای راست‌چین کردن متن
import arabic_reshaper  # برای اصلاح نمایش متن فارسی
import pandas as pd
import os
import io
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename
import webbrowser

# مقادیر پیش‌فرض
DEFAULT_FONT_PATH = './fonts/BNazanin.ttf'
DEFAULT_FONT_SIZE = 12
INITIAL_DIR = os.path.dirname(os.path.abspath(__file__))

class FontSelector(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.font_path = DEFAULT_FONT_PATH
        self.font_size = DEFAULT_FONT_SIZE
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle('انتخاب فونت و اندازه')
        self.setGeometry(100, 100, 400, 200)
        
        layout = QtWidgets.QVBoxLayout()
        
        # دکمه انتخاب فایل فونت
        self.font_button = QtWidgets.QPushButton('انتخاب فایل فونت (اختیاری)', self)
        self.font_button.clicked.connect(self.select_font)
        layout.addWidget(self.font_button)
        
        self.font_label = QtWidgets.QLabel(f'فونت پیش‌فرض: {DEFAULT_FONT_PATH}')
        layout.addWidget(self.font_label)
        
        self.font_size_label = QtWidgets.QLabel(f'اندازه فونت (پیش‌فرض: {DEFAULT_FONT_SIZE}):')
        layout.addWidget(self.font_size_label)
        
        self.font_size_input = QtWidgets.QLineEdit(self)
        layout.addWidget(self.font_size_input)
        
        self.submit_button = QtWidgets.QPushButton('تأیید', self)
        self.submit_button.clicked.connect(self.submit)
        layout.addWidget(self.submit_button)
        
        self.setLayout(layout)
        
    def select_font(self):
        # انتخاب فایل فونت از پوشه‌ی fonts کنار فایل اصلی پروژه
        self.font_path, _ = QtWidgets.QFileDialog.getOpenFileName(self, 'انتخاب فایل فونت', os.path.join(INITIAL_DIR, 'fonts'), 'Font Files (*.ttf)')
        if self.font_path:
            self.font_label.setText(f'فونت انتخاب شده: {self.font_path}')
    
    def submit(self):
        self.font_size = int(self.font_size_input.text()) if self.font_size_input.text() else DEFAULT_FONT_SIZE
        self.close()

    def get_font_settings(self):
        return self.font_path, self.font_size

def add_farsi_font(canvas, font_path, font_size):
    pdfmetrics.registerFont(TTFont('SelectedFont', font_path))
    canvas.setFont("SelectedFont", font_size)

def draw_text_rtl(canvas, text, x, y):
    reshaped_text = arabic_reshaper.reshape(text)
    bidi_text = get_display(reshaped_text)
    canvas.drawRightString(x, y, bidi_text)

def add_info_to_pdf(reader, writer, number, name, positions, font_path, font_size):
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)
    add_farsi_font(can, font_path, font_size)
    
    number_text = f"{number}"
    name_text = f"{name}"

    number_pos, name_pos = positions
    draw_text_rtl(can, number_text, *number_pos)
    draw_text_rtl(can, name_text, *name_pos)
    can.save()
    packet.seek(0)
    new_pdf = PdfReader(packet)
    overlay_page = new_pdf.pages[0]
    page_with_info = reader.pages[0]
    page_with_info.merge_page(overlay_page)
    mediabox = page_with_info.mediabox
    mediabox.lower_left = (0, 0)
    mediabox.upper_right = (mediabox.right, mediabox.top)
    writer.add_page(page_with_info)

class PdfViewer(QtWidgets.QWidget):
    def __init__(self, pdf_path):
        super().__init__()
        self.setWindowTitle('انتخاب موقعیت')
        self.resize(800, 1000)
        self.pdf_path = pdf_path
        self.selected_position = None
        self.init_ui()

    def init_ui(self):
        layout = QtWidgets.QVBoxLayout()
        self.setLayout(layout)
        self.pdf_label = QtWidgets.QLabel(self)
        layout.addWidget(self.pdf_label)
        self.load_pdf()
        self.pdf_label.mousePressEvent = self.get_position

    def load_pdf(self):
        pdf_document = fitz.open(self.pdf_path)
        page = pdf_document.load_page(0)
        self.page_width = page.rect.width
        self.page_height = page.rect.height
        pix = page.get_pixmap()
        qt_image = QtGui.QImage(pix.samples, pix.width, pix.height, pix.stride, QtGui.QImage.Format_RGB888)
        self.pdf_label.setPixmap(QtGui.QPixmap.fromImage(qt_image))
        self.pdf_label.setAlignment(QtCore.Qt.AlignCenter)

    def get_position(self, event):
        label_width = self.pdf_label.pixmap().width()
        label_height = self.pdf_label.pixmap().height()
        label_x_offset = (self.pdf_label.width() - label_width) // 2
        label_y_offset = (self.pdf_label.height() - label_height) // 2
        x_ratio = self.page_width / label_width
        y_ratio = self.page_height / label_height
        x = (event.x() - label_x_offset) * x_ratio
        y = (event.y() - label_y_offset) * y_ratio
        y = self.page_height - y
        self.selected_position = (x, y)
        self.close()

    def get_selected_position(self):
        return self.selected_position

def main():
    Tk().withdraw()  # مخفی کردن پنجره اصلی Tkinter
    pdf_path = askopenfilename(title="انتخاب فایل PDF", initialdir=INITIAL_DIR, filetypes=[("PDF files", "*.pdf")])
    if not pdf_path:
        messagebox.showerror("خطا", "لطفاً یک فایل PDF انتخاب کنید.")
        return

    excel_path = askopenfilename(title="انتخاب فایل Excel", initialdir=INITIAL_DIR, filetypes=[("Excel files", "*.xlsx")])
    if not excel_path:
        messagebox.showerror("خطا", "لطفاً یک فایل Excel انتخاب کنید.")
        return

    df = pd.read_excel(excel_path)
    app = QtWidgets.QApplication(sys.argv)
    
    # پنجره انتخاب فونت و اندازه
    font_selector = FontSelector()
    font_selector.show()
    app.exec_()
    font_path, font_size = font_selector.get_font_settings()

    column_positions = {column: None for column in ['Number', 'Name']}
    
    for column in column_positions.keys():
        viewer = PdfViewer(pdf_path)
        viewer.show()
        app.exec_()
        column_positions[column] = viewer.get_selected_position()

    writer = PdfWriter()
    for index, row in df.iterrows():
        reader = PdfReader(pdf_path)
        number = row['Number']
        name = row['Name']
        positions = [column_positions['Number'], column_positions['Name']]
        add_info_to_pdf(reader, writer, number, name, positions, font_path, font_size)

    output_path = "combined_output.pdf"
    with open(output_path, "wb") as output_file:
        writer.write(output_file)
    
    # باز کردن فایل PDF ایجاد شده
    webbrowser.open(output_path)

if __name__ == "__main__":
    main()
    