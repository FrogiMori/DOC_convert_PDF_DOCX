import sys
from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton,
    QVBoxLayout, QFileDialog, QLabel
)
from converter import docx_to_pdf, pdf_to_docx

class App(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Doc Converter")

        self.label = QLabel("Chọn file để convert")

        btn_docx = QPushButton("DOCX → PDF")
        btn_pdf = QPushButton("PDF → DOCX")

        btn_docx.clicked.connect(self.convert_docx)
        btn_pdf.clicked.connect(self.convert_pdf)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(btn_docx)
        layout.addWidget(btn_pdf)

        self.setLayout(layout)

    def convert_docx(self):
        file, _ = QFileDialog.getOpenFileName(self, "Chọn file DOCX", "", "*.docx")
        if file:
            out = docx_to_pdf(file)
            self.label.setText(f"Done: {out}")

    def convert_pdf(self):
        file, _ = QFileDialog.getOpenFileName(self, "Chọn file PDF", "", "*.pdf")
        if file:
            out = pdf_to_docx(file)
            self.label.setText(f"Done: {out}")

app = QApplication(sys.argv)
window = App()
window.show()
sys.exit(app.exec())
