import sys
import os
import time
from PySide6.QtCore import Qt, QThread, Signal, QSize, QUrl
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QTabBar, QGroupBox, QPushButton, QLabel, QCheckBox,
    QRadioButton, QComboBox, QLineEdit, QFileDialog, QTableWidget,
    QTableWidgetItem, QProgressBar, QTextEdit, QStackedWidget, QHeaderView,
    QFrame, QMessageBox
)
from PySide6.QtGui import QIcon, QFont, QColor, QDesktopServices, QDragEnterEvent, QDropEvent, QTextCursor
from converter import docx_to_pdf, pdf_to_docx

# ----------------- Translation Dictionary -----------------
TRANSLATIONS = {
    "en": {
        "title": "Document Converter",
        "tab_convert": "Batch Convert",
        "tab_settings": "Global Settings",
        "tab_about": "About",
        "always_on_top": "Always on Top",
        "lock_layout": "Lock Settings",
        "drop_title": "Drag & Drop Files Here",
        "drop_subtitle": "Supports PDF and DOCX files. Click to select files manually.",
        "btn_add_files": "Add Files",
        "btn_clear_list": "Clear List",
        "col_name": "File Name",
        "col_size": "Size",
        "col_type": "Convert Option",
        "col_status": "Status",
        "col_progress": "Progress",
        "col_action": "Action",
        "cfg_title": "Task Configuration",
        "cfg_direction": "Conversion Direction",
        "cfg_dir_auto": "Auto-detect by extension",
        "cfg_dir_docx2pdf": "DOCX → PDF",
        "cfg_dir_pdf2docx": "PDF → DOCX",
        "cfg_save_title": "Output Save Location",
        "cfg_save_same": "Save in the source directory",
        "cfg_save_custom": "Custom output directory",
        "cfg_browse": "Browse",
        "cfg_post_title": "Post-Conversion Actions",
        "cfg_post_open": "Open destination folder when done",
        "cfg_post_sound": "Play system notification sound",
        "btn_start": "Start Batch Conversion",
        "btn_stop": "Stop Conversion",
        "status_ready": "Ready",
        "status_converting": "Converting...",
        "status_completed": "Completed",
        "status_failed": "Failed",
        "console_title": "Real-time Execution Output",
        "settings_ui_title": "Appearance & UI",
        "settings_theme": "Theme Mode",
        "settings_lang": "Language",
        "settings_concurrency": "Max Parallel Tasks",
        "settings_theme_dark": "Dark Theme",
        "settings_theme_light": "Light Theme",
        "about_desc": "Document Converter is a lightweight, high-performance batch PDF and DOCX conversion utility built using PySide6. Inspired by the clean, modular layout of DOC-OCR.",
        "about_help": "How to use:\n1. Drag and drop PDF or DOCX files into the application.\n2. Choose output directory and conversion settings on the right panel.\n3. Click 'Start Batch Conversion' and view real-time log details below.",
        "about_ver": "Version: 1.0.0 (Stable)\nFramework: PySide 6.11\nRenderer: Headless LibreOffice & pdf2docx"
    },
    "vi": {
        "title": "Bộ chuyển đổi tài liệu Doc-OCR",
        "tab_convert": "Chuyển đổi hàng loạt",
        "tab_settings": "Cài đặt chung",
        "tab_about": "Giới thiệu",
        "always_on_top": "Luôn trên cùng",
        "lock_layout": "Khóa cài đặt",
        "drop_title": "Kéo & Thả tệp vào đây",
        "drop_subtitle": "Hỗ trợ tệp PDF và DOCX. Bấm để chọn tệp thủ công.",
        "btn_add_files": "Thêm tệp",
        "btn_clear_list": "Xóa danh sách",
        "col_name": "Tên tệp",
        "col_size": "Dung lượng",
        "col_type": "Tùy chọn chuyển đổi",
        "col_status": "Trạng thái",
        "col_progress": "Tiến trình",
        "col_action": "Hành động",
        "cfg_title": "Cấu hình tác vụ",
        "cfg_direction": "Hướng chuyển đổi",
        "cfg_dir_auto": "Tự động nhận diện",
        "cfg_dir_docx2pdf": "DOCX → PDF",
        "cfg_dir_pdf2docx": "PDF → DOCX",
        "cfg_save_title": "Nơi lưu tệp đầu ra",
        "cfg_save_same": "Lưu cùng thư mục gốc",
        "cfg_save_custom": "Thư mục lưu tùy chỉnh",
        "cfg_browse": "Duyệt",
        "cfg_post_title": "Hành động sau khi hoàn thành",
        "cfg_post_open": "Mở thư mục lưu khi hoàn thành",
        "cfg_post_sound": "Phát âm thanh thông báo",
        "btn_start": "Bắt đầu chuyển đổi",
        "btn_stop": "Dừng chuyển đổi",
        "status_ready": "Sẵn sàng",
        "status_converting": "Đang chuyển...",
        "status_completed": "Hoàn thành",
        "status_failed": "Thất bại",
        "console_title": "Thông tin thực thi thời gian thực",
        "settings_ui_title": "Giao diện & Ngôn ngữ",
        "settings_theme": "Chế độ giao diện",
        "settings_lang": "Ngôn ngữ",
        "settings_concurrency": "Số tác vụ song song tối đa",
        "settings_theme_dark": "Giao diện tối",
        "settings_theme_light": "Giao diện sáng",
        "about_desc": "DOC Document Converter là một công cụ chuyển đổi hàng loạt tệp PDF và DOCX hiệu năng cao, nhẹ nhàng, được xây dựng bằng PySide6. Lấy cảm hứng từ giao diện gọn gàng, dạng thẻ của DOC-OCR.",
        "about_help": "Hướng dẫn sử dụng:\n1. Kéo thả các tệp PDF hoặc DOCX vào ứng dụng.\n2. Chọn thư mục lưu và hướng chuyển đổi ở bảng cấu hình bên phải.\n3. Bấm 'Bắt đầu chuyển đổi' và theo dõi nhật ký thực thi chi tiết bên dưới.",
        "about_ver": "Phiên bản: 1.0.0 (Ổn định)\nNền tảng: PySide 6.11\nBộ dựng: Headless LibreOffice & pdf2docx"
    }
}

# ----------------- Stylesheets (QSS) -----------------
DARK_THEME_QSS = """
QMainWindow {
    background-color: #121214;
}
QWidget {
    background-color: #1e1e24;
    color: #cbd5e1;
    font-family: 'Segoe UI', 'Inter', -apple-system, sans-serif;
    font-size: 13px;
}

#HeaderBar {
    background-color: #121214;
    border-bottom: 1px solid #27272a;
}
#AppLogo {
    font-size: 18px;
    font-weight: bold;
    color: #3b82f6;
}
#HeaderBtn {
    background-color: #27272c;
    border: 1px solid #3f3f46;
    border-radius: 6px;
    padding: 6px 12px;
    color: #cbd5e1;
    font-weight: 500;
}
#HeaderBtn:hover {
    background-color: #3b82f6;
    border-color: #3b82f6;
    color: white;
}
#HeaderBtn:checked {
    background-color: #2563eb;
    border-color: #2563eb;
    color: white;
}

QTabWidget::pane {
    border: none;
    background-color: #121214;
}
QTabWidget {
    background-color: #121214;
}
QTabBar::tab {
    background-color: #121214;
    color: #a1a1aa;
    padding: 12px 24px;
    font-weight: 600;
    border-bottom: 2px solid transparent;
}
QTabBar::tab:hover {
    background-color: #1e1e24;
    color: #f4f4f5;
}
QTabBar::tab:selected {
    background-color: #1e1e24;
    color: #3b82f6;
    border-bottom: 2px solid #3b82f6;
}

#DropZone {
    border: 2px dashed #3f3f46;
    border-radius: 12px;
    background-color: #18181b;
}
#DropZone:hover {
    border-color: #3b82f6;
    background-color: #1c1d22;
}
#DropTitle {
    font-size: 18px;
    font-weight: bold;
    color: #f4f4f5;
}
#DropSubtitle {
    font-size: 13px;
    color: #71717a;
}

QTableWidget {
    background-color: #18181b;
    border: 1px solid #27272a;
    border-radius: 10px;
    gridline-color: #27272a;
    color: #e4e4e7;
}
QHeaderView::section {
    background-color: #202024;
    color: #a1a1aa;
    border: none;
    border-bottom: 1px solid #27272a;
    padding: 8px;
    font-weight: 600;
}
QTableWidget::item {
    border-bottom: 1px solid #27272a;
}
QTableWidget::item:selected {
    background-color: #27272c;
    color: #3b82f6;
}

QGroupBox {
    background-color: #18181b;
    border: 1px solid #27272a;
    border-radius: 10px;
    margin-top: 15px;
    padding: 15px 10px 10px 10px;
    font-weight: bold;
    font-size: 13px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 15px;
    padding: 0 5px;
    color: #3b82f6;
    background-color: #121214;
}

QLineEdit {
    background-color: #27272a;
    border: 1px solid #3f3f46;
    border-radius: 6px;
    padding: 6px 10px;
    color: #f4f4f5;
}
QLineEdit:focus {
    border: 1px solid #3b82f6;
}
QLineEdit:disabled {
    background-color: #18181b;
    color: #52525b;
    border-color: #27272a;
}

QComboBox {
    background-color: #27272a;
    border: 1px solid #3f3f46;
    border-radius: 6px;
    padding: 6px 10px;
    color: #f4f4f5;
}
QComboBox:focus {
    border: 1px solid #3b82f6;
}
QComboBox QAbstractItemView {
    background-color: #27272a;
    border: 1px solid #3f3f46;
    color: #f4f4f5;
    selection-background-color: #3b82f6;
}
QComboBox:disabled {
    background-color: #18181b;
    color: #52525b;
    border-color: #27272a;
}

QPushButton {
    background-color: #27272a;
    border: 1px solid #3f3f46;
    border-radius: 6px;
    padding: 6px 12px;
    color: #cbd5e1;
    font-weight: 500;
}
QPushButton:hover {
    background-color: #3b82f6;
    border-color: #3b82f6;
    color: white;
}
QPushButton:pressed {
    background-color: #2563eb;
}
QPushButton:disabled {
    background-color: #18181b;
    color: #52525b;
    border-color: #27272a;
}

#StartButton {
    background-color: #3b82f6;
    border: none;
    color: white;
    font-size: 14px;
    font-weight: bold;
    border-radius: 8px;
    padding: 12px;
}
#StartButton:hover {
    background-color: #2563eb;
}
#StartButton:pressed {
    background-color: #1d4ed8;
}
#StartButton:disabled {
    background-color: #27272a;
    color: #71717a;
}

#StopButton {
    background-color: #ef4444;
    border: none;
    color: white;
    font-size: 14px;
    font-weight: bold;
    border-radius: 8px;
    padding: 12px;
}
#StopButton:hover {
    background-color: #dc2626;
}
#StopButton:pressed {
    background-color: #b91c1c;
}

QRadioButton, QCheckBox {
    spacing: 8px;
}

#LogHeader {
    background-color: #18181b;
    border: 1px solid #27272a;
    border-radius: 6px 6px 0 0;
}
#LogConsole {
    background-color: #0c0a09;
    border: 1px solid #27272a;
    border-top: none;
    border-radius: 0 0 6px 6px;
    font-family: 'Courier New', monospace;
    color: #a3e635;
    padding: 8px;
}

QProgressBar {
    background-color: #27272a;
    border: 1px solid #3f3f46;
    border-radius: 8px;
    text-align: center;
    color: white;
}
QProgressBar::chunk {
    background-color: #3b82f6;
    border-radius: 7px;
}

QScrollBar:vertical {
    background-color: #121214;
    width: 10px;
    margin: 0;
}
QScrollBar::handle:vertical {
    background-color: #27272a;
    min-height: 20px;
    border-radius: 5px;
}
QScrollBar::handle:vertical:hover {
    background-color: #3f3f46;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
"""

LIGHT_THEME_QSS = """
QMainWindow {
    background-color: #f1f5f9;
}
QWidget {
    background-color: #ffffff;
    color: #334155;
    font-family: 'Segoe UI', 'Inter', -apple-system, sans-serif;
    font-size: 13px;
}

#HeaderBar {
    background-color: #f8fafc;
    border-bottom: 1px solid #e2e8f0;
}
#AppLogo {
    font-size: 18px;
    font-weight: bold;
    color: #2563eb;
}
#HeaderBtn {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 6px 12px;
    color: #334155;
    font-weight: 500;
}
#HeaderBtn:hover {
    background-color: #f1f5f9;
    border-color: #3b82f6;
    color: #2563eb;
}
#HeaderBtn:checked {
    background-color: #e0f2fe;
    border-color: #2563eb;
    color: #0369a1;
}

QTabWidget::pane {
    border: none;
    background-color: #f1f5f9;
}
QTabWidget {
    background-color: #f1f5f9;
}
QTabBar::tab {
    background-color: #f1f5f9;
    color: #64748b;
    padding: 12px 24px;
    font-weight: 600;
    border-bottom: 2px solid transparent;
}
QTabBar::tab:hover {
    background-color: #e2e8f0;
    color: #0f172a;
}
QTabBar::tab:selected {
    background-color: #ffffff;
    color: #2563eb;
    border-bottom: 2px solid #2563eb;
}

#DropZone {
    border: 2px dashed #cbd5e1;
    border-radius: 12px;
    background-color: #f8fafc;
}
#DropZone:hover {
    border-color: #2563eb;
    background-color: #eff6ff;
}
#DropTitle {
    font-size: 18px;
    font-weight: bold;
    color: #0f172a;
}
#DropSubtitle {
    font-size: 13px;
    color: #64748b;
}

QTableWidget {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 10px;
    gridline-color: #e2e8f0;
    color: #334155;
}
QHeaderView::section {
    background-color: #f1f5f9;
    color: #475569;
    border: none;
    border-bottom: 1px solid #cbd5e1;
    padding: 8px;
    font-weight: 600;
}
QTableWidget::item {
    border-bottom: 1px solid #e2e8f0;
}
QTableWidget::item:selected {
    background-color: #eff6ff;
    color: #2563eb;
}

QGroupBox {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 10px;
    margin-top: 15px;
    padding: 15px 10px 10px 10px;
    font-weight: bold;
    font-size: 13px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 15px;
    padding: 0 5px;
    color: #2563eb;
    background-color: #f1f5f9;
}

QLineEdit {
    background-color: #f8fafc;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 6px 10px;
    color: #0f172a;
}
QLineEdit:focus {
    border: 1px solid #2563eb;
}
QLineEdit:disabled {
    background-color: #f1f5f9;
    color: #94a3b8;
    border-color: #e2e8f0;
}

QComboBox {
    background-color: #f8fafc;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 6px 10px;
    color: #0f172a;
}
QComboBox:focus {
    border: 1px solid #2563eb;
}
QComboBox QAbstractItemView {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    color: #0f172a;
    selection-background-color: #eff6ff;
    selection-color: #2563eb;
}
QComboBox:disabled {
    background-color: #f1f5f9;
    color: #94a3b8;
    border-color: #e2e8f0;
}

QPushButton {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 6px 12px;
    color: #334155;
    font-weight: 500;
}
QPushButton:hover {
    background-color: #f8fafc;
    border-color: #2563eb;
    color: #2563eb;
}
QPushButton:pressed {
    background-color: #eff6ff;
}
QPushButton:disabled {
    background-color: #f1f5f9;
    color: #94a3b8;
    border-color: #e2e8f0;
}

#StartButton {
    background-color: #2563eb;
    border: none;
    color: white;
    font-size: 14px;
    font-weight: bold;
    border-radius: 8px;
    padding: 12px;
}
#StartButton:hover {
    background-color: #1d4ed8;
}
#StartButton:pressed {
    background-color: #1e40af;
}
#StartButton:disabled {
    background-color: #e2e8f0;
    color: #94a3b8;
}

#StopButton {
    background-color: #ef4444;
    border: none;
    color: white;
    font-size: 14px;
    font-weight: bold;
    border-radius: 8px;
    padding: 12px;
}
#StopButton:hover {
    background-color: #dc2626;
}
#StopButton:pressed {
    background-color: #b91c1c;
}

QRadioButton, QCheckBox {
    spacing: 8px;
}

#LogHeader {
    background-color: #f8fafc;
    border: 1px solid #cbd5e1;
    border-radius: 6px 6px 0 0;
}
#LogConsole {
    background-color: #0f172a;
    border: 1px solid #cbd5e1;
    border-top: none;
    border-radius: 0 0 6px 6px;
    font-family: 'Courier New', monospace;
    color: #38bdf8;
    padding: 8px;
}

QProgressBar {
    background-color: #e2e8f0;
    border: 1px solid #cbd5e1;
    border-radius: 8px;
    text-align: center;
    color: #0f172a;
    font-weight: bold;
}
QProgressBar::chunk {
    background-color: #2563eb;
    border-radius: 7px;
}

QScrollBar:vertical {
    background-color: #f1f5f9;
    width: 10px;
    margin: 0;
}
QScrollBar::handle:vertical {
    background-color: #cbd5e1;
    min-height: 20px;
    border-radius: 5px;
}
QScrollBar::handle:vertical:hover {
    background-color: #94a3b8;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
"""


# ----------------- Format Helper -----------------
def format_size(bytes_size):
    if bytes_size < 1024:
        return f"{bytes_size} B"
    elif bytes_size < 1024 * 1024:
        return f"{bytes_size / 1024:.1f} KB"
    else:
        return f"{bytes_size / (1024 * 1024):.1f} MB"


# ----------------- Asynchronous Thread Worker -----------------
class FileConvertWorker(QThread):
    progress_signal = Signal(int, int)  # row, percentage
    status_signal = Signal(int, str, str)  # row, status, output_path/error
    log_signal = Signal(str)  # log text

    def __init__(self, row, file_path, direction, output_dir):
        super().__init__()
        self.row = row
        self.file_path = file_path
        self.direction = direction
        self.output_dir = output_dir

    def run(self):
        try:
            filename = os.path.basename(self.file_path)
            self.status_signal.emit(self.row, "Converting", "")
            self.progress_signal.emit(self.row, 20)

            def log_callback(msg):
                self.log_signal.emit(f"[{filename}] {msg}")

            if self.direction == "docx2pdf":
                self.progress_signal.emit(self.row, 50)
                out = docx_to_pdf(self.file_path, self.output_dir, log_callback)
            else:
                self.progress_signal.emit(self.row, 50)
                out = pdf_to_docx(self.file_path, self.output_dir, log_callback)

            self.progress_signal.emit(self.row, 100)
            self.status_signal.emit(self.row, "Completed", out)
        except Exception as e:
            self.progress_signal.emit(self.row, 0)
            self.status_signal.emit(self.row, "Failed", str(e))


# ----------------- Custom UI Components -----------------
class DropZoneWidget(QFrame):
    clicked = Signal()
    files_dropped = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setObjectName("DropZone")

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(15)

        self.icon_label = QLabel("📥")
        self.icon_label.setStyleSheet("font-size: 48px; color: #3b82f6;")
        self.icon_label.setAlignment(Qt.AlignCenter)

        self.title_label = QLabel()
        self.title_label.setObjectName("DropTitle")
        self.title_label.setAlignment(Qt.AlignCenter)

        self.sub_label = QLabel()
        self.sub_label.setObjectName("DropSubtitle")
        self.sub_label.setAlignment(Qt.AlignCenter)
        self.sub_label.setWordWrap(True)

        layout.addWidget(self.icon_label)
        layout.addWidget(self.title_label)
        layout.addWidget(self.sub_label)

        self.setCursor(Qt.PointingHandCursor)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit()
        else:
            super().mousePressEvent(event)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            files = []
            for url in event.mimeData().urls():
                local_path = url.toLocalFile()
                if os.path.exists(local_path):
                    files.append(local_path)
            if files:
                self.files_dropped.emit(files)
            event.acceptProposedAction()


class FileDropTable(QTableWidget):
    files_dropped = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragEnabled(False)
        self.setDropIndicatorShown(True)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            files = []
            for url in event.mimeData().urls():
                local_path = url.toLocalFile()
                if os.path.exists(local_path):
                    files.append(local_path)
            if files:
                self.files_dropped.emit(files)
            event.acceptProposedAction()


class CollapsibleLogPanel(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Header bar
        self.header = QWidget()
        self.header.setObjectName("LogHeader")
        header_layout = QHBoxLayout(self.header)
        header_layout.setContentsMargins(10, 6, 10, 6)

        self.toggle_btn = QPushButton("▼")
        self.toggle_btn.setFixedWidth(24)
        self.toggle_btn.setStyleSheet("border: none; background: transparent; font-weight: bold;")

        self.title_lbl = QLabel()
        self.title_lbl.setStyleSheet("font-weight: bold; color: #3b82f6;")

        self.clear_btn = QPushButton()
        self.clear_btn.setFixedWidth(80)
        self.clear_btn.setStyleSheet("padding: 2px 8px; font-size: 11px;")

        header_layout.addWidget(self.toggle_btn)
        header_layout.addWidget(self.title_lbl)
        header_layout.addStretch()
        header_layout.addWidget(self.clear_btn)

        self.console = QTextEdit()
        self.console.setObjectName("LogConsole")
        self.console.setReadOnly(True)
        self.console.setMinimumHeight(110)
        self.console.setMaximumHeight(180)

        layout.addWidget(self.header)
        layout.addWidget(self.console)

        self.toggle_btn.clicked.connect(self.toggle_console)
        self.expanded = True

    def toggle_console(self):
        if self.expanded:
            self.console.hide()
            self.toggle_btn.setText("▲")
            self.expanded = False
        else:
            self.console.show()
            self.toggle_btn.setText("▼")
            self.expanded = True

    def append_log(self, text):
        timestamp = time.strftime("[%H:%M:%S]")
        self.console.append(f"{timestamp} {text}")
        self.console.moveCursor(QTextCursor.End)

    def clear_logs(self):
        self.console.clear()


# ----------------- Main Window -----------------
class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_lang = "en"
        self.is_converting = False
        self.active_workers = {}
        self.queue = []
        self.last_output_dir = None

        self.setMinimumSize(980, 680)

        # Main layouts
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)

        # Initialize UI Components
        self.setup_header()
        self.setup_tabs()
        
        # Apply defaults
        self.apply_theme("dark")
        self.retranslate_ui()

    def tr_text(self, key):
        return TRANSLATIONS[self.current_lang].get(key, key)

    # ----------------- Header Setup -----------------
    def setup_header(self):
        self.header_bar = QWidget()
        self.header_bar.setObjectName("HeaderBar")
        header_layout = QHBoxLayout(self.header_bar)
        header_layout.setContentsMargins(15, 10, 15, 10)

        self.logo_label = QLabel("📄 Doc Convert")
        self.logo_label.setObjectName("AppLogo")

        self.always_on_top_btn = QPushButton()
        self.always_on_top_btn.setObjectName("HeaderBtn")
        self.always_on_top_btn.setCheckable(True)
        self.always_on_top_btn.clicked.connect(self.toggle_always_on_top)

        self.lock_settings_btn = QPushButton()
        self.lock_settings_btn.setObjectName("HeaderBtn")
        self.lock_settings_btn.setCheckable(True)
        self.lock_settings_btn.clicked.connect(self.toggle_lock_settings)

        header_layout.addWidget(self.logo_label)
        header_layout.addStretch()
        header_layout.addWidget(self.always_on_top_btn)
        header_layout.addWidget(self.lock_settings_btn)

        self.main_layout.addWidget(self.header_bar)

    # ----------------- Tabs Setup -----------------
    def setup_tabs(self):
        self.tabs = QTabWidget()
        self.tabs.setTabBar(QTabBar())
        self.main_layout.addWidget(self.tabs)

        # Create tabs
        self.setup_convert_tab()
        self.setup_settings_tab()
        self.setup_about_tab()

    # ----------------- Convert Tab Setup -----------------
    def setup_convert_tab(self):
        self.convert_tab = QWidget()
        layout = QHBoxLayout(self.convert_tab)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)

        # Left Column: Stack (Drop zone vs File list) + Controls + Collapsible Console
        left_col = QVBoxLayout()
        left_col.setSpacing(10)

        self.stacked_widget = QStackedWidget()
        
        # 1. Drop Zone Widget
        self.drop_zone = DropZoneWidget()
        self.drop_zone.clicked.connect(self.add_files_dialog)
        self.drop_zone.files_dropped.connect(self.add_files)
        self.stacked_widget.addWidget(self.drop_zone)

        # 2. File List Table Widget
        self.table_container = QWidget()
        table_layout = QVBoxLayout(self.table_container)
        table_layout.setContentsMargins(0, 0, 0, 0)
        table_layout.setSpacing(10)

        self.table = FileDropTable()
        self.table.setColumnCount(7)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.files_dropped.connect(self.add_files)

        # Adjust header settings
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.setColumnHidden(1, True) # Hidden path column
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)

        self.table.setColumnWidth(2, 80)   # Size
        self.table.setColumnWidth(3, 140)  # Convert direction option
        self.table.setColumnWidth(4, 90)   # Status
        self.table.setColumnWidth(5, 120)  # Progress

        table_layout.addWidget(self.table)
        self.stacked_widget.addWidget(self.table_container)

        left_col.addWidget(self.stacked_widget, stretch=1)

        # Manage List Buttons
        btn_layout = QHBoxLayout()
        self.btn_add_files = QPushButton()
        self.btn_add_files.clicked.connect(self.add_files_dialog)
        
        self.btn_clear = QPushButton()
        self.btn_clear.clicked.connect(self.clear_list)
        self.btn_clear.setEnabled(False)

        btn_layout.addWidget(self.btn_add_files)
        btn_layout.addWidget(self.btn_clear)
        btn_layout.addStretch()
        left_col.addLayout(btn_layout)

        # Collapsible log console
        self.log_panel = CollapsibleLogPanel()
        self.log_panel.clear_btn.clicked.connect(self.log_panel.clear_logs)
        left_col.addWidget(self.log_panel)

        # Right Column: Configuration & Run Button
        right_col = QVBoxLayout()
        right_col.setContentsMargins(0, 0, 0, 0)
        right_col.setSpacing(15)
        
        self.config_box = QGroupBox()
        config_layout = QVBoxLayout(self.config_box)
        config_layout.setContentsMargins(15, 20, 15, 15)
        config_layout.setSpacing(15)

        # Group 1: Force Conversion Direction (Override)
        self.grp_dir = QGroupBox()
        grp_dir_layout = QVBoxLayout(self.grp_dir)
        self.combo_global_dir = QComboBox()
        self.combo_global_dir.currentIndexChanged.connect(self.on_global_dir_changed)
        grp_dir_layout.addWidget(self.combo_global_dir)
        config_layout.addWidget(self.grp_dir)

        # Group 2: Output Saving Path
        self.grp_save = QGroupBox()
        grp_save_layout = QVBoxLayout(self.grp_save)
        
        self.radio_save_same = QRadioButton()
        self.radio_save_same.setChecked(True)
        self.radio_save_same.toggled.connect(self.on_save_location_toggled)

        self.radio_save_custom = QRadioButton()
        self.radio_save_custom.toggled.connect(self.on_save_location_toggled)

        custom_path_layout = QHBoxLayout()
        self.line_save_dir = QLineEdit()
        self.line_save_dir.setReadOnly(True)
        self.line_save_dir.setEnabled(False)

        self.btn_browse_dir = QPushButton()
        self.btn_browse_dir.setEnabled(False)
        self.btn_browse_dir.clicked.connect(self.browse_output_dir)

        custom_path_layout.addWidget(self.line_save_dir)
        custom_path_layout.addWidget(self.btn_browse_dir)

        grp_save_layout.addWidget(self.radio_save_same)
        grp_save_layout.addWidget(self.radio_save_custom)
        grp_save_layout.addLayout(custom_path_layout)
        config_layout.addWidget(self.grp_save)

        # Group 3: Post action settings
        self.grp_post = QGroupBox()
        grp_post_layout = QVBoxLayout(self.grp_post)

        self.chk_post_open = QCheckBox()
        self.chk_post_open.setChecked(True)

        self.chk_post_sound = QCheckBox()
        self.chk_post_sound.setChecked(True)

        grp_post_layout.addWidget(self.chk_post_open)
        grp_post_layout.addWidget(self.chk_post_sound)
        config_layout.addWidget(self.grp_post)

        config_layout.addStretch()

        # Run Button
        self.btn_start = QPushButton()
        self.btn_start.setObjectName("StartButton")
        self.btn_start.clicked.connect(self.start_conversion)
        config_layout.addWidget(self.btn_start)

        right_col.addWidget(self.config_box)

        # Combine left and right columns
        layout.addLayout(left_col, stretch=7)
        layout.addLayout(right_col, stretch=3)

        self.tabs.addTab(self.convert_tab, "")

    # ----------------- Settings Tab Setup -----------------
    def setup_settings_tab(self):
        self.settings_tab = QWidget()
        layout = QVBoxLayout(self.settings_tab)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)

        # Global configs card
        self.settings_grp = QGroupBox()
        s_layout = QVBoxLayout(self.settings_grp)
        s_layout.setContentsMargins(20, 25, 20, 20)
        s_layout.setSpacing(15)

        # 1. Theme Selection
        theme_lay = QHBoxLayout()
        self.lbl_theme = QLabel()
        self.combo_theme = QComboBox()
        self.combo_theme.addItems(["Dark", "Light"])
        self.combo_theme.currentIndexChanged.connect(self.on_theme_changed)
        theme_lay.addWidget(self.lbl_theme)
        theme_lay.addWidget(self.combo_theme)
        theme_lay.addStretch()
        s_layout.addLayout(theme_lay)

        # 2. Language Selection
        lang_lay = QHBoxLayout()
        self.lbl_lang = QLabel()
        self.combo_lang = QComboBox()
        self.combo_lang.addItems(["English", "Tiếng Việt"])
        self.combo_lang.currentIndexChanged.connect(self.on_lang_changed)
        lang_lay.addWidget(self.lbl_lang)
        lang_lay.addWidget(self.combo_lang)
        lang_lay.addStretch()
        s_layout.addLayout(lang_lay)

        # 3. Parallel tasks concurrency
        concur_lay = QHBoxLayout()
        self.lbl_concurrency = QLabel()
        self.concurrency_combo = QComboBox()
        self.concurrency_combo.addItems(["1 Task", "2 Tasks", "3 Tasks", "4 Tasks"])
        self.concurrency_combo.setCurrentIndex(1) # default 2 concurrent conversions
        concur_lay.addWidget(self.lbl_concurrency)
        concur_lay.addWidget(self.concurrency_combo)
        concur_lay.addStretch()
        s_layout.addLayout(concur_lay)

        layout.addWidget(self.settings_grp)
        layout.addStretch()

        self.tabs.addTab(self.settings_tab, "")

    # ----------------- About Tab Setup -----------------
    def setup_about_tab(self):
        self.about_tab = QWidget()
        layout = QVBoxLayout(self.about_tab)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)

        # Content Card
        self.about_grp = QGroupBox()
        a_layout = QVBoxLayout(self.about_grp)
        a_layout.setContentsMargins(20, 25, 20, 20)
        a_layout.setSpacing(15)

        self.about_logo = QLabel("📄 Document Converter")
        self.about_logo.setStyleSheet("font-size: 20px; font-weight: bold; color: #3b82f6;")
        
        self.about_desc_lbl = QLabel()
        self.about_desc_lbl.setWordWrap(True)
        self.about_desc_lbl.setStyleSheet("font-size: 13px; line-height: 1.5; color: #a1a1aa;")

        self.about_help_lbl = QLabel()
        self.about_help_lbl.setWordWrap(True)
        self.about_help_lbl.setStyleSheet("line-height: 1.5; font-size: 13px; margin-top: 10px;")

        self.about_ver_lbl = QLabel()
        self.about_ver_lbl.setStyleSheet("font-size: 12px; color: #71717a;")

        a_layout.addWidget(self.about_logo)
        a_layout.addWidget(self.about_desc_lbl)
        a_layout.addWidget(self.about_help_lbl)
        a_layout.addWidget(self.about_ver_lbl)

        layout.addWidget(self.about_grp)
        layout.addStretch()

        self.tabs.addTab(self.about_tab, "")

    # ----------------- Controller / Actions -----------------
    def toggle_always_on_top(self, checked):
        # Update window state flags dynamically
        flags = self.windowFlags()
        if checked:
            flags |= Qt.WindowStaysOnTopHint
        else:
            flags &= ~Qt.WindowStaysOnTopHint
        
        # We need to save geometry, set flags, and show again
        geom = self.geometry()
        self.setWindowFlags(flags)
        self.setGeometry(geom)
        self.show()
        self.log_message(f"Always on Top: {'ON' if checked else 'OFF'}")

    def toggle_lock_settings(self, checked):
        # Disables configuration changes to prevent adjustments
        self.set_config_enabled(not checked)
        self.log_message(f"Configuration locked: {'YES' if checked else 'NO'}")

    def set_config_enabled(self, enabled):
        # If locked by user, we override setting to false. If converting, we also disable.
        is_locked = self.lock_settings_btn.isChecked()
        actual_state = enabled and not is_locked and not self.is_converting
        
        self.combo_global_dir.setEnabled(actual_state)
        self.radio_save_same.setEnabled(actual_state)
        self.radio_save_custom.setEnabled(actual_state)
        
        custom_save = self.radio_save_custom.isChecked()
        self.line_save_dir.setEnabled(actual_state and custom_save)
        self.btn_browse_dir.setEnabled(actual_state and custom_save)
        self.chk_post_open.setEnabled(actual_state)
        self.chk_post_sound.setEnabled(actual_state)

        # Allow / block changing convert direction inside cells
        for r in range(self.table.rowCount()):
            cb = self.table.cellWidget(r, 3)
            if cb:
                cb.setEnabled(actual_state)

    def on_save_location_toggled(self):
        custom_save = self.radio_save_custom.isChecked()
        self.line_save_dir.setEnabled(custom_save and not self.is_converting and not self.lock_settings_btn.isChecked())
        self.btn_browse_dir.setEnabled(custom_save and not self.is_converting and not self.lock_settings_btn.isChecked())

    def browse_output_dir(self):
        folder = QFileDialog.getExistingDirectory(self, self.tr_text("cfg_save_custom"), "")
        if folder:
            self.line_save_dir.setText(folder)
            self.log_message(f"Selected custom output path: {folder}")

    def add_files_dialog(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, self.tr_text("btn_add_files"), "", "Documents (*.pdf *.docx)"
        )
        if files:
            self.add_files(files)

    def add_files(self, paths):
        added_count = 0
        for path in paths:
            if not os.path.exists(path):
                continue
            ext = os.path.splitext(path)[1].lower()
            if ext not in [".docx", ".pdf"]:
                self.log_message(f"[Ignore] Unsupported format: {os.path.basename(path)}")
                continue

            # Duplicate check
            is_dup = False
            for r in range(self.table.rowCount()):
                if self.table.item(r, 1).text() == path:
                    is_dup = True
                    break
            if is_dup:
                continue

            row = self.table.rowCount()
            self.table.insertRow(row)

            # Col 0: File name
            filename_item = QTableWidgetItem(os.path.basename(path))
            filename_item.setFlags(filename_item.flags() ^ Qt.ItemIsEditable)
            self.table.setItem(row, 0, filename_item)

            # Col 1: Full path (Hidden)
            path_item = QTableWidgetItem(path)
            self.table.setItem(row, 1, path_item)

            # Col 2: Size
            try:
                size_bytes = os.path.getsize(path)
                size_str = format_size(size_bytes)
            except Exception:
                size_str = "Unknown"
            size_item = QTableWidgetItem(size_str)
            size_item.setFlags(size_item.flags() ^ Qt.ItemIsEditable)
            self.table.setItem(row, 2, size_item)

            # Col 3: Direct choice
            combo = QComboBox()
            combo.addItems([self.tr_text("cfg_dir_docx2pdf"), self.tr_text("cfg_dir_pdf2docx")])
            if ext == ".pdf":
                combo.setCurrentIndex(1)
            else:
                combo.setCurrentIndex(0)
            
            # Hook global sync or lock settings
            combo.setEnabled(not self.is_converting and not self.lock_settings_btn.isChecked())
            self.table.setCellWidget(row, 3, combo)

            # Col 4: Status
            status_item = QTableWidgetItem(self.tr_text("status_ready"))
            status_item.setForeground(QColor("#a1a1aa"))
            status_item.setFlags(status_item.flags() ^ Qt.ItemIsEditable)
            self.table.setItem(row, 4, status_item)

            # Col 5: Progress Bar
            pbar = QProgressBar()
            pbar.setRange(0, 100)
            pbar.setValue(0)
            pbar.setFixedHeight(8)
            pbar.setTextVisible(False)
            self.table.setCellWidget(row, 5, pbar)

            # Col 6: Delete Button
            del_btn = QPushButton("🗑️")
            del_btn.setToolTip("Remove")
            del_btn.setStyleSheet("""
                QPushButton { border: none; background: transparent; font-size: 14px; color: #ef4444; }
                QPushButton:hover { background-color: #fee2e2; border-radius: 4px; }
            """)
            del_btn.clicked.connect(self.remove_table_row)
            del_btn.setEnabled(not self.is_converting)
            self.table.setCellWidget(row, 6, del_btn)

            added_count += 1

        if added_count > 0:
            self.log_message(f"Added {added_count} files to batch convert queue.")
            self.stacked_widget.setCurrentIndex(1)
            self.btn_clear.setEnabled(True)

    def remove_table_row(self):
        btn = self.sender()
        if not btn:
            return
        
        # Scan for matching action cell widget
        for r in range(self.table.rowCount()):
            if self.table.cellWidget(r, 6) == btn:
                filename = self.table.item(r, 0).text()
                self.table.removeRow(r)
                self.log_message(f"Removed from list: {filename}")
                break

        if self.table.rowCount() == 0:
            self.stacked_widget.setCurrentIndex(0)
            self.btn_clear.setEnabled(False)

    def clear_list(self):
        if self.is_converting:
            return
        self.table.setRowCount(0)
        self.stacked_widget.setCurrentIndex(0)
        self.btn_clear.setEnabled(False)
        self.log_message("Cleared list.")

    def on_global_dir_changed(self, idx):
        # Sync all row convert options to match global override
        if idx == 0:
            # Auto-detect mode: reset row comboboxes based on extensions
            for r in range(self.table.rowCount()):
                path = self.table.item(r, 1).text()
                cb = self.table.cellWidget(r, 3)
                if cb:
                    if path.lower().endswith(".pdf"):
                        cb.setCurrentIndex(1)
                    else:
                        cb.setCurrentIndex(0)
        elif idx == 1:
            # Force DOCX to PDF
            for r in range(self.table.rowCount()):
                cb = self.table.cellWidget(r, 3)
                if cb:
                    cb.setCurrentIndex(0)
        elif idx == 2:
            # Force PDF to DOCX
            for r in range(self.table.rowCount()):
                cb = self.table.cellWidget(r, 3)
                if cb:
                    cb.setCurrentIndex(1)

    # ----------------- Core Conversion Execution -----------------
    def start_conversion(self):
        if self.is_converting:
            self.cancel_conversion()
            return

        num_rows = self.table.rowCount()
        if num_rows == 0:
            QMessageBox.warning(self, self.tr_text("title"), self.tr_text("drop_title"))
            return

        # Check custom path validity
        if self.radio_save_custom.isChecked():
            out_dir = self.line_save_dir.text().strip()
            if not out_dir:
                QMessageBox.warning(self, self.tr_text("title"), self.tr_text("cfg_save_custom"))
                return
            self.last_output_dir = out_dir
        else:
            self.last_output_dir = None

        # Lock UI
        self.is_converting = True
        self.btn_start.setText(self.tr_text("btn_stop"))
        self.btn_start.setObjectName("StopButton")
        self.btn_start.setStyleSheet("")  # resets styling to apply StopButton styling
        self.btn_clear.setEnabled(False)
        
        # Disable trash row buttons and options
        for r in range(num_rows):
            del_btn = self.table.cellWidget(r, 6)
            if del_btn:
                del_btn.setEnabled(False)
        self.set_config_enabled(False)

        # Clear progress and status items
        for r in range(num_rows):
            status_item = self.table.item(r, 4)
            status_item.setText(self.tr_text("status_ready"))
            status_item.setForeground(QColor("#a1a1aa"))
            
            pbar = self.table.cellWidget(r, 5)
            if pbar:
                pbar.setValue(0)

        self.log_message(f"=== Starting batch conversion queue ({num_rows} files) ===")
        self.queue = list(range(num_rows))
        self.active_workers = {}
        
        self.process_queue()

    def process_queue(self):
        if not self.is_converting:
            return

        # Fetch concurrency capacity limit
        concurrency = self.concurrency_combo.currentIndex() + 1 # 1 to 4

        while len(self.active_workers) < concurrency and self.queue:
            row = self.queue.pop(0)
            file_path = self.table.item(row, 1).text()

            # Determine conversion direction
            cb = self.table.cellWidget(row, 3)
            if cb.currentIndex() == 0:
                direction = "docx2pdf"
            else:
                direction = "pdf2docx"

            # Determine output directory
            if self.radio_save_same.isChecked():
                out_dir = os.path.dirname(file_path)
            else:
                out_dir = self.line_save_dir.text()

            if not self.last_output_dir:
                self.last_output_dir = out_dir

            worker = FileConvertWorker(row, file_path, direction, out_dir)
            worker.status_signal.connect(self.on_worker_status)
            worker.progress_signal.connect(self.on_worker_progress)
            worker.log_signal.connect(self.log_message)
            worker.finished.connect(lambda r=row: self.on_worker_finished(r))

            self.active_workers[row] = worker
            worker.start()

        # Check if complete
        if not self.active_workers and not self.queue:
            self.batch_completed()

    def on_worker_status(self, row, status, details):
        status_item = self.table.item(row, 4)
        if not status_item:
            return

        if status == "Converting":
            status_item.setText(self.tr_text("status_converting"))
            status_item.setForeground(QColor("#3b82f6")) # Blue
        elif status == "Completed":
            status_item.setText(self.tr_text("status_completed"))
            status_item.setForeground(QColor("#10b981")) # Green
            status_item.setToolTip(details) # show output path on hover
        elif status == "Failed":
            status_item.setText(self.tr_text("status_failed"))
            status_item.setForeground(QColor("#ef4444")) # Red
            status_item.setToolTip(details) # show error stack on hover

    def on_worker_progress(self, row, val):
        pbar = self.table.cellWidget(row, 5)
        if pbar:
            pbar.setValue(val)

    def on_worker_finished(self, row):
        if row in self.active_workers:
            self.active_workers[row].deleteLater()
            del self.active_workers[row]
        self.process_queue()

    def cancel_conversion(self):
        self.log_message("Stopping batch conversion queue...")
        self.queue.clear()
        
        # Stop and terminate any active workers
        for row, worker in list(self.active_workers.items()):
            if worker.isRunning():
                worker.terminate()
                worker.wait()
            
            # Reset row status to failed/interrupted
            status_item = self.table.item(row, 4)
            if status_item:
                status_item.setText(self.tr_text("status_failed"))
                status_item.setForeground(QColor("#ef4444"))
                status_item.setToolTip("Conversion interrupted by user.")
            
            pbar = self.table.cellWidget(row, 5)
            if pbar:
                pbar.setValue(0)
                
            del self.active_workers[row]

        self.is_converting = False
        self.btn_start.setText(self.tr_text("btn_start"))
        self.btn_start.setObjectName("StartButton")
        self.btn_start.setStyleSheet("") # Reset stylesheet styles
        
        # Re-enable UI components
        for r in range(self.table.rowCount()):
            del_btn = self.table.cellWidget(r, 6)
            if del_btn:
                del_btn.setEnabled(True)
        self.btn_clear.setEnabled(True)
        self.set_config_enabled(True)
        
        self.log_message("=== Batch conversion stopped ===")

    def batch_completed(self):
        self.is_converting = False
        self.btn_start.setText(self.tr_text("btn_start"))
        self.btn_start.setObjectName("StartButton")
        self.btn_start.setStyleSheet("")
        
        # Re-enable UI
        for r in range(self.table.rowCount()):
            del_btn = self.table.cellWidget(r, 6)
            if del_btn:
                del_btn.setEnabled(True)
        self.btn_clear.setEnabled(True)
        self.set_config_enabled(True)

        self.log_message("=== All tasks in batch completed ===")

        # Play notification beep
        if self.chk_post_sound.isChecked():
            QApplication.beep()

        # Open destination folder
        if self.chk_post_open.isChecked() and self.last_output_dir:
            if os.path.exists(self.last_output_dir):
                QDesktopServices.openUrl(QUrl.fromLocalFile(self.last_output_dir))

    def log_message(self, text):
        self.log_panel.append_log(text)

    # ----------------- Settings / Localization Slots -----------------
    def on_theme_changed(self, idx):
        if idx == 0:
            self.apply_theme("dark")
        else:
            self.apply_theme("light")

    def on_lang_changed(self, idx):
        if idx == 0:
            self.current_lang = "en"
        else:
            self.current_lang = "vi"
        self.retranslate_ui()

    def apply_theme(self, theme_name):
        app = QApplication.instance()
        if theme_name == "dark":
            app.setStyleSheet(DARK_THEME_QSS)
            self.combo_theme.setCurrentIndex(0)
        else:
            app.setStyleSheet(LIGHT_THEME_QSS)
            self.combo_theme.setCurrentIndex(1)
        self.log_message(f"Applied visual theme: {theme_name.upper()}")

    def retranslate_ui(self):
        # Update Window title
        self.setWindowTitle(self.tr_text("title"))
        
        # Header Controls
        self.always_on_top_btn.setText(self.tr_text("always_on_top"))
        self.lock_settings_btn.setText(self.tr_text("lock_layout"))

        # Tabs
        self.tabs.setTabText(0, self.tr_text("tab_convert"))
        self.tabs.setTabText(1, self.tr_text("tab_settings"))
        self.tabs.setTabText(2, self.tr_text("tab_about"))

        # Drop Zone
        self.drop_zone.title_label.setText(self.tr_text("drop_title"))
        self.drop_zone.sub_label.setText(self.tr_text("drop_subtitle"))

        # Action Buttons
        self.btn_add_files.setText("➕ " + self.tr_text("btn_add_files"))
        self.btn_clear.setText("🧹 " + self.tr_text("btn_clear_list"))
        
        if self.is_converting:
            self.btn_start.setText(self.tr_text("btn_stop"))
        else:
            self.btn_start.setText(self.tr_text("btn_start"))

        # Log Panel
        self.log_panel.title_lbl.setText(self.tr_text("console_title"))
        self.log_panel.clear_btn.setText(self.tr_text("btn_clear_list"))

        # Task Configuration right column
        self.config_box.setTitle(self.tr_text("cfg_title"))
        
        self.grp_dir.setTitle(self.tr_text("cfg_direction"))
        current_override_idx = self.combo_global_dir.currentIndex()
        self.combo_global_dir.clear()
        self.combo_global_dir.addItems([
            self.tr_text("cfg_dir_auto"),
            self.tr_text("cfg_dir_docx2pdf"),
            self.tr_text("cfg_dir_pdf2docx")
        ])
        if current_override_idx != -1:
            self.combo_global_dir.setCurrentIndex(current_override_idx)
        else:
            self.combo_global_dir.setCurrentIndex(0)

        self.grp_save.setTitle(self.tr_text("cfg_save_title"))
        self.radio_save_same.setText(self.tr_text("cfg_save_same"))
        self.radio_save_custom.setText(self.tr_text("cfg_save_custom"))
        self.btn_browse_dir.setText(self.tr_text("cfg_browse"))

        self.grp_post.setTitle(self.tr_text("cfg_post_title"))
        self.chk_post_open.setText(self.tr_text("cfg_post_open"))
        self.chk_post_sound.setText(self.tr_text("cfg_post_sound"))

        # Table Column Headers
        self.table.setHorizontalHeaderLabels([
            self.tr_text("col_name"),
            "Path",  # Hidden column
            self.tr_text("col_size"),
            self.tr_text("col_type"),
            self.tr_text("col_status"),
            self.tr_text("col_progress"),
            self.tr_text("col_action")
        ])

        # Dynamic table row cell item texts
        for r in range(self.table.rowCount()):
            # Update Convert Direction combo choices
            cb = self.table.cellWidget(r, 3)
            if cb:
                cur_idx = cb.currentIndex()
                cb.clear()
                cb.addItems([self.tr_text("cfg_dir_docx2pdf"), self.tr_text("cfg_dir_pdf2docx")])
                cb.setCurrentIndex(cur_idx)

            # Update status cells if they aren't converting/completed
            status_item = self.table.item(r, 4)
            if status_item:
                raw_status = status_item.text()
                # Find matching state
                if raw_status in [TRANSLATIONS["en"]["status_ready"], TRANSLATIONS["vi"]["status_ready"]]:
                    status_item.setText(self.tr_text("status_ready"))
                elif raw_status in [TRANSLATIONS["en"]["status_converting"], TRANSLATIONS["vi"]["status_converting"]]:
                    status_item.setText(self.tr_text("status_converting"))
                elif raw_status in [TRANSLATIONS["en"]["status_completed"], TRANSLATIONS["vi"]["status_completed"]]:
                    status_item.setText(self.tr_text("status_completed"))
                elif raw_status in [TRANSLATIONS["en"]["status_failed"], TRANSLATIONS["vi"]["status_failed"]]:
                    status_item.setText(self.tr_text("status_failed"))

        # Settings page translations
        self.settings_grp.setTitle(self.tr_text("settings_ui_title"))
        self.lbl_theme.setText(self.tr_text("settings_theme"))
        self.lbl_lang.setText(self.tr_text("settings_lang"))
        self.lbl_concurrency.setText(self.tr_text("settings_concurrency"))

        # Sync settings combos
        cur_theme_idx = self.combo_theme.currentIndex()
        self.combo_theme.clear()
        self.combo_theme.addItems([self.tr_text("settings_theme_dark"), self.tr_text("settings_theme_light")])
        self.combo_theme.setCurrentIndex(cur_theme_idx if cur_theme_idx != -1 else 0)

        # About page translations
        self.about_grp.setTitle(self.tr_text("tab_about"))
        self.about_desc_lbl.setText(self.tr_text("about_desc"))
        self.about_help_lbl.setText(self.tr_text("about_help"))
        self.about_ver_lbl.setText(self.tr_text("about_ver"))

        self.log_message(f"Language switched to: {self.current_lang.upper()}")


# ----------------- App Entry Point -----------------
if __name__ == "__main__":
    # Ensure subprocesses inherit correct environment flags on headless displays
    if sys.platform.startswith("linux") and "DISPLAY" not in os.environ:
        os.environ["QT_QPA_PLATFORM"] = "minimal"

    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec())