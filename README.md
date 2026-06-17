# Document Converter

A modern, high-performance batch PDF and DOCX conversion utility built using **PySide6**. Inspired by the clean, modular layout of Umi-OCR / DOC-OCR, this tool allows for seamless batch document conversions in a responsive desktop environment.

---

## Key Features

- **Dual-Pane Batch Interface**: Quickly load and inspect files in a structured queue, and configure parameters on a dedicated settings sidebar.
- **Native Drag-and-Drop**: Drag PDF or DOCX files directly from your system file manager into the application window to add them instantly.
- **Asynchronous Execution Queue**: Tasks are executed on background threads. The UI remains fully responsive, preventing freezing or lockups.
- **Parallel Task Support**: Configure up to 4 concurrent conversions from the Settings tab to maximize performance.
- **Collapsible Real-Time Logs**: View step-by-step processing outputs in a toggleable command terminal at the bottom of the window.
- **Multi-Theme Support**: Instantly toggle between **Dark Theme** and **Light Theme** styles.
- **Dynamic Localization**: Full language support for **English** and **Vietnamese (Tiếng Việt)**.
- **Post-Action Operations**: Automatically open destination folders and play system audio notifications upon batch completion.
- **Settings Safety**: Toggle the "Lock Settings" header switch to freeze your configuration and avoid accidental adjustments mid-run.

---

## System Dependencies

To convert documents, the application utilizes:
1. **Python 3.9+**
2. **LibreOffice** (running in headless mode for DOCX to PDF conversions)
3. **pdf2docx** (for converting PDF documents back to DOCX)

### Installing LibreOffice

Ensure LibreOffice is installed and available in your system path:
- **Ubuntu/Debian**:
  ```bash
  sudo apt update
  sudo apt install libreoffice
  ```
- **macOS** (via Homebrew):
  ```bash
  brew install libreoffice
  ```
- **Windows**: Download and run the installer from the [official LibreOffice website](https://www.libreoffice.org/download/download/).

---

## Installation & Setup

1. Clone or copy this repository to your local workspace.
2. Install the required Python packages:
   ```bash
   pip install -r requirement.txt
   ```

---

## Usage

1. Launch the application:
   ```bash
   python3 main.py
   ```
2. **Add Files**: Drag and drop documents onto the window, or click the **Add Files** button at the bottom of the table.
3. **Configure Settings**:
   - Select the output saving path (use the source directory or set a custom target).
   - Configure conversion direction: let the engine auto-detect based on file extension, or force a global direction.
   - Adjust post-run settings.
4. **Convert**: Click **Start Batch Conversion** to execute. You can view logs inside the Collapsible Log Terminal.

---

# Bộ chuyển đổi tài liệu (Document Converter)

Công cụ chuyển đổi tài liệu hàng loạt PDF và DOCX hiệu năng cao, nhẹ nhàng, được xây dựng bằng **PySide6**. Lấy cảm hứng từ giao diện dạng thẻ, gọn gàng của Umi-OCR / DOC-OCR, ứng dụng này giúp chuyển đổi tài liệu hàng loạt một cách mượt mà trên môi trường máy tính để bàn.

---

## Các tính năng chính

- **Giao diện chuyển đổi hàng loạt hai khung**: Dễ dàng tải và kiểm tra các tệp trong hàng đợi cấu trúc, và cấu hình các tùy chọn trên bảng cài đặt riêng biệt.
- **Kéo và thả tự nhiên**: Kéo các tệp PDF hoặc DOCX trực tiếp từ trình quản lý tệp hệ thống vào ứng dụng để thêm chúng ngay lập tức.
- **Hàng đợi thực thi bất đồng bộ**: Các tác vụ được chạy trên luồng nền (background threads). Giao diện luôn phản hồi mượt mà, không bị treo hay đơ.
- **Hỗ trợ xử lý song song**: Cấu hình tối đa 4 tác vụ chuyển đổi song song từ tab Cài đặt để tối ưu hiệu suất.
- **Nhật ký thời gian thực thu gọn**: Theo dõi kết quả thực thi từng bước trong bảng điều khiển logs có thể thu gọn/mở rộng ở phía dưới.
- **Hỗ trợ đa giao diện**: Chuyển đổi nhanh chóng giữa chế độ **Giao diện tối (Dark Theme)** và **Giao diện sáng (Light Theme)**.
- **Đa ngôn ngữ động**: Hỗ trợ đầy đủ cho tiếng **Anh (English)** và tiếng **Việt (Tiếng Việt)**.
- **Tác vụ sau khi hoàn thành**: Tự động mở thư mục lưu và phát âm thanh thông báo của hệ thống khi hoàn tất chuyển đổi hàng loạt.
- **Khóa cài đặt an toàn**: Nút gạt "Khóa cài đặt" ở thanh tiêu đề giúp cố định cấu hình, tránh thay đổi ngoài ý muốn trong lúc đang chạy.

---

## Yêu cầu hệ thống

Để thực hiện chuyển đổi, ứng dụng sử dụng các thành phần sau:
1. **Python 3.9+**
2. **LibreOffice** (chạy ở chế độ headless để chuyển từ DOCX sang PDF)
3. **pdf2docx** (để chuyển từ PDF sang DOCX)

### Cài đặt LibreOffice

Hãy đảm bảo LibreOffice đã được cài đặt và cấu hình trong đường dẫn hệ thống của bạn:
- **Ubuntu/Debian**:
  ```bash
  sudo apt update
  sudo apt install libreoffice
  ```
- **macOS** (qua Homebrew):
  ```bash
  brew install libreoffice
  ```
- **Windows**: Tải và chạy bộ cài từ [trang chủ LibreOffice](https://www.libreoffice.org/download/download/).

---

## Cài đặt & Thiết lập

1. Sao chép (clone/copy) thư mục này về không gian làm việc cục bộ của bạn.
2. Cài đặt các gói Python cần thiết:
   ```bash
   pip install -r requirement.txt
   ```

---

## Hướng dẫn sử dụng

1. Khởi chạy ứng dụng:
   ```bash
   python3 main.py
   ```
2. **Thêm tệp**: Kéo và thả các tài liệu vào cửa sổ ứng dụng, hoặc nhấn nút **Thêm tệp** ở cuối bảng.
3. **Cấu hình cài đặt**:
   - Chọn vị trí lưu tệp đầu ra (lưu trong cùng thư mục gốc hoặc đặt một thư mục tùy chỉnh).
   - Chọn hướng chuyển đổi: để hệ thống tự động nhận diện theo đuôi tệp, hoặc ép buộc một hướng chuyển đổi cụ thể.
   - Điều chỉnh các cài đặt sau khi hoàn thành.
4. **Chuyển đổi**: Bấm **Bắt đầu chuyển đổi** để thực thi. Bạn có thể theo dõi tiến trình trong bảng logs bên dưới.

