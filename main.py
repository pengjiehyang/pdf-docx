# --- START OF FILE main.py ---

import sys, os, time, tempfile
# Removed subprocess as terminal functionality is removed

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel,
    QProgressBar, QMessageBox, QHBoxLayout, QRadioButton,
    QButtonGroup, QComboBox, QListWidget, QListWidgetItem, QSizePolicy
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QIcon, QPixmap, QPainter, QFont
from pdf2docx import Converter
from docx import Document

try:
    import win32com.client
except ImportError:
    win32com = None

from pypdf import PdfReader

def set_emoji_icon(window, emoji="📝", size=64):
    """Sets an emoji as the window icon."""
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    font = QFont("Segoe UI Emoji")
    font.setPointSize(size - 10)
    painter.setFont(font)
    painter.drawText(pixmap.rect(), Qt.AlignmentFlag.AlignCenter, emoji)
    painter.end()
    icon = QIcon(pixmap)
    window.setWindowIcon(icon)

class ConvertWorker(QThread):
    """
    Worker thread to handle PDF to Word conversion and Word to PDF conversion
    to keep the UI responsive.
    Supports normal conversion and per-page conversion for PDF to Word.
    """
    progress = pyqtSignal(int) # Emits overall progress (files or pages)
    page_progress = pyqtSignal(int) # Emits progress specifically for per-page mode
    result = pyqtSignal(list) # Emits a list of conversion results (success/failure)

    def __init__(self, files, output_folder, mode, fmt, direction):
        super().__init__()
        self.files = files
        self.output_folder = output_folder
        self.mode = mode
        self.fmt = fmt # This will be 'docx', 'doc', or 'pdf'
        self.direction = direction # 'pdf_to_word' or 'word_to_pdf'
        self.total_steps = 0

        # Calculate total steps for the progress bar based on the conversion mode and direction
        if self.direction == 'pdf_to_word' and self.mode == 'per_page':
            for f in self.files:
                try:
                    reader = PdfReader(f)
                    self.total_steps += len(reader.pages)
                except Exception:
                    pass # Skip files that cannot be read
        else:
            self.total_steps = len(self.files) # One step per file for normal mode or Word to PDF

    def run(self):
        """Performs the conversion in the background."""
        results = []
        current_step = 0
        for i, path in enumerate(self.files):
            name = os.path.splitext(os.path.basename(path))[0]
            
            if self.direction == 'pdf_to_word':
                out_name = name + f".{self.fmt}"
                dest_path = os.path.join(self.output_folder, out_name)
                tmp_docx = "" # Initialize tmp_docx for doc conversion later

                try:
                    if self.mode == 'normal':
                        # Normal conversion: convert the whole PDF to one DOCX
                        converter = Converter(path)
                        tmp_docx = dest_path if self.fmt == 'docx' else os.path.join(self.output_folder, name + '.docx')
                        converter.convert(tmp_docx, start=0, end=None)
                        converter.close()

                        current_step += 1
                        self.progress.emit(current_step)

                    else: # 'per_page' mode (PDF to Word)
                        # Per-page conversion: convert each page to a separate DOCX, then merge
                        temp_docs = []
                        reader = PdfReader(path)
                        page_count = len(reader.pages)

                        converter = Converter(path)
                        for p in range(page_count):
                            tmp = os.path.join(tempfile.gettempdir(), f"{name}_p{p}.docx")
                            converter.convert(tmp, start=p, end=p+1)
                            temp_docs.append(tmp)
                            current_step += 1
                            self.progress.emit(current_step)
                            self.page_progress.emit(current_step) # Specific signal for page progress
                        converter.close()

                        # Merge the individual page DOCX files into one
                        merged = Document()
                        for td in temp_docs:
                            doc = Document(td)
                            for el in doc.element.body:
                                merged.element.body.append(el)
                            os.remove(td) # Clean up temporary page files
                        tmp_docx = os.path.join(self.output_folder, name + ".docx")
                        merged.save(tmp_docx)

                    # If output format is 'doc' and win32com is available, convert DOCX to DOC
                    if self.fmt == 'doc' and win32com:
                        word = win32com.client.Dispatch('Word.Application')
                        word.Visible = False
                        doc = word.Documents.Open(tmp_docx)
                        doc.SaveAs(dest_path, FileFormat=0) # FileFormat=0 saves as Word Document (.doc)
                        doc.Close()
                        word.Quit()
                        if os.path.exists(tmp_docx) and dest_path != tmp_docx:
                            os.remove(tmp_docx) # Remove the intermediate .docx file

                    results.append(f"✅ {out_name}")
                except Exception as e:
                    results.append(f"❌ {out_name} - 失敗：{e}")
                    # Ensure progress is updated even on failure for overall count
                    if self.mode == 'normal':
                        current_step += 1
                        self.progress.emit(current_step)
                    # For per_page, progress is updated per page, so no extra update here

            elif self.direction == 'word_to_pdf':
                out_name = name + ".pdf"
                dest_path = os.path.join(self.output_folder, out_name)
                
                if not win32com:
                    results.append(f"❌ {out_name} - 失敗：需要安裝 Microsoft Word 和 pywin32 才能轉換 Word 到 PDF。")
                    current_step += 1
                    self.progress.emit(current_step)
                    continue

                try:
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = False # Keep Word application hidden
                    doc = word.Documents.Open(path)
                    # FileFormat=17 is wdFormatPDF
                    doc.SaveAs(dest_path, FileFormat=17)
                    doc.Close()
                    word.Quit()
                    results.append(f"✅ {out_name}")
                except Exception as e:
                    results.append(f"❌ {out_name} - 失敗：{e}")
                finally:
                    current_step += 1
                    self.progress.emit(current_step)
                    # Ensure Word process is closed even if an error occurs
                    try:
                        word.Quit()
                    except Exception:
                        pass # Already closed or not initialized

        self.progress.emit(self.total_steps) # Ensure progress bar reaches 100%
        self.result.emit(results)

class PDFConverterApp(QWidget):
    """Main application window for the PDF to Word converter."""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("📝 PDF/Word 轉換工具")
        # Initial size, will be maximized later
        self.resize(800, 700) # Increased default size for better initial look
        self.file_paths = [] # Stores full paths of selected PDF/Word files
        self.output_folder = ""
        self.is_dark = False # Theme state
        self.convert_start_time = None
        self.convert_total_time = 0
        self.current_page_progress = 0
        self.total_pages = 0
        self.init_ui()
        self.setAcceptDrops(True) # Enable drag and drop
        self.apply_light_theme() # Apply initial theme
        set_emoji_icon(self)
        self.showMaximized() # Maximize the window on startup

    def init_ui(self):
        """Initializes the user interface."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30) # Increased margins
        layout.setSpacing(15) # Increased spacing between widgets

        # Conversion Direction selection
        layout.addWidget(QLabel("選擇轉換方向:"))
        self.direction_group = QButtonGroup(self)
        self.pdf_to_word_rb = QRadioButton("PDF 轉 Word")
        self.word_to_pdf_rb = QRadioButton("Word 轉 PDF")
        self.pdf_to_word_rb.setChecked(True) # Default direction
        self.direction_group.addButton(self.pdf_to_word_rb)
        self.direction_group.addButton(self.word_to_pdf_rb)
        
        direction_layout = QHBoxLayout()
        direction_layout.addWidget(self.pdf_to_word_rb)
        direction_layout.addWidget(self.word_to_pdf_rb)
        direction_layout.addStretch(1) # Push radio buttons to left
        layout.addLayout(direction_layout)

        # Connect direction change to update UI
        self.pdf_to_word_rb.toggled.connect(self.update_ui_for_direction)

        # ==================== FIX START ====================
        # Group PDF-to-Word specific widgets into a list to enable/disable them together.
        # This is a robust way to avoid layout issues caused by hiding/showing widgets.
        self.pdf_mode_widgets = []

        # Conversion Mode selection (Only for PDF to Word)
        mode_label = QLabel("選擇轉換模式 (僅限 PDF 轉 Word):")
        layout.addWidget(mode_label)
        self.pdf_mode_widgets.append(mode_label)
        
        self.mode_group = QButtonGroup(self)
        self.normal_rb = QRadioButton("正常模式 (將 PDF 轉換為單一 Word 檔案)")
        self.perpage_rb = QRadioButton("逐頁模式 (將 PDF 逐頁轉換後合併為單一 Word 檔案)")
        self.normal_rb.setChecked(True)
        self.mode_group.addButton(self.normal_rb)
        self.mode_group.addButton(self.perpage_rb)
        
        layout.addWidget(self.normal_rb)
        layout.addWidget(self.perpage_rb)
        self.pdf_mode_widgets.append(self.normal_rb)
        self.pdf_mode_widgets.append(self.perpage_rb)
        # ===================== FIX END =====================
        
        # Output Format selection (Horizontal layout)
        output_format_layout = QHBoxLayout()
        self.output_format_label = QLabel("選擇輸出格式:")
        output_format_layout.addWidget(self.output_format_label)
        self.format_combo = QComboBox()
        self.format_combo.addItems(["docx", "doc"]) # Default for PDF to Word
        output_format_layout.addWidget(self.format_combo)
        output_format_layout.addStretch(1) # Push combo box to the left
        layout.addLayout(output_format_layout)

        # File selection and output folder buttons
        file_folder_btns = QHBoxLayout()
        self.select_file_btn = QPushButton("📂 選擇檔案")
        self.select_file_btn.clicked.connect(self.select_files)
        file_folder_btns.addWidget(self.select_file_btn)

        self.select_output_btn = QPushButton("📁 選擇輸出資料夾")
        self.select_output_btn.clicked.connect(self.select_output_folder)
        file_folder_btns.addWidget(self.select_output_btn)
        layout.addLayout(file_folder_btns)

        # List to display selected files
        layout.addWidget(QLabel("已選取的檔案:"))
        self.file_list_widget = QListWidget()
        self.file_list_widget.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection) # Allow multiple selections
        self.file_list_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding) # Allow it to expand
        layout.addWidget(self.file_list_widget, 1) # Give it a stretch factor of 1 to expand vertically

        # Buttons for managing selected files
        file_management_btns = QHBoxLayout()
        self.remove_selected_btn = QPushButton("🗑️ 移除選取檔案")
        self.remove_selected_btn.clicked.connect(self.remove_selected_files)
        file_management_btns.addWidget(self.remove_selected_btn)

        self.clear_all_btn = QPushButton("🧹 清除所有檔案")
        self.clear_all_btn.clicked.connect(self.clear_all_files)
        file_management_btns.addWidget(self.clear_all_btn)
        layout.addLayout(file_management_btns)

        # Selected Output Folder Display (Horizontal layout)
        output_folder_display_layout = QHBoxLayout()
        output_folder_display_layout.addWidget(QLabel("已選取的輸出資料夾:"))
        self.selected_folder_label = QLabel("無") # Displays the selected folder path
        self.selected_folder_label.setWordWrap(True) # Allow text to wrap
        self.selected_folder_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred) # Allow it to expand horizontally
        output_folder_display_layout.addWidget(self.selected_folder_label)
        
        self.clear_folder_btn = QPushButton("🗑️ 清除") # Shorter text for button
        self.clear_folder_btn.clicked.connect(self.clear_output_folder)
        output_folder_display_layout.addWidget(self.clear_folder_btn)
        layout.addLayout(output_folder_display_layout)

        # Status and Progress Bar
        self.status = QLabel("請選擇檔案與輸出路徑")
        self.status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress = QProgressBar()
        self.progress.setFormat("%p%") # Display percentage
        layout.addWidget(self.status)
        layout.addWidget(self.progress)

        # Control buttons (Convert, Open Folder, Show Progress)
        ctrl_btns = QHBoxLayout()
        self.convert_btn = QPushButton("🚀 開始轉換")
        self.convert_btn.clicked.connect(self.convert_files)
        ctrl_btns.addWidget(self.convert_btn)

        self.open_btn = QPushButton("📂 開啟輸出資料夾")
        self.open_btn.clicked.connect(self.open_folder)
        self.open_btn.setEnabled(False) # Disabled until conversion is done
        ctrl_btns.addWidget(self.open_btn)

        self.show_progress_btn = QPushButton("📊 顯示逐頁進度")
        self.show_progress_btn.setEnabled(False) # Only enabled for per-page mode
        self.show_progress_btn.clicked.connect(self.show_progress_dialog)
        ctrl_btns.addWidget(self.show_progress_btn)

        layout.addLayout(ctrl_btns)

        # Initial UI update based on default direction
        self.update_ui_for_direction()
        self.update_file_count_status()
        self.update_output_folder_display() # Initial update for folder display

    def update_ui_for_direction(self):
        """Updates UI elements based on the selected conversion direction."""
        # Check which radio button is checked to determine the state
        is_pdf_to_word = self.pdf_to_word_rb.isChecked()
        
        # ==================== FIX START ====================
        # Enable/disable the PDF-to-Word specific widgets instead of hiding them.
        # This prevents layout corruption issues.
        for widget in self.pdf_mode_widgets:
            widget.setEnabled(is_pdf_to_word)
        # ===================== FIX END =====================

        if is_pdf_to_word:
            # PDF to Word selected
            self.output_format_label.setText("選擇輸出格式:")
            self.format_combo.clear()
            self.format_combo.addItems(["docx", "doc"])
            self.format_combo.setEnabled(True) # Enable combo box for selection
            self.show_progress_btn.setEnabled(self.perpage_rb.isChecked()) # Only if per-page is also selected
        else:
            # Word to PDF selected
            self.output_format_label.setText("輸出格式: PDF (自動)")
            self.format_combo.clear() # Clear existing items
            self.format_combo.addItem("pdf") # Add "pdf" as the only option
            self.format_combo.setEnabled(False) # Disable combo box as format is fixed
            self.show_progress_btn.setEnabled(False) # Per-page not applicable for Word to PDF

        # Clear selected files when direction changes to avoid confusion
        self.clear_all_files(show_message=False) # Clear silently

    def update_file_count_status(self):
        """Updates the status label with the current number of selected files."""
        count = len(self.file_paths)
        if count == 0:
            self.status.setText("請選擇檔案與輸出路徑")
        else:
            self.status.setText(f"已選 {count} 份檔案")

    def update_output_folder_display(self):
        """Updates the label displaying the selected output folder."""
        if self.output_folder:
            self.selected_folder_label.setText(self.output_folder)
        else:
            self.selected_folder_label.setText("無")

    def select_files(self):
        """Opens a file dialog to select files based on the current conversion direction."""
        if self.pdf_to_word_rb.isChecked():
            file_filter = "PDF Files (*.pdf)"
            dialog_title = "選擇 PDF 檔案"
        else: # word_to_pdf_rb is checked
            file_filter = "Word Files (*.docx *.doc)"
            dialog_title = "選擇 Word 檔案"

        files, _ = QFileDialog.getOpenFileNames(self, dialog_title, "", file_filter)
        if files:
            self.add_files_to_list(files)

    def add_files_to_list(self, files):
        """
        Adds new files to the internal list and the QListWidget, avoiding duplicates.
        Handles duplicate file names by appending (n) to the display name.
        """
        added_count = 0
        for f in files:
            if f not in self.file_paths:
                self.file_paths.append(f)
                
                original_file_name = os.path.basename(f)
                display_file_name = original_file_name
                
                # Check for duplicate display names in the QListWidget
                existing_display_names = set()
                for i in range(self.file_list_widget.count()):
                    existing_display_names.add(self.file_list_widget.item(i).text())

                counter = 1
                while display_file_name in existing_display_names:
                    name_parts = os.path.splitext(original_file_name)
                    display_file_name = f"{name_parts[0]}({counter}){name_parts[1]}"
                    counter += 1

                item = QListWidgetItem(display_file_name)
                item.setData(Qt.ItemDataRole.UserRole, f) # Store full path in item's data
                self.file_list_widget.addItem(item)
                added_count += 1
        self.update_file_count_status()
        if added_count > 0 and self.isVisible(): # Check if window is visible to avoid message on startup
            QMessageBox.information(self, "檔案選擇", f"已新增 {added_count} 個新檔案。")

    def remove_selected_files(self):
        """Removes selected files from the list widget and internal file_paths."""
        selected_items = self.file_list_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "提醒", "請先選取要移除的檔案。")
            return

        reply = QMessageBox.question(self, "確認移除",
                                     f"您確定要移除選取的 {len(selected_items)} 個檔案嗎？",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            for item in selected_items:
                file_path_to_remove = item.data(Qt.ItemDataRole.UserRole)
                if file_path_to_remove in self.file_paths:
                    self.file_paths.remove(file_path_to_remove)
                self.file_list_widget.takeItem(self.file_list_widget.row(item))
            self.update_file_count_status()

    def clear_all_files(self, show_message=True):
        """Clears all selected files from the list and resets the UI."""
        if not self.file_paths and show_message:
            QMessageBox.information(self, "提醒", "目前沒有選取的檔案。")
            return

        if show_message:
            reply = QMessageBox.question(self, "確認清除",
                                         "您確定要清除所有選取的檔案嗎？",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.No:
                return

        self.file_paths.clear()
        self.file_list_widget.clear()
        self.progress.setValue(0)
        self.open_btn.setEnabled(False)
        self.show_progress_btn.setEnabled(False)
        self.update_file_count_status()
        if show_message:
            QMessageBox.information(self, "清除完成", "所有檔案已清除。")

    def select_output_folder(self):
        """Opens a dialog to select the output folder."""
        folder = QFileDialog.getExistingDirectory(self, "選擇輸出資料夾")
        if folder:
            self.output_folder = folder
            self.update_output_folder_display() # Update the new label

    def clear_output_folder(self):
        """Clears the selected output folder."""
        if not self.output_folder:
            QMessageBox.information(self, "提醒", "目前沒有選取的輸出資料夾。")
            return

        reply = QMessageBox.question(self, "確認清除",
                                     "您確定要清除選取的輸出資料夾嗎？",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.output_folder = ""
            self.update_output_folder_display()
            QMessageBox.information(self, "清除完成", "輸出資料夾已清除。")


    def convert_files(self):
        """Initiates the conversion process."""
        if not self.file_paths:
            QMessageBox.warning(self, "提醒", "請先選擇至少一個檔案")
            return
        if not os.path.isdir(self.output_folder):
            QMessageBox.warning(self, "提醒", "請先選擇一個有效的輸出資料夾")
            return

        direction = 'pdf_to_word' if self.pdf_to_word_rb.isChecked() else 'word_to_pdf'
        mode = 'normal' # Default for Word to PDF, or for PDF to Word if not per-page
        fmt = '' # Output format, will be determined by direction

        if direction == 'pdf_to_word':
            mode = 'normal' if self.normal_rb.isChecked() else 'per_page'
            fmt = self.format_combo.currentText()
            # Calculate total pages for per-page mode for progress bar maximum
            if mode == 'per_page':
                total_pages_for_progress = 0
                for f in self.file_paths:
                    try:
                        reader = PdfReader(f)
                        total_pages_for_progress += len(reader.pages)
                    except Exception:
                        pass
                self.total_pages = total_pages_for_progress
                self.progress.setMaximum(total_pages_for_progress)
                self.show_progress_btn.setEnabled(True)
            else:
                self.total_pages = 0 # Reset for normal mode
                self.progress.setMaximum(len(self.file_paths))
                self.show_progress_btn.setEnabled(False)
        else: # word_to_pdf
            fmt = 'pdf' # Output is always PDF
            self.total_pages = 0 # Per-page not applicable
            self.progress.setMaximum(len(self.file_paths))
            self.show_progress_btn.setEnabled(False) # Per-page not applicable

        self.progress.setValue(0)
        self.open_btn.setEnabled(False)
        self.status.setText("🔄 正在轉換中... 請稍候")

        self.convert_start_time = time.time()
        self.current_page_progress = 0 # Reset page progress for new conversion

        self.worker = ConvertWorker(self.file_paths, self.output_folder, mode, fmt, direction)

        # Connect progress signals based on mode and direction
        if direction == 'pdf_to_word' and mode == 'per_page':
            self.worker.page_progress.connect(self.progress.setValue)
        else:
            self.worker.progress.connect(self.progress.setValue)

        # Connect a shared signal to update current_page_progress for the dialog
        self.worker.progress.connect(self.update_current_page_progress)
        self.worker.page_progress.connect(self.update_current_page_progress)

        self.worker.result.connect(self.convert_finished)
        self.worker.start()

    def update_current_page_progress(self, val):
        """Updates the internal variable tracking current page progress."""
        self.current_page_progress = val

    def show_progress_dialog(self):
        """Shows a message box with the current per-page conversion progress."""
        if self.total_pages > 0:
            QMessageBox.information(self, "逐頁轉換進度",
                                    f"已轉換 {self.current_page_progress} 頁 / 總共 {self.total_pages} 頁")
        else:
            QMessageBox.information(self, "逐頁轉換進度", "目前非逐頁模式，或尚未開始轉換。")

    def convert_finished(self, results):
        """Called when the conversion worker thread finishes."""
        if self.convert_start_time is not None:
            elapsed = time.time() - self.convert_start_time
            self.convert_total_time += elapsed
            self.convert_start_time = None

        self.status.setText("✅ 所有轉換已完成！")
        QMessageBox.information(self, "轉換結果", "\n".join(results))
        self.open_btn.setEnabled(True)
        self.progress.setValue(self.progress.maximum()) # Ensure progress bar is full

    def open_folder(self):
        """Opens the output folder in the system's file explorer."""
        if os.path.isdir(self.output_folder):
            os.startfile(self.output_folder)
        else:
            QMessageBox.warning(self, "錯誤", "輸出資料夾不存在。")

    def dragEnterEvent(self, event):
        """Handles drag enter event for drag and drop."""
        if event.mimeData().hasUrls():
            # Accept drag if it contains files
            event.acceptProposedAction()

    def dropEvent(self, event):
        """Handles drop event for drag and drop, adding files based on current direction."""
        files_to_add = []
        current_direction = 'pdf_to_word' if self.pdf_to_word_rb.isChecked() else 'word_to_pdf'
        
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if current_direction == 'pdf_to_word' and path.lower().endswith(".pdf"):
                files_to_add.append(path)
            elif current_direction == 'word_to_pdf' and (path.lower().endswith(".docx") or path.lower().endswith(".doc")):
                files_to_add.append(path)
        
        if files_to_add:
            self.add_files_to_list(files_to_add)
        event.acceptProposedAction()

    def apply_light_theme(self):
        """Applies the light theme stylesheet."""
        self.setStyleSheet("""
            QWidget {
                background-color: #f7f9fc;
                color: #000000;
                font-family: 'Microsoft JhengHei';
                font-size: 16px; /* Increased font size for better readability on larger screens */
            }
            QLabel {
                color: #333333;
                margin-top: 5px;
                margin-bottom: 2px;
                font-weight: bold;
            }
            QLabel:disabled {
                color: #aaaaaa;
            }
            QPushButton {
                border-radius: 20px; /* Larger radius for buttons */
                padding: 10px 20px; /* Larger padding for buttons */
                font-weight: bold;
                border: none;
                background-color: #4a90e2;
                color: white;
                min-height: 40px; /* Larger consistent button height */
            }
            QPushButton:hover {
                background-color: #357ABD;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
            QProgressBar {
                border-radius: 12px; /* Smoother/more rounded progress bar */
                height: 24px;
                text-align: center;
                font-weight: bold;
                border: 1px solid #bbb;
                background-color: #e0e0e0;
                color: #000000;
            }
            QProgressBar::chunk {
                border-radius: 12px; /* Smoother/more rounded progress bar chunk */
                background-color: #4CAF50;
            }
            QComboBox {
                border: 2px solid #4a90e2;
                border-radius: 15px;
                padding: 5px 15px;
                min-width: 100px;
                background-color: #f0f6ff;
                color: #2a2a2a;
                min-height: 35px; /* Adjusted combo box height for consistency with larger buttons */
            }
            QComboBox:hover {
                border-color: #357ABD;
                background-color: #e6f0ff;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 30px;
                border-left-width: 1px;
                border-left-color: #4a90e2;
                border-left-style: solid;
                border-top-right-radius: 15px;
                border-bottom-right-radius: 15px;
                background-color: #4a90e2;
            }
            QComboBox::down-arrow {
                /* Updated SVG with blue fill color */
                image: url(data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTIiIGhlaWdodD0iOCIgdmlld0JveD0iMCAwIDEyIDgiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+PHBhdGggZD0iTTIgMmg4bC00IDQiIGZpbGw9IiM0YTkwZTIiLz48L3N2Zz4=);
                width: 12px;
                height: 8px;
            }
            QComboBox QAbstractItemView { /* Styling for the dropdown list itself */
                border: 2px solid #4a90e2;
                border-radius: 10px; /* Rounded corners for the dropdown list */
                background-color: #f0f6ff;
                selection-background-color: #357ABD;
                color: #2a2a2a;
                padding: 5px;
            }
            QRadioButton {
                font-weight: bold;
                color: #333333;
                margin-bottom: 5px;
            }
            QRadioButton:disabled {
                color: #aaaaaa;
            }
            QListWidget {
                border: 2px solid #4a90e2;
                border-radius: 10px;
                background-color: #ffffff;
                padding: 5px;
                min-height: 100px; /* Keep a minimum height */
            }
            QListWidget::item {
                padding: 5px;
                margin-bottom: 2px;
                border-radius: 5px;
            }
            QListWidget::item:selected {
                background-color: #e0f0ff;
                color: #000000;
            }
            /* Scrollbar styling */
            QScrollBar:vertical {
                border: none;
                background: #f0f0f0; /* Light background for the track */
                width: 12px; /* Width of the vertical scrollbar */
                margin: 0px 0px 0px 0px;
                border-radius: 6px; /* Rounded track */
            }
            QScrollBar::handle:vertical {
                background: #a0a0a0; /* Color of the scrollbar handle */
                border-radius: 6px; /* Rounded handle */
                min-height: 20px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                background: none; /* No buttons/arrows */
                height: 0px;
                width: 0px;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none; /* No page area background */
            }
            QScrollBar:horizontal {
                border: none;
                background: #f0f0f0; /* Light background for the track */
                height: 12px; /* Height of the horizontal scrollbar */
                margin: 0px 0px 0px 0px;
                border-radius: 6px; /* Rounded track */
            }
            QScrollBar::handle:horizontal {
                background: #a0a0a0; /* Color of the scrollbar handle */
                border-radius: 6px; /* Rounded handle */
                min-width: 20px;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                background: none; /* No buttons/arrows */
                height: 0px;
                width: 0px;
            }
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background: none; /* No page area background */
            }
        """)

    def apply_dark_theme(self):
        """Applies the dark theme stylesheet."""
        self.setStyleSheet("""
            QWidget {
                background-color: #1e1e1e;
                color: #eeeeee;
                font-family: 'Microsoft JhengHei';
                font-size: 16px; /* Increased font size for better readability on larger screens */
            }
            QLabel {
                color: #dddddd;
                margin-top: 5px;
                margin-bottom: 2px;
                font-weight: bold;
            }
            QLabel:disabled {
                color: #777777;
            }
            QPushButton {
                border-radius: 20px; /* Larger radius for buttons */
                padding: 10px 20px; /* Larger padding for buttons */
                font-weight: bold;
                border: none;
                background-color: #3c3f41;
                color: white;
                min-height: 40px; /* Larger consistent button height */
            }
            QPushButton:hover {
                background-color: #505357;
            }
            QPushButton:disabled {
                background-color: #2a2a2a;
                color: #777777;
            }
            QProgressBar {
                border-radius: 12px; /* Smoother/more rounded progress bar */
                height: 24px;
                text-align: center;
                font-weight: bold;
                border: 1px solid #444;
                background-color: #222;
                color: #eeeeee;
            }
            QProgressBar::chunk {
                border-radius: 12px; /* Smoother/more rounded progress bar chunk */
                background-color: #29b6f6;
            }
            QComboBox {
                border: 2px solid #555;
                border-radius: 15px;
                padding: 5px 15px;
                min-width: 100px;
                background-color: #2a2a2a;
                color: #dddddd;
                min-height: 35px; /* Adjusted combo box height for consistency with larger buttons */
            }
            QComboBox:hover {
                border-color: #29b6f6;
                background-color: #1a1a1a;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 30px;
                border-left-width: 1px;
                border-left-color: #29b6f6;
                border-left-style: solid;
                border-top-right-radius: 15px;
                border-bottom-right-radius: 15px;
                background-color: #29b6f6;
            }
            QComboBox::down-arrow {
                /* Updated SVG with blue fill color */
                image: url(data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTIiIGhlaWdodD0iOCIgdmlld0JveD0iMCAwIDEyIDgiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+PHBhdGggZD0iTTIgMmg4bC00IDQiIGZpbGw9IiMyOWI2ZjYiLz48L3N2Zz4=);
                width: 12px;
                height: 8px;
            }
            QComboBox QAbstractItemView { /* Styling for the dropdown list itself */
                border: 2px solid #29b6f6;
                border-radius: 10px; /* Rounded corners for the dropdown list */
                background-color: #1a1a1a;
                selection-background-color: #29b6f6;
                color: #dddddd;
                padding: 5px;
            }
            QRadioButton {
                font-weight: bold;
                color: #eeeeee;
                margin-bottom: 5px;
            }
            QRadioButton:disabled {
                color: #777777;
            }
            QListWidget {
                border: 2px solid #555;
                border-radius: 10px;
                background-color: #2a2a2a;
                padding: 5px;
                min-height: 100px; /* Keep a minimum height */
            }
            QListWidget::item {
                padding: 5px;
                margin-bottom: 2px;
                border-radius: 5px;
            }
            QListWidget::item:selected {
                background-color: #3c3f41;
                color: #ffffff;
            }
            /* Scrollbar styling */
            QScrollBar:vertical {
                border: none;
                background: #2a2a2a; /* Dark background for the track */
                width: 12px; /* Width of the vertical scrollbar */
                margin: 0px 0px 0px 0px;
                border-radius: 6px; /* Rounded track */
            }
            QScrollBar::handle:vertical {
                background: #555555; /* Color of the scrollbar handle */
                border-radius: 6px; /* Rounded handle */
                min-height: 20px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                background: none; /* No buttons/arrows */
                height: 0px;
                width: 0px;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none; /* No page area background */
            }
            QScrollBar:horizontal {
                border: none;
                background: #2a2a2a; /* Dark background for the track */
                height: 12px; /* Height of the horizontal scrollbar */
                margin: 0px 0px 0px 0px;
                border-radius: 6px; /* Rounded track */
            }
            QScrollBar::handle:horizontal {
                background: #555555; /* Color of the scrollbar handle */
                border-radius: 6px; /* Rounded handle */
                min-width: 20px;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                background: none; /* No buttons/arrows */
                height: 0px;
                width: 0px;
            }
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background: none; /* No page area background */
            }
        """)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = PDFConverterApp()
    win.show()
    sys.exit(app.exec())