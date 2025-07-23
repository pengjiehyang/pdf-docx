import sys, os, time, tempfile
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel,
    QProgressBar, QMessageBox, QHBoxLayout, QRadioButton,
    QButtonGroup, QComboBox
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
    progress = pyqtSignal(int)
    page_progress = pyqtSignal(int)
    result = pyqtSignal(list)

    def __init__(self, files, output_folder, mode, fmt):
        super().__init__()
        self.files = files
        self.output_folder = output_folder
        self.mode = mode
        self.fmt = fmt

        if self.mode == 'per_page':
            total_pages = 0
            for f in self.files:
                try:
                    reader = PdfReader(f)
                    total_pages += len(reader.pages)
                except Exception:
                    pass
            self.total_steps = total_pages
        else:
            self.total_steps = len(self.files)

    def run(self):
        results = []
        current_step = 0
        for i, path in enumerate(self.files):
            name = os.path.splitext(os.path.basename(path))[0]
            out_name = name + f".{self.fmt}"
            dest_path = os.path.join(self.output_folder, out_name)
            try:
                if self.mode == 'normal':
                    converter = Converter(path)
                    tmp_docx = dest_path if self.fmt == 'docx' else dest_path + '.docx'
                    converter.convert(tmp_docx, start=0, end=None)
                    converter.close()

                    current_step += 1
                    self.progress.emit(current_step)

                else:
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
                        self.page_progress.emit(current_step)
                    converter.close()

                    merged = Document()
                    for td in temp_docs:
                        doc = Document(td)
                        for el in doc.element.body:
                            merged.element.body.append(el)
                    tmp_docx = os.path.join(self.output_folder, name + ".docx")
                    merged.save(tmp_docx)

                if self.fmt == 'doc' and win32com:
                    word = win32com.client.Dispatch('Word.Application')
                    word.Visible = False
                    doc = word.Documents.Open(tmp_docx)
                    doc.SaveAs(dest_path, FileFormat=0)
                    doc.Close()
                    word.Quit()

                results.append(f"✅ {out_name}")
            except Exception as e:
                results.append(f"❌ {out_name} - 失敗：{e}")
        self.progress.emit(self.total_steps)
        self.result.emit(results)

class PDFConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("📝 PDF 轉 Word 工具")
        self.resize(520, 380)
        self.file_paths = []
        self.output_folder = ""
        self.is_dark = False
        self.convert_start_time = None
        self.convert_total_time = 0
        self.current_page_progress = 0
        self.total_pages = 0
        self.init_ui()
        self.setAcceptDrops(True)
        self.apply_light_theme()
        set_emoji_icon(self)

    def init_ui(self):
        layout = QVBoxLayout(self)

        mode_group = QButtonGroup(self)
        self.normal_rb = QRadioButton("正常模式")
        self.perpage_rb = QRadioButton("逐頁模式")
        self.normal_rb.setChecked(True)
        mode_group.addButton(self.normal_rb)
        mode_group.addButton(self.perpage_rb)
        layout.addWidget(self.normal_rb)
        layout.addWidget(self.perpage_rb)
        self.mode_group = mode_group

        self.format_combo = QComboBox()
        self.format_combo.addItems(["docx", "doc"])
        layout.addWidget(self.format_combo)

        btns = QHBoxLayout()
        for text, slot in [("📂 選擇 PDF 檔案", self.select_pdfs),
                           ("📁 選擇輸出資料夾", self.select_output_folder)]:
            btn = QPushButton(text)
            btn.clicked.connect(slot)
            btns.addWidget(btn)
        layout.addLayout(btns)

        self.status = QLabel("請選擇檔案與輸出路徑")
        self.status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress = QProgressBar()
        layout.addWidget(self.status)
        layout.addWidget(self.progress)

        ctrl = QHBoxLayout()
        self.convert_btn = QPushButton("🚀 開始轉換")
        self.convert_btn.clicked.connect(self.convert_files)
        self.open_btn = QPushButton("📂 開啟輸出資料夾")
        self.open_btn.clicked.connect(self.open_folder)
        self.open_btn.setEnabled(False)
        ctrl.addWidget(self.convert_btn)
        ctrl.addWidget(self.open_btn)

        self.show_progress_btn = QPushButton("📊 顯示進度")
        self.show_progress_btn.setEnabled(False)
        self.show_progress_btn.clicked.connect(self.show_progress_dialog)
        ctrl.addWidget(self.show_progress_btn)

        layout.addLayout(ctrl)

        self.setStyleSheet("""
            QWidget {
                font-family: 'Microsoft JhengHei';
                font-size: 15px;
            }
            QPushButton {
                border-radius: 20px;
                padding: 10px 20px;
                font-weight: bold;
                border: none;
            }
            QPushButton:hover {
                background-color: #357ABD;
                color: white;
            }
            QProgressBar {
                border-radius: 15px;
                height: 24px;
                text-align: center;
                font-weight: bold;
                border: 1px solid #bbb;
                background-color: #f0f0f0;
                color: #000000;
            }
            QProgressBar::chunk {
                border-radius: 15px;
                background-color: #4CAF50;
            }
            QComboBox {
                border: 2px solid #4a90e2;
                border-radius: 15px;
                padding: 5px 15px;
                min-width: 100px;
                background-color: #f0f6ff;
                color: #2a2a2a;
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
                image: url(data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTIiIGhlaWdodD0iOCIgdmlld0JveD0iMCAwIDEyIDgiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+PHBhdGggZD0iTTIgMmg4bC00IDQiLz48L3N2Zz4=);
                width: 12px;
                height: 8px;
            }
            QComboBox QAbstractItemView {
                border: 2px solid #4a90e2;
                border-radius: 10px;
                background-color: #f0f6ff;
                selection-background-color: #357ABD;
                color: #2a2a2a;
                padding: 5px;
            }
            QRadioButton {
                font-weight: bold;
            }
        """)

    def select_pdfs(self):
        files, _ = QFileDialog.getOpenFileNames(self, "選擇 PDF 檔案", "", "PDF Files (*.pdf)")
        if files:
            self.file_paths.extend(f for f in files if f not in self.file_paths)
            self.status.setText(f"已選 {len(self.file_paths)} 份 PDF 檔案")

    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "選擇輸出資料夾")
        if folder:
            self.output_folder = folder
            self.status.setText(f"輸出資料夾：{folder}")

    def convert_files(self):
        if not self.file_paths:
            QMessageBox.warning(self, "提醒", "請先選擇至少一個 PDF 檔案")
            return
        if not os.path.isdir(self.output_folder):
            QMessageBox.warning(self, "提醒", "請先選擇一個有效的輸出資料夾")
            return

        mode = 'normal' if self.normal_rb.isChecked() else 'per_page'
        fmt = self.format_combo.currentText()

        if mode == 'per_page':
            total_pages = 0
            for f in self.file_paths:
                try:
                    reader = PdfReader(f)
                    total_pages += len(reader.pages)
                except Exception:
                    pass
            self.total_pages = total_pages
            self.progress.setMaximum(total_pages)
            self.show_progress_btn.setEnabled(True)
        else:
            self.total_pages = 0
            self.progress.setMaximum(len(self.file_paths))
            self.show_progress_btn.setEnabled(False)

        self.progress.setValue(0)
        self.open_btn.setEnabled(False)
        self.status.setText("🔄 正在轉換中...")

        self.convert_start_time = time.time()
        self.current_page_progress = 0

        self.worker = ConvertWorker(self.file_paths, self.output_folder, mode, fmt)

        if mode == 'per_page':
            self.worker.page_progress.connect(self.progress.setValue)
        else:
            self.worker.progress.connect(self.progress.setValue)

        self.worker.page_progress.connect(self.update_current_page_progress)
        self.worker.progress.connect(self.update_current_page_progress)

        self.worker.result.connect(self.convert_finished)
        self.worker.start()

    def update_current_page_progress(self, val):
        self.current_page_progress = val

    def show_progress_dialog(self):
        if self.total_pages > 0:
            QMessageBox.information(self, "逐頁轉換進度",
                                    f"已轉換 {self.current_page_progress} 頁 / 總共 {self.total_pages} 頁")
        else:
            QMessageBox.information(self, "逐頁轉換進度", "尚未開始或非逐頁模式")

    def convert_finished(self, results):
        if self.convert_start_time is not None:
            elapsed = time.time() - self.convert_start_time
            self.convert_total_time += elapsed
            self.convert_start_time = None

        self.status.setText("✅ 所有轉換已完成")
        QMessageBox.information(self, "轉換結果", "\n".join(results))
        self.open_btn.setEnabled(True)

    def open_folder(self):
        if os.path.isdir(self.output_folder):
            os.startfile(self.output_folder)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        added = 0
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(".pdf") and path not in self.file_paths:
                self.file_paths.append(path)
                added += 1
        self.status.setText(f"已選 {len(self.file_paths)} 份 PDF 檔案 (+{added} 個新檔案)" if added else f"已選 {len(self.file_paths)} 份 PDF 檔案")

    def apply_light_theme(self):
        self.setStyleSheet("""
            QWidget {
                background-color: #f7f9fc;
                color: #000000;
                font-family: 'Microsoft JhengHei';
                font-size: 15px;
            }
            QPushButton {
                border-radius: 20px;
                padding: 10px 20px;
                font-weight: bold;
                border: none;
                background-color: #4a90e2;
                color: white;
            }
            QPushButton:hover {
                background-color: #357ABD;
            }
            QProgressBar {
                border-radius: 15px;
                height: 24px;
                text-align: center;
                font-weight: bold;
                border: 1px solid #bbb;
                background-color: #f0f0f0;
                color: #000000;
            }
            QProgressBar::chunk {
                border-radius: 15px;
                background-color: #4CAF50;
            }
            QComboBox {
                border: 2px solid #4a90e2;
                border-radius: 15px;
                padding: 5px 15px;
                min-width: 100px;
                background-color: #f0f6ff;
                color: #2a2a2a;
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
                image: url(data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTIiIGhlaWdodD0iOCIgdmlld0JveD0iMCAwIDEyIDgiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+PHBhdGggZD0iTTIgMmg4bC00IDQiLz48L3N2Zz4=);
                width: 12px;
                height: 8px;
            }
            QComboBox QAbstractItemView {
                border: 2px solid #4a90e2;
                border-radius: 10px;
                background-color: #f0f6ff;
                selection-background-color: #357ABD;
                color: #2a2a2a;
                padding: 5px;
            }
            QRadioButton {
                font-weight: bold;
            }
        """)

    def apply_dark_theme(self):
        self.setStyleSheet("""
            QWidget {
                background-color: #1e1e1e;
                color: #eeeeee;
                font-family: 'Microsoft JhengHei';
                font-size: 15px;
            }
            QPushButton {
                border-radius: 20px;
                padding: 10px 20px;
                font-weight: bold;
                border: none;
                background-color: #3c3f41;
                color: white;
            }
            QPushButton:hover {
                background-color: #505357;
            }
            QProgressBar {
                border-radius: 15px;
                height: 24px;
                text-align: center;
                font-weight: bold;
                border: 1px solid #444;
                background-color: #222;
                color: #eeeeee;
            }
            QProgressBar::chunk {
                border-radius: 15px;
                background-color: #29b6f6;
            }
            QComboBox {
                border: 2px solid #555;
                border-radius: 15px;
                padding: 5px 15px;
                min-width: 100px;
                background-color: #2a2a2a;
                color: #dddddd;
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
                image: url(data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTIiIGhlaWdodD0iOCIgdmlld0JveD0iMCAwIDEyIDgiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+PHBhdGggZD0iTTIgMmg4bC00IDQiLz48L3N2Zz4=);
                width: 12px;
                height: 8px;
            }
            QComboBox QAbstractItemView {
                border: 2px solid #29b6f6;
                border-radius: 10px;
                background-color: #1a1a1a;
                selection-background-color: #29b6f6;
                color: #dddddd;
                padding: 5px;
            }
            QRadioButton {
                font-weight: bold;
            }
        """)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = PDFConverterApp()
    win.show()
    sys.exit(app.exec())
