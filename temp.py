import sys
import os
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QPushButton,
    QLabel,
    QVBoxLayout,
    QHBoxLayout,
    QFileDialog,
    QWidget,
    QTextEdit,
    QProgressBar,
    QListWidget,
    QComboBox,
    QGroupBox,
    QCheckBox,
    QMessageBox,
    QFrame,
    QSplitter,
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QPalette, QColor

import pymupdf as fitz


class PDFExtractorThread(QThread):
    progress_update = pyqtSignal(int)
    extraction_complete = pyqtSignal(str, list)
    extraction_error = pyqtSignal(str)

    def __init__(self, pdf_paths, output_path, extract_text=True, extract_images=False):
        super().__init__()
        self.pdf_paths = pdf_paths
        self.output_path = output_path
        self.extract_text = extract_text
        self.extract_images = extract_images

    def run(self):
        processed_files = []
        try:
            total_pdfs = len(self.pdf_paths)
            for pdf_index, pdf_path in enumerate(self.pdf_paths):
                doc = fitz.open(pdf_path)
                total_pages = len(doc)
                extracted_text = ""

                base_name = os.path.splitext(os.path.basename(pdf_path))[0]

                if self.extract_images:
                    images_dir = os.path.join(self.output_path, "images")
                    os.makedirs(images_dir, exist_ok=True)

                for page_num, page in enumerate(doc):
                    if self.extract_text:
                        extracted_text += f"--- {base_name} - Page {page_num + 1} ---\n"
                        extracted_text += page.get_text()
                        extracted_text += "\n\n"

                    if self.extract_images:
                        image_list = page.get_images(full=True)
                        for img_index, img_info in enumerate(image_list):
                            xref = img_info[0]
                            base_image = doc.extract_image(xref)
                            image_bytes = base_image["image"]
                            image_ext = base_image["ext"]

                            image_filename = f"{base_name}_page{page_num + 1}_img{img_index + 1}.{image_ext}"
                            with open(
                                os.path.join(images_dir, image_filename), "wb"
                            ) as img_file:
                                img_file.write(image_bytes)

                    progress = int(
                        ((pdf_index + (page_num + 1) / total_pages) / total_pdfs) * 100
                    )
                    self.progress_update.emit(progress)

                if self.extract_text:
                    text_file_path = os.path.join(
                        self.output_path, f"{base_name}_text.txt"
                    )
                    with open(text_file_path, "w", encoding="utf-8") as text_file:
                        text_file.write(extracted_text)

                processed_files.append(base_name + ".pdf")

            if total_pdfs > 1:
                self.extraction_complete.emit("All PDFs processed.", processed_files)
            else:
                self.extraction_complete.emit(extracted_text, processed_files)

        except Exception as e:
            self.extraction_error.emit(f"Error extracting PDFs: {str(e)}")


class PDFExtractorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Extractor")
        self.setMinimumSize(900, 600)
        self.set_dark_mode()

        self.pdf_paths = []
        self.output_dir = None

        self.init_ui()

    def set_dark_mode(self):
        dark_palette = QPalette()
        dark_palette.setColor(QPalette.Window, QColor(53, 53, 53))
        dark_palette.setColor(QPalette.WindowText, Qt.white)
        dark_palette.setColor(QPalette.Base, QColor(35, 35, 35))
        dark_palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
        dark_palette.setColor(QPalette.ToolTipBase, Qt.white)
        dark_palette.setColor(QPalette.ToolTipText, Qt.white)
        dark_palette.setColor(QPalette.Text, Qt.white)
        dark_palette.setColor(QPalette.Button, QColor(53, 53, 53))
        dark_palette.setColor(QPalette.ButtonText, Qt.white)
        dark_palette.setColor(QPalette.BrightText, Qt.red)
        dark_palette.setColor(QPalette.Link, QColor(42, 130, 218))
        QApplication.setPalette(dark_palette)

    def init_ui(self):
        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)

        title_label = QLabel("PDF Extractor")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter) # Attribute "AlignCenter" is unknown error in editor
        main_layout.addWidget(title_label)

        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["Single PDF", "All PDFs in Folder"])
        self.mode_combo.currentIndexChanged.connect(self.mode_changed)
        main_layout.addWidget(self.mode_combo)

        self.file_path_label = QLabel("No file or folder selected")
        self.file_path_label.setWordWrap(True)
        self.file_select_button = QPushButton("Select PDF or Folder")
        self.file_select_button.clicked.connect(self.select_file_or_folder)

        file_layout = QHBoxLayout()
        file_layout.addWidget(self.file_path_label)
        file_layout.addWidget(self.file_select_button)
        main_layout.addLayout(file_layout)

        self.output_path_label = QLabel("No output directory selected")
        self.output_path_label.setWordWrap(True)
        output_select_button = QPushButton("Select Output Directory")
        output_select_button.clicked.connect(self.select_output_directory)

        output_layout = QHBoxLayout()
        output_layout.addWidget(self.output_path_label)
        output_layout.addWidget(output_select_button)
        main_layout.addLayout(output_layout)

        options_group = QGroupBox("Extraction Options")
        options_layout = QVBoxLayout()

        self.extract_text_checkbox = QCheckBox("Extract Text")
        self.extract_text_checkbox.setChecked(True)
        options_layout.addWidget(self.extract_text_checkbox)

        self.extract_images_checkbox = QCheckBox("Extract Images")
        options_layout.addWidget(self.extract_images_checkbox)

        options_group.setLayout(options_layout)
        main_layout.addWidget(options_group)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        main_layout.addWidget(self.progress_bar)

        self.extract_button = QPushButton("Extract PDF Contents")
        self.extract_button.setEnabled(False)
        self.extract_button.clicked.connect(self.start_extraction)
        main_layout.addWidget(self.extract_button)

        self.processed_list = QListWidget()
        main_layout.addWidget(QLabel("Processed Files:"))
        main_layout.addWidget(self.processed_list)

        self.preview_toggle_button = QPushButton("Show/Hide Extracted Text Preview")
        self.preview_toggle_button.setCheckable(True)
        self.preview_toggle_button.setChecked(True)
        self.preview_toggle_button.clicked.connect(self.toggle_preview)
        main_layout.addWidget(self.preview_toggle_button)

        self.preview_group = QGroupBox("Extracted Text Preview")
        preview_layout = QVBoxLayout()

        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setMinimumHeight(900)
        # self.preview_text.setMaximumHeight(300)
        preview_layout.addWidget(self.preview_text)

        self.preview_group.setLayout(preview_layout)
        main_layout.addWidget(self.preview_group)

        self.setCentralWidget(main_widget)

    def toggle_preview(self):
        self.preview_group.setVisible(self.preview_toggle_button.isChecked())

    def mode_changed(self):
        self.file_path_label.setText("No file or folder selected")
        self.pdf_paths = []
        self.check_extract_button()

    def select_file_or_folder(self):
        if self.mode_combo.currentIndex() == 0:  # Single PDF mode
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Select PDF File", "", "PDF Files (*.pdf)"
            )
            if file_path:
                self.pdf_paths = [file_path]
                self.file_path_label.setText(file_path)
        else:  # All PDFs in Folder mode (with subfolder traversal)
            folder_path = QFileDialog.getExistingDirectory(
                self, "Select Folder Containing PDFs"
            )
            if folder_path:
                self.pdf_paths = []
                # Recursively walk through all subdirectories
                for root, dirs, files in os.walk(folder_path):
                    for file in files:
                        if file.lower().endswith(".pdf"):
                            full_path = os.path.join(root, file)
                            self.pdf_paths.append(full_path)

                if not self.pdf_paths:
                    self.file_path_label.setText(
                        "No PDF files found in selected folder or subfolders"
                    )
                else:
                    # Warn if many PDFs found
                    if len(self.pdf_paths) > 200:
                        reply = QMessageBox.question(
                            self,
                            "Many PDFs Found",
                            f"Found {len(self.pdf_paths)} PDF files. Processing may take a while. Continue?",
                            QMessageBox.Yes | QMessageBox.No,
                            QMessageBox.No,
                        )
                        if reply == QMessageBox.No:
                            self.pdf_paths = []
                            self.file_path_label.setText("Operation canceled")
                            self.check_extract_button()
                            return

                    self.file_path_label.setText(
                        f"{len(self.pdf_paths)} PDF file(s) found in folder and subfolders"
                    )

        self.check_extract_button()

    def select_output_directory(self):
        output_dir = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if output_dir:
            self.output_dir = output_dir
            self.output_path_label.setText(output_dir)
            self.check_extract_button()

    def check_extract_button(self):
        self.extract_button.setEnabled(
            bool(self.pdf_paths) and self.output_dir is not None
        )

    def start_extraction(self):
        # Automatically hide preview in batch mode
        if self.mode_combo.currentText() == "All PDFs in Folder":
            self.preview_toggle_button.setChecked(False)
            self.preview_group.setVisible(False)
        else:
            self.preview_toggle_button.setChecked(True)
            self.preview_group.setVisible(True)
        if (
            not self.extract_text_checkbox.isChecked()
            and not self.extract_images_checkbox.isChecked()
        ):
            QMessageBox.warning(
                self, "Warning", "Please select at least one extraction option."
            )
            return

        self.extract_button.setEnabled(False)
        self.preview_text.clear()
        self.progress_bar.setValue(0)
        self.processed_list.clear()

        self.extraction_thread = PDFExtractorThread(
            self.pdf_paths,
            self.output_dir,
            self.extract_text_checkbox.isChecked(),
            self.extract_images_checkbox.isChecked(),
        )

        self.extraction_thread.progress_update.connect(self.update_progress)
        self.extraction_thread.extraction_complete.connect(self.on_extraction_complete)
        self.extraction_thread.extraction_error.connect(self.on_extraction_error)

        self.extraction_thread.start()

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def on_extraction_complete(self, extracted_text, processed_files):
        if (
            self.extract_text_checkbox.isChecked()
            and self.mode_combo.currentText() == "Single PDF"
        ):
            self.preview_text.setText(extracted_text)

        self.processed_list.addItems(processed_files)
        self.extract_button.setEnabled(True)

        QMessageBox.information(
            self,
            "Extraction Complete",
            f"PDF extraction completed successfully.\n\nFiles saved to: {self.output_dir}",
        )

    def on_extraction_error(self, error_message):
        self.extract_button.setEnabled(True)
        QMessageBox.critical(self, "Error", error_message)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFExtractorApp()
    window.show()
    sys.exit(app.exec_())
