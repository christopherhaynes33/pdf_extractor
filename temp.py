import sys
import os
import csv
import json
import re
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
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QPalette, QColor

import pymupdf as fitz


class FieldExtractor:
    def __init__(self, config_path=None):
        self.fields = []
        self.config_path = config_path or "pdf_extractor_config.json"
        self.load_config()

    def load_config(self):
        """Load field configuration from JSON file"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, "r") as f:
                    config = json.load(f)
                    self.fields = config.get("fields", [])
            else:
                # Create default config if doesn't exist
                self.fields = [
                    {
                        "name": "invoice_number",
                        "pattern": r"Invoice\s*Number:\s*(\w+)",
                        "required": True,
                    },
                    {
                        "name": "date",
                        "pattern": r"Date:\s*(\d{2}/\d{2}/\d{4})",
                        "required": False,
                    },
                    {
                        "name": "total_amount",
                        "pattern": r"Total\s*Amount:\s*\$?([\d,]+\.\d{2})",
                        "required": True,
                    },
                ]
                self.save_config()
        except Exception as e:
            raise Exception(f"Error loading config: {str(e)}")

    def save_config(self):
        """Save current field configuration to JSON file"""
        try:
            with open(self.config_path, "w") as f:
                json.dump({"fields": self.fields}, f, indent=4)
        except Exception as e:
            raise Exception(f"Error saving config: {str(e)}")

    def extract_fields(self, text):
        """Extract configured fields from text using regex patterns"""
        results = {}
        missing_required = []

        for field in self.fields:
            try:
                pattern = field.get("pattern", "")
                matches = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
                if matches:
                    print(matches)  # NOTE: Testing purposes
                    # Use the first capture group if exists, else the whole match
                    results[field["name"]] = (
                        matches.group(1) if matches.groups() else matches.group(0)
                    )
                elif field.get("required", False):
                    missing_required.append(field["name"])
            except Exception as e:
                raise Exception(f"Error processing field '{field['name']}': {str(e)}")

        if missing_required:
            raise Exception(f"Missing required fields: {', '.join(missing_required)}")

        return results


class PDFExtractorThread(QThread):
    progress_update = pyqtSignal(int)
    extraction_complete = pyqtSignal(str, list)
    extraction_error = pyqtSignal(str)

    def __init__(
        self,
        pdf_paths,
        output_path,
        extract_text=True,
        extract_images=False,
        export_csv=False,
    ):
        super().__init__()
        self.pdf_paths = pdf_paths
        self.output_path = output_path
        self.extract_text = extract_text
        self.extract_images = extract_images
        self.export_csv = export_csv
        self.field_extractor = FieldExtractor()

    def run(self):
        processed_files = []
        try:
            total_pdfs = len(self.pdf_paths)
            for pdf_index, pdf_path in enumerate(self.pdf_paths):
                doc = fitz.open(pdf_path)
                total_pages = len(doc)
                extracted_text = ""
                csv_data = []

                base_name = os.path.splitext(os.path.basename(pdf_path))[0]

                if self.extract_images:
                    images_dir = os.path.join(self.output_path, "images")
                    os.makedirs(images_dir, exist_ok=True)

                for page_num, page in enumerate(doc):
                    page_text = page.get_text()

                    if self.extract_text or self.export_csv:
                        extracted_text += f"--- {base_name} - Page {page_num + 1} ---\n"
                        extracted_text += page_text
                        extracted_text += "\n\n"

                    if self.export_csv:
                        try:
                            fields = self.field_extractor.extract_fields(page_text)
                            row = {"File": base_name, "Page": page_num + 1}
                            row.update(fields)
                            csv_data.append(row)
                        except Exception as e:
                            # Skip page if field extraction fails but continue processing
                            print(f"Warning processing page {page_num + 1}: {str(e)}")
                            continue

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

                if self.export_csv and csv_data:
                    csv_file_path = os.path.join(
                        self.output_path, f"{base_name}_data.csv"
                    )
                    # Get all possible field names from all rows
                    fieldnames = set()
                    for row in csv_data:
                        fieldnames.update(row.keys())
                    fieldnames = sorted(fieldnames)

                    with open(
                        csv_file_path, "w", newline="", encoding="utf-8"
                    ) as csv_file:
                        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
                        writer.writeheader()
                        writer.writerows(csv_data)

                processed_files.append(base_name + ".pdf")

            if total_pdfs > 1:
                self.extraction_complete.emit("All PDFs processed.", processed_files)
            else:
                self.extraction_complete.emit(extracted_text, processed_files)

        except Exception as e:
            self.extraction_error.emit(f"Error extracting PDFs: {str(e)}")


class ConfigEditorDialog(QWidget):
    def __init__(self, field_extractor):
        super().__init__()
        self.field_extractor = field_extractor
        self.setWindowTitle("Field Configuration Editor")
        self.setMinimumSize(600, 400)

        layout = QVBoxLayout()

        # Field list
        self.field_list = QListWidget()
        self.refresh_field_list()
        layout.addWidget(self.field_list)

        # Field details
        self.name_edit = QTextEdit()
        self.name_edit.setMaximumHeight(30)
        self.pattern_edit = QTextEdit()
        self.required_checkbox = QCheckBox("Required Field")

        details_group = QGroupBox("Field Details")
        details_layout = QVBoxLayout()
        details_layout.addWidget(QLabel("Field Name:"))
        details_layout.addWidget(self.name_edit)
        details_layout.addWidget(QLabel("Regex Pattern:"))
        details_layout.addWidget(self.pattern_edit)
        details_layout.addWidget(self.required_checkbox)
        details_group.setLayout(details_layout)
        layout.addWidget(details_group)

        # Buttons
        button_layout = QHBoxLayout()
        self.add_button = QPushButton("Add Field")
        self.add_button.clicked.connect(self.add_field)
        self.update_button = QPushButton("Update Field")
        self.update_button.clicked.connect(self.update_field)
        self.remove_button = QPushButton("Remove Field")
        self.remove_button.clicked.connect(self.remove_field)
        self.save_button = QPushButton("Save Configuration")
        self.save_button.clicked.connect(self.save_config)

        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.update_button)
        button_layout.addWidget(self.remove_button)
        button_layout.addWidget(self.save_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

        # Connect selection change
        self.field_list.currentItemChanged.connect(self.field_selected)

    def refresh_field_list(self):
        self.field_list.clear()
        for field in self.field_extractor.fields:
            self.field_list.addItem(
                f"{field['name']} ({'required' if field.get('required', False) else 'optional'})"
            )

    def field_selected(self, current):
        if current:
            index = self.field_list.row(current)
            field = self.field_extractor.fields[index]
            self.name_edit.setPlainText(field["name"])
            self.pattern_edit.setPlainText(field["pattern"])
            self.required_checkbox.setChecked(field.get("required", False))

    def add_field(self):
        name = self.name_edit.toPlainText().strip()
        pattern = self.pattern_edit.toPlainText().strip()
        required = self.required_checkbox.isChecked()

        if not name or not pattern:
            QMessageBox.warning(self, "Warning", "Field name and pattern are required.")
            return

        self.field_extractor.fields.append(
            {"name": name, "pattern": pattern, "required": required}
        )
        self.refresh_field_list()

    def update_field(self):
        current = self.field_list.currentItem()
        if not current:
            return

        index = self.field_list.row(current)
        name = self.name_edit.toPlainText().strip()
        pattern = self.pattern_edit.toPlainText().strip()
        required = self.required_checkbox.isChecked()

        if not name or not pattern:
            QMessageBox.warning(self, "Warning", "Field name and pattern are required.")
            return

        self.field_extractor.fields[index] = {
            "name": name,
            "pattern": pattern,
            "required": required,
        }
        self.refresh_field_list()

    def remove_field(self):
        current = self.field_list.currentItem()
        if not current:
            return

        index = self.field_list.row(current)
        self.field_extractor.fields.pop(index)
        self.refresh_field_list()

    def save_config(self):
        try:
            self.field_extractor.save_config()
            QMessageBox.information(
                self, "Success", "Configuration saved successfully."
            )
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save config: {str(e)}")


class PDFExtractorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Extractor")
        self.setMinimumSize(900, 600)
        self.set_dark_mode()

        self.pdf_paths = []
        self.output_dir = None
        self.field_extractor = FieldExtractor()

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
        title_label.setAlignment(Qt.AlignCenter)
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

        self.export_csv_checkbox = QCheckBox("Export as CSV (with field extraction)")
        options_layout.addWidget(self.export_csv_checkbox)

        self.extract_images_checkbox = QCheckBox("Extract Images")
        options_layout.addWidget(self.extract_images_checkbox)

        options_group.setLayout(options_layout)
        main_layout.addWidget(options_group)

        # Add configuration button
        self.config_button = QPushButton("Configure CSV Fields...")
        self.config_button.clicked.connect(self.open_config_editor)
        main_layout.addWidget(self.config_button)

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
        preview_layout.addWidget(self.preview_text)

        self.preview_group.setLayout(preview_layout)
        main_layout.addWidget(self.preview_group)

        self.setCentralWidget(main_widget)

    def open_config_editor(self):
        self.config_editor = ConfigEditorDialog(self.field_extractor)
        self.config_editor.show()

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
            and not self.export_csv_checkbox.isChecked()
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
            self.export_csv_checkbox.isChecked(),
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
