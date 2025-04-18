import sys
import os
import csv
import json
import re
from typing import List, Dict, Optional, Any, cast, Iterable, Tuple
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
    QListWidgetItem,
    QComboBox,
    QGroupBox,
    QCheckBox,
    QMessageBox,
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QPalette, QColor

import pymupdf as fitz


class FieldExtractor:
    """Handles configuration and extraction of specific fields from text using regex patterns."""

    def __init__(self, config_path: Optional[str] = None) -> None:
        """Initialize the FieldExtractor with optional custom config path.

        Args:
            config_path: Path to JSON configuration file. Defaults to 'pdf_extractor_config.json'.
        """
        self.fields: List[Dict[str, Any]] = []
        self.config_path = config_path or "pdf_extractor_config.json"
        self.load_config()

    def load_config(self) -> None:
        """Load field configuration from JSON file or create default config if none exists.

        Raises:
            Exception: If there's an error loading or parsing the config file.
        """
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, "r") as f:
                    config: Dict[str, Any] = json.load(f)
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

    def save_config(self) -> None:
        """Save current field configuration to JSON file.

        Raises:
            Exception: If there's an error writing the config file.
        """
        try:
            with open(self.config_path, "w") as f:
                json.dump({"fields": self.fields}, f, indent=4)
        except Exception as e:
            raise Exception(f"Error saving config: {str(e)}")

    def extract_fields(self, text: str) -> Dict[str, str]:
        """Extract configured fields from text using regex patterns.

        Args:
            text: The text to extract fields from.

        Returns:
            Dictionary of extracted field names and values.

        Raises:
            Exception: If required fields are missing or pattern processing fails.
        """
        results: Dict[str, str] = {}
        missing_required: List[str] = []

        for field in self.fields:
            try:
                pattern: str = field.get("pattern", "")
                matches = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
                if matches:
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
    """Worker thread for PDF extraction to prevent UI freezing."""

    progress_update = pyqtSignal(int)
    extraction_complete = pyqtSignal(str, list)
    extraction_error = pyqtSignal(str)

    def __init__(
        self,
        pdf_paths: List[str],
        output_path: str,
        extract_text: bool = True,
        extract_images: bool = False,
        export_csv: bool = False,
    ) -> None:
        """Initialize the PDF extraction thread.

        Args:
            pdf_paths: List of PDF file paths to process.
            output_path: Directory to save extracted content.
            extract_text: Whether to extract raw text.
            extract_images: Whether to extract images.
            export_csv: Whether to export data as CSV.
        """
        super().__init__()
        self.pdf_paths = pdf_paths
        self.output_path = output_path
        self.extract_text = extract_text
        self.extract_images = extract_images
        self.export_csv = export_csv
        self.field_extractor = FieldExtractor()

    def run(self) -> None:
        """Main execution method for the thread, performs the PDF extraction."""
        processed_files: List[str] = []
        try:
            total_pdfs = len(self.pdf_paths)
            for pdf_index, pdf_path in enumerate(self.pdf_paths):
                processed_files.append(
                    self.process_single_pdf(pdf_index, pdf_path, total_pdfs)
                )

            self.emit_completion_result(processed_files, total_pdfs)
        except Exception as e:
            self.extraction_error.emit(f"Error extracting PDFs: {str(e)}")

    def process_single_pdf(self, pdf_index: int, pdf_path: str, total_pdfs: int) -> str:
        """Process a single PDF file and return its base name."""
        with fitz.open(pdf_path) as doc:
            base_name = self.get_base_name(pdf_path)
            extracted_text, csv_data = self.process_pdf_pages(
                doc, base_name, pdf_index, total_pdfs
            )
            self.save_output_files(base_name, extracted_text, csv_data)
        return f"{base_name}.pdf"

    def get_base_name(self, pdf_path: str) -> str:
        """Extract the base filename without extension."""
        return os.path.splitext(os.path.basename(pdf_path))[0]

    def process_pdf_pages(
        self, doc: fitz.Document, base_name: str, pdf_index: int, total_pdfs: int
    ) -> Tuple[str, List[Dict[str, Any]]]:
        """Process all pages in a PDF document."""
        extracted_text = ""
        csv_data = []
        total_pages = len(doc)

        if self.extract_images:
            self.setup_images_directory(base_name)

        for page_num, page in enumerate(cast(Iterable[Any], doc)):
            page_text = page.get_text()
            extracted_text = self.process_page_text(
                extracted_text, base_name, page_num, page_text
            )
            csv_data = self.process_page_csv(csv_data, base_name, page_num, page_text)
            self.process_page_images(page, base_name, page_num)
            self.update_progress(pdf_index, page_num, total_pages, total_pdfs)

        return extracted_text, csv_data

    def setup_images_directory(self, base_name: str) -> None:
        """Create directory for images if needed."""
        images_dir = os.path.join(self.output_path, "images")
        os.makedirs(images_dir, exist_ok=True)

    def process_page_text(
        self, current_text: str, base_name: str, page_num: int, page_text: str
    ) -> str:
        """Process and append text for a single page."""
        if self.extract_text or self.export_csv:
            return (
                current_text
                + f"--- {base_name} - Page {page_num + 1} ---\n{page_text}\n\n"
            )
        return current_text

    def process_page_csv(
        self,
        csv_data: List[Dict[str, Any]],
        base_name: str,
        page_num: int,
        page_text: str,
    ) -> List[Dict[str, Any]]:
        """Process CSV data for a single page."""
        if not self.export_csv:
            return csv_data

        try:
            fields = self.field_extractor.extract_fields(page_text)
            row = {"File": base_name, "Page": page_num + 1}
            row.update(fields)
            return csv_data + [row]
        except Exception as e:
            print(f"Warning processing page {page_num + 1}: {str(e)}")
            return csv_data

    def process_page_images(
        self, page: fitz.Page, base_name: str, page_num: int
    ) -> None:
        """Extract images from a single page."""
        if not self.extract_images:
            return

        images_dir = os.path.join(self.output_path, "images")
        for img_index, img_info in enumerate(page.get_images(full=True)):
            self.save_page_image(img_info, base_name, page_num, img_index, images_dir)

    def save_page_image(
        self,
        img_info: Tuple,
        base_name: str,
        page_num: int,
        img_index: int,
        images_dir: str,
    ) -> None:
        """Save a single image from a page."""
        xref = img_info[0]
        base_image = img_info[1].extract_image(xref)
        image_filename = (
            f"{base_name}_page{page_num + 1}_img{img_index + 1}.{base_image['ext']}"
        )
        with open(os.path.join(images_dir, image_filename), "wb") as img_file:
            img_file.write(base_image["image"])

    def save_output_files(
        self, base_name: str, extracted_text: str, csv_data: List[Dict[str, Any]]
    ) -> None:
        """Save all output files for a single PDF."""
        if self.extract_text and extracted_text:
            self.save_text_file(base_name, extracted_text)
        if self.export_csv and csv_data:
            self.save_csv_file(base_name, csv_data)

    def save_text_file(self, base_name: str, text_content: str) -> None:
        """Save extracted text to file."""
        text_file_path = os.path.join(self.output_path, f"{base_name}_text.txt")
        with open(text_file_path, "w", encoding="utf-8") as text_file:
            text_file.write(text_content)

    def save_csv_file(self, base_name: str, csv_data: List[Dict[str, Any]]) -> None:
        """Save CSV data to file."""
        csv_file_path = os.path.join(self.output_path, f"{base_name}_data.csv")
        fieldnames = sorted({field for row in csv_data for field in row.keys()})

        with open(csv_file_path, "w", newline="", encoding="utf-8") as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(csv_data)

    def update_progress(
        self, pdf_index: int, page_num: int, total_pages: int, total_pdfs: int
    ) -> None:
        """Calculate and emit progress update."""
        progress = int(((pdf_index + (page_num + 1) / total_pages) / total_pdfs) * 100)
        self.progress_update.emit(progress)

    def emit_completion_result(
        self, processed_files: List[str], total_pdfs: int
    ) -> None:
        """Emit the appropriate completion signal."""
        if total_pdfs > 1:
            self.extraction_complete.emit("All PDFs processed.", processed_files)
        else:
            self.extraction_complete.emit(
                "", processed_files
            )  # Text is handled separately


class ConfigEditorDialog(QWidget):
    """Dialog window for editing field extraction configurations."""

    def __init__(self, field_extractor: FieldExtractor) -> None:
        """Initialize the configuration editor dialog.

        Args:
            field_extractor: The FieldExtractor instance to configure.
        """
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

    def refresh_field_list(self) -> None:
        """Refresh the list of fields in the UI."""
        self.field_list.clear()
        for field in self.field_extractor.fields:
            self.field_list.addItem(
                f"{field['name']} ({'required' if field.get('required', False) else 'optional'})"
            )

    def field_selected(self, current: Optional[QListWidgetItem]) -> None:
        """Update field details when a field is selected.

        Args:
            current: The currently selected list item.
        """
        if current:
            index: int = self.field_list.row(current)
            field: Dict[str, Any] = self.field_extractor.fields[index]
            self.name_edit.setPlainText(field["name"])
            self.pattern_edit.setPlainText(field["pattern"])
            self.required_checkbox.setChecked(field.get("required", False))

    def add_field(self) -> None:
        """Add a new field to the configuration."""
        name: str = self.name_edit.toPlainText().strip()
        pattern: str = self.pattern_edit.toPlainText().strip()
        required: bool = self.required_checkbox.isChecked()

        if not name or not pattern:
            QMessageBox.warning(self, "Warning", "Field name and pattern are required.")
            return

        self.field_extractor.fields.append(
            {"name": name, "pattern": pattern, "required": required}
        )
        self.refresh_field_list()

    def update_field(self) -> None:
        """Update the currently selected field."""
        current: Optional[QListWidgetItem] = self.field_list.currentItem()
        if not current:
            return

        index: int = self.field_list.row(current)
        name: str = self.name_edit.toPlainText().strip()
        pattern: str = self.pattern_edit.toPlainText().strip()
        required: bool = self.required_checkbox.isChecked()

        if not name or not pattern:
            QMessageBox.warning(self, "Warning", "Field name and pattern are required.")
            return

        self.field_extractor.fields[index] = {
            "name": name,
            "pattern": pattern,
            "required": required,
        }
        self.refresh_field_list()

    def remove_field(self) -> None:
        """Remove the currently selected field."""
        current: Optional[QListWidgetItem] = self.field_list.currentItem()
        if not current:
            return

        index: int = self.field_list.row(current)
        self.field_extractor.fields.pop(index)
        self.refresh_field_list()

    def save_config(self) -> None:
        """Save the current configuration to file."""
        try:
            self.field_extractor.save_config()
            QMessageBox.information(
                self, "Success", "Configuration saved successfully."
            )
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save config: {str(e)}")


class PDFExtractorApp(QMainWindow):
    """Main application window for PDF Extractor."""

    def __init__(self) -> None:
        """Initialize the main application window."""
        super().__init__()
        self.setWindowTitle("PDF Extractor")
        self.setMinimumSize(900, 600)
        self.set_dark_mode()

        self.pdf_paths: List[str] = []
        self.output_dir: Optional[str] = None
        self.field_extractor = FieldExtractor()

        self.init_ui()

    def set_dark_mode(self) -> None:
        """Configure the application to use a dark theme."""
        dark_palette = QPalette()
        dark_palette.setColor(QPalette.Window, QColor(53, 53, 53))
        dark_palette.setColor(QPalette.WindowText, Qt.GlobalColor.white)
        dark_palette.setColor(QPalette.Base, QColor(35, 35, 35))
        dark_palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
        dark_palette.setColor(QPalette.ToolTipBase, Qt.GlobalColor.white)
        dark_palette.setColor(QPalette.ToolTipText, Qt.GlobalColor.white)
        dark_palette.setColor(QPalette.Text, Qt.GlobalColor.white)
        dark_palette.setColor(QPalette.Button, QColor(53, 53, 53))
        dark_palette.setColor(QPalette.ButtonText, Qt.GlobalColor.white)
        dark_palette.setColor(QPalette.BrightText, Qt.GlobalColor.red)
        dark_palette.setColor(QPalette.Link, QColor(42, 130, 218))
        QApplication.setPalette(dark_palette)

    def init_ui(self) -> None:
        """Initialize the user interface components."""
        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)

        title_label = QLabel("PDF Extractor")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
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

    def open_config_editor(self) -> None:
        """Open the field configuration editor dialog."""
        self.config_editor = ConfigEditorDialog(self.field_extractor)
        self.config_editor.show()

    def toggle_preview(self) -> None:
        """Toggle visibility of the extracted text preview panel."""
        self.preview_group.setVisible(self.preview_toggle_button.isChecked())

    def mode_changed(self) -> None:
        """Handle changes to the processing mode (single file vs folder)."""
        self.file_path_label.setText("No file or folder selected")
        self.pdf_paths = []
        self.check_extract_button()

    def select_file_or_folder(self) -> None:
        """Select PDF file(s) to process based on current mode."""
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

    def select_output_directory(self) -> None:
        """Select output directory for extracted content."""
        output_dir = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if output_dir:
            self.output_dir = output_dir
            self.output_path_label.setText(output_dir)
            self.check_extract_button()

    def check_extract_button(self) -> None:
        """Enable/disable extract button based on current selections."""
        self.extract_button.setEnabled(
            bool(self.pdf_paths) and self.output_dir is not None
        )

    def start_extraction(self) -> None:
        """Start the PDF extraction process with current settings."""
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

        if self.output_dir is None:
            raise ValueError("Output directory must be selected.")

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

    def update_progress(self, value: int) -> None:
        """Update the progress bar.

        Args:
            value: Current progress value (0-100).
        """
        self.progress_bar.setValue(value)

    def on_extraction_complete(
        self, extracted_text: str, processed_files: List[str]
    ) -> None:
        """Handle completion of PDF extraction.

        Args:
            extracted_text: The extracted text content (for single file mode).
            processed_files: List of processed file names.
        """
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

    def on_extraction_error(self, error_message: str) -> None:
        """Handle extraction errors.

        Args:
            error_message: Description of the error that occurred.
        """
        self.extract_button.setEnabled(True)
        QMessageBox.critical(self, "Error", error_message)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFExtractorApp()
    window.show()
    sys.exit(app.exec_())
