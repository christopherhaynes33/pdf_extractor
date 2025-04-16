import os
import threading
import pymupdf as fitz  # PyMuPDF
import customtkinter as ctk
from tkinter import filedialog, messagebox

ctk.set_appearance_mode("Dark")  # Dark or Light
ctk.set_default_color_theme("blue")

class PDFExtractorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.extracted_texts = {}  # Dictionary to store extracted texts by filename
        self.title("🧠 PDF Extractor - Modern Edition")
        self.geometry("1000x700")
        self.resizable(True, True)

        self.pdf_paths = []
        self.output_dir = ""
        self.extract_text = ctk.BooleanVar(value=True)
        self.extract_images = ctk.BooleanVar(value=False)
        self.extracted_text = ""

        self.build_ui()

    def change_theme(self, mode):
        ctk.set_appearance_mode(mode)

    def build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Layout Frame
        layout = ctk.CTkFrame(self)
        layout.grid(row=0, column=0, sticky="nsew")
        layout.pack(padx=20, pady=20, fill="both", expand=True)

        # Title
        title = ctk.CTkLabel(layout, text="PDF Extractor", font=ctk.CTkFont(size=28, weight="bold"))
        title.pack(pady=(0, 20))

        # Theme Toggle
        self.theme_mode = ctk.StringVar(value="Dark")
        theme_frame = ctk.CTkFrame(layout)
        theme_frame.pack(pady=(0, 10))

        ctk.CTkLabel(theme_frame, text="Theme:", font=ctk.CTkFont(size=14)).pack(side="left", padx=(5, 10))
        theme_toggle = ctk.CTkOptionMenu(theme_frame, variable=self.theme_mode, values=["Light", "Dark"], command=self.change_theme)
        theme_toggle.pack(side="left")

        # Mode Select
        self.mode_option = ctk.CTkOptionMenu(layout, values=["Single PDF", "All PDFs in Folder"])
        self.mode_option.set("Single PDF")
        self.mode_option.pack(pady=5)

        # File Selector
        self.file_label = ctk.CTkLabel(layout, text="No file/folder selected", wraplength=900)
        self.file_label.pack(pady=2)
        ctk.CTkButton(layout, text="Select PDF or Folder", command=self.select_file_or_folder).pack(pady=5)

        # Output Dir
        self.output_label = ctk.CTkLabel(layout, text="No output directory selected", wraplength=900)
        self.output_label.pack(pady=2)
        ctk.CTkButton(layout, text="Select Output Directory", command=self.select_output_dir).pack(pady=5)

        # Checkboxes
        ctk.CTkCheckBox(layout, text="Extract Text", variable=self.extract_text).pack(pady=2)
        ctk.CTkCheckBox(layout, text="Extract Images", variable=self.extract_images).pack(pady=2)

        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(layout, width=600)
        self.progress_bar.set(0)
        self.progress_bar.pack(pady=10)

        # Extract button
        self.extract_button = ctk.CTkButton(layout, text="Extract PDF Contents", command=self.start_extraction, state="disabled")
        self.extract_button.pack(pady=10)

        # Processed list + Preview area
        splitter = ctk.CTkFrame(layout)
        splitter.pack(fill="both", expand=True, pady=10)
        splitter.grid_rowconfigure(0, weight=1)
        splitter.grid_columnconfigure(0, weight=1)
        splitter.grid_columnconfigure(1, weight=4)

        self.processed_box = ctk.CTkTextbox(splitter, height=100, font=ctk.CTkFont(size=14))
        self.processed_box.grid(row=0, column=0, sticky="nsew", padx=(0,10))
        self.processed_box.tag_config("clickable", foreground="blue", underline=True)
        self.processed_box.tag_bind("clickable", "<Button-1>", self.on_file_click)

        self.preview = ctk.CTkTextbox(splitter, font=ctk.CTkFont(size=14))
        self.preview.grid(row=0, column=1, sticky="nsew")

    def select_file_or_folder(self):
        self.pdf_paths.clear()
        if self.mode_option.get() == "Single PDF":
            file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if file:
                self.pdf_paths = [file]
                self.file_label.configure(text=file)
        else:
            folder = filedialog.askdirectory()
            if folder:
                for root, _, files in os.walk(folder):
                    for f in files:
                        if f.lower().endswith(".pdf"):
                            self.pdf_paths.append(os.path.join(root, f))
                self.file_label.configure(text=f"{len(self.pdf_paths)} PDFs selected")
        self.update_extract_button()

    def select_output_dir(self):
        self.output_dir = filedialog.askdirectory()
        if self.output_dir:
            self.output_label.configure(text=self.output_dir)
        self.update_extract_button()

    def update_extract_button(self):
        self.extract_button.configure(state="normal" if self.pdf_paths and self.output_dir else "disabled")

    def on_file_click(self, event):
        # Get the index of the click
        index = self.processed_box.index(f"@{event.x},{event.y}")
        line_start = index.split(".")[0] + ".0"
        line_end = index.split(".")[0] + ".end"

        # Get the clicked line
        clicked_line = self.processed_box.get(line_start, line_end).strip()

        # Find the filename in the extracted_texts dictionary
        for filename, text in self.extracted_texts.items():
            if filename in clicked_line:
                self.preview.delete("1.0", "end")
                self.preview.insert("1.0", text)
                break

    def start_extraction(self):
        if not self.extract_text.get() and not self.extract_images.get():
            messagebox.showwarning("Warning", "Select at least one extraction option.")
            return

        self.progress_bar.set(0)
        self.processed_box.delete("1.0", "end")
        self.preview.delete("1.0", "end")
        self.extracted_texts = {}  # Clear previous extracted texts
        self.extract_button.configure(state="disabled")

        def update_progress(value):
            self.progress_bar.set(value / 100)

        def on_complete(text, files):
            for f in files:
                # Add each file as clickable text
                self.processed_box.insert("end", f + "\n", "clickable")
            self.extract_button.configure(state="normal")
            messagebox.showinfo("Done", f"Extraction complete. Files saved to:\n{self.output_dir}")

        def on_error(error_msg):
            self.extract_button.configure(state="normal")
            messagebox.showerror("Error", error_msg)

        thread = PDFExtractorThread(
            pdf_paths=self.pdf_paths,
            output_path=self.output_dir,
            extract_text=self.extract_text.get(),
            extract_images=self.extract_images.get(),
            update_progress=update_progress,
            on_complete=on_complete,
            on_error=on_error,
            extracted_texts_ref=self.extracted_texts  # Pass reference to store texts
        )
        thread.start()

class PDFExtractorThread(threading.Thread):
    def __init__(self, pdf_paths, output_path, extract_text, extract_images, update_progress, on_complete, on_error, extracted_texts_ref):
        super().__init__()
        self.pdf_paths = pdf_paths
        self.output_path = output_path
        self.extract_text = extract_text
        self.extract_images = extract_images
        self.update_progress = update_progress
        self.on_complete = on_complete
        self.on_error = on_error
        self.extracted_texts_ref = extracted_texts_ref  # Reference to store texts

    def run(self):
        processed_files = []
        try:
            total_pdfs = len(self.pdf_paths)
            for pdf_index, pdf_path in enumerate(self.pdf_paths):
                doc = fitz.open(pdf_path)
                total_pages = len(doc)
                extracted_text = ""

                base_name = os.path.splitext(os.path.basename(pdf_path))[0]
                filename = os.path.basename(pdf_path)

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
                            with open(os.path.join(images_dir, image_filename), "wb") as img_file:
                                img_file.write(image_bytes)

                    progress = int(((pdf_index + (page_num + 1) / total_pages) / total_pdfs) * 100)
                    self.update_progress(progress)

                if self.extract_text:
                    text_file_path = os.path.join(self.output_path, f"{base_name}_text.txt")
                    with open(text_file_path, "w", encoding="utf-8") as text_file:
                        text_file.write(extracted_text)

                    # Store the extracted text with filename as key
                    self.extracted_texts_ref[filename] = extracted_text

                processed_files.append(filename)

            self.on_complete(extracted_text if total_pdfs == 1 else "", processed_files)
        except Exception as e:
            self.on_error(f"Error extracting PDFs: {str(e)}")

if __name__ == "__main__":
    app = PDFExtractorApp()
    app.mainloop()
