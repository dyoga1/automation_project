import os
import re
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog, QHBoxLayout, QFrame
import PyPDF2
import pyperclip
import pandas as pd


class PDFExtractorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Extractor")
        self.setGeometry(100, 100, 500, 200)

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)

        self.regex_page_label = QLabel("Regex for Page Identifier:")
        layout.addWidget(self.regex_page_label)

        self.regex_page_edit = QLineEdit()
        layout.addWidget(self.regex_page_edit)

        self.regex_title_label = QLabel("Regex for PDF Title:")
        layout.addWidget(self.regex_title_label)

        self.regex_title_edit = QLineEdit()
        layout.addWidget(self.regex_title_edit)

        # Create a horizontal layout for file selection
        file_layout = QHBoxLayout()
        self.file_label = QLabel("No file selected.")
        file_layout.addWidget(self.file_label)

        self.file_button = QPushButton("Select PDF File")
        self.file_button.clicked.connect(self.select_pdf_file)
        file_layout.addWidget(self.file_button)

        layout.addLayout(file_layout)

        # Create a horizontal layout for output folder selection
        output_layout = QHBoxLayout()
        self.output_label = QLabel("No output folder selected.")
        output_layout.addWidget(self.output_label)

        self.output_button = QPushButton("Select Output Folder")
        self.output_button.clicked.connect(self.select_output_folder)
        output_layout.addWidget(self.output_button)

        layout.addLayout(output_layout)

        self.extract_button = QPushButton("Extract PDF")
        self.extract_button.clicked.connect(self.extract_pdf)
        layout.addWidget(self.extract_button)

        self.copy_text_button = QPushButton("Copy First Page Text")
        self.copy_text_button.clicked.connect(self.copy_first_page_text)
        layout.addWidget(self.copy_text_button)

        # Separator line at the bottom
        separator_line = QFrame()
        separator_line.setFrameShape(QFrame.HLine)
        separator_line.setFrameShadow(QFrame.Sunken)
        layout.addWidget(separator_line)

        # Input PDF file label and button
        file_wrapper_widget = QWidget()
        file_layout = QHBoxLayout(file_wrapper_widget)
        self.input_file_label = QLabel("Input PDF File:")
        file_layout.addWidget(self.input_file_label)
        self.input_file_button = QPushButton("Select PDF Path List")
        self.input_file_button.clicked.connect(self.select_path_list)
        file_layout.addWidget(self.input_file_button)
        layout.addWidget(file_wrapper_widget)

        # Merge PDF button
        self.merge_pdf_button = QPushButton("Merge PDF")
        self.merge_pdf_button.clicked.connect(self.merge_pdfs)
        layout.addWidget(self.merge_pdf_button)


        self.status_label = QLabel("")
        layout.addWidget(self.status_label)

        self.pdf_file_path = ""
        self.output_folder = ""

        # Load last session regex patterns and output folder
        self.load_last_session()

    def select_pdf_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Select PDF File", "", "PDF Files (*.pdf)", options=options)
        if file_path:
            self.pdf_file_path = file_path
            self.file_label.setText(os.path.basename(file_path))

    def clean_filename(self, filename):
        # Define the regex pattern to remove invalid characters
        cleaned_filename = re.sub(r'[<>:"/\\|?*:\n]', '', filename)
        return cleaned_filename

    def select_output_folder(self):
        options = QFileDialog.Options()
        output_folder = QFileDialog.getExistingDirectory(self, "Select Output Folder", options=options)
        if output_folder:
            self.output_folder = output_folder
            self.output_label.setText(output_folder)

    def extract_pdf(self):
        regex_page = self.regex_page_edit.text()
        regex_title = self.regex_title_edit.text()

        if not self.pdf_file_path or not self.output_folder or not regex_page or not regex_title:
            self.status_label.setText("Please fill in all fields.")
            return

        with open(self.pdf_file_path, "rb") as file:
            pdf_reader = PyPDF2.PdfReader(file)
            num_pages = len(pdf_reader.pages)
            # Find the start and end pages
            start_page_text = None
            extracted_pdf = None

            for page_number in range(num_pages):
                page = pdf_reader.pages[page_number]
                page_text = page.extract_text()
                
                if re.search(regex_page, page_text) and not start_page_text:
                    extracted_pdf = PyPDF2.PdfWriter()
                    extracted_pdf.add_page(page)
                    start_page_text = page_text
                elif re.search(regex_page, page_text) and start_page_text:
                    title_match = re.search(regex_title, start_page_text)
                    title = title_match.group() if title_match else f"Page{page_number + 1}"
                    title = self.clean_filename(title)
                    output_file_path = os.path.join(self.output_folder, f"{title}_{page_number}.pdf")
                    with open(output_file_path, "wb") as output_file:
                        extracted_pdf.write(output_file)
                    
                    # Mulai tandai lagi
                    extracted_pdf = PyPDF2.PdfWriter()
                    extracted_pdf.add_page(page)
                    start_page_text = page_text
                else:
                    if page_number == num_pages-1:
                        extracted_pdf.add_page(page)
                        title_match = re.search(regex_title, start_page_text)
                        title = title_match.group() if title_match else f"Page{page_number + 1}"
                        title = self.clean_filename(title)
                        output_file_path = os.path.join(self.output_folder, f"{title}_{page_number}.pdf")
                        with open(output_file_path, "wb") as output_file:
                            extracted_pdf.write(output_file)
                    else:
                        extracted_pdf.add_page(page)

        self.status_label.setText("PDF extraction completed.")
        # Save current regex patterns to remember for the next session
        self.save_last_session()

    def copy_first_page_text(self):
        if not self.pdf_file_path:
            self.status_label.setText("Please select a PDF file first.")
            return

        with open(self.pdf_file_path, "rb") as file:
            pdf_reader = PyPDF2.PdfReader(file)
            if len(pdf_reader.pages) > 0:
                first_page = pdf_reader.pages[0]
                page_text = first_page.extract_text()
                pyperclip.copy(page_text)
                self.status_label.setText("Text from first page copied to clipboard.")
            else:
                self.status_label.setText("The PDF is empty.")

    def save_last_session(self):
        with open("last_session.txt", "w") as file:
            file.write(self.regex_page_edit.text() + "\n")
            file.write(self.regex_title_edit.text())

    def load_last_session(self):
        if os.path.exists("last_session.txt"):
            with open("last_session.txt", "r") as file:
                lines = file.readlines()
                if len(lines) >= 2:
                    self.regex_page_edit.setText(lines[0].strip())
                    self.regex_title_edit.setText(lines[1].strip())

    def select_path_list(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if file_name:
            self.path_list = self.get_path_list_from_excel(file_name)

    def get_path_list_from_excel(self, excel_file):
        df = pd.read_excel(excel_file, header=None)
        path_list = df.iloc[:, 0].tolist()
        return path_list

    def merge_pdfs(self):
        if not hasattr(self, "path_list"):
            self.status_label.setText("Please select an Excel file first.")
            return

        if len(self.path_list) == 0:
            self.status_label.setText("No PDF files found in the Excel file.")
            return

        merged_pdf = PyPDF2.PdfMerger()

        for pdf_path in self.path_list:
            if not pdf_path.endswith(".pdf"):
                self.status_label.setText(f"Invalid file: {pdf_path}. Only PDF files can be merged.")
                continue

            try:
                merged_pdf.append(pdf_path)
                self.status_label.setText(f"Merging {os.path.basename(pdf_path)}...")
            except Exception as e:
                self.status_label.setText(f"Error merging {os.path.basename(pdf_path)}: {str(e)}")

        # Save the merged PDF to a new file
        output_file_path, _ = QFileDialog.getSaveFileName(self, "Save Merged PDF", "", "PDF Files (*.pdf);;All Files (*)")
        if output_file_path:
            with open(output_file_path, "wb") as output_file:
                merged_pdf.write(output_file)
            self.status_label.setText(f"Merged PDF saved to: {output_file_path}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFExtractorApp()
    window.show()
    sys.exit(app.exec_())
