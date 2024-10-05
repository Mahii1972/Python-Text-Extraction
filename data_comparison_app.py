import sys
import pandas as pd
import re
import os
import pdfplumber
from bs4 import BeautifulSoup
import msoffcrypto
import io
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QLineEdit, QFileDialog, QTableWidget, QTableWidgetItem, QMessageBox, QInputDialog
from PyQt5.QtGui import QIcon

# Function to clean phone numbers
def clean_phone_number(phone):
    return re.sub(r'\D', '', str(phone).split('.')[0])

# Function to extract text from PDF
def extract_text_from_pdf(file_path):
    text = ''
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text()
    return text

# Function to extract text from TXT
def extract_text_from_txt(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        return file.read()

# Function to extract text from HTML
def extract_text_from_html(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        soup = BeautifulSoup(file, 'html.parser')
        return soup.get_text()

# Function to extract text from Excel
def extract_text_from_excel(file_path, password=None):
    try:
        if password:
            with open(file_path, 'rb') as file:
                decrypted = io.BytesIO()
                office_file = msoffcrypto.OfficeFile(file)
                office_file.load_key(password=password)
                office_file.decrypt(decrypted)
                df = pd.read_excel(decrypted)
        else:
            df = pd.read_excel(file_path)
        return df.to_string()
    except msoffcrypto.exceptions.InvalidKeyError:
        return "encrypted"
    except Exception as e:
        if "Workbook is encrypted" in str(e) or "Can't find workbook in OLE2 compound document" in str(e):
            return "encrypted"
        else:
            QMessageBox.critical(None, "Error", f"Failed to read Excel file: {e}")
            return None

# Function to compare data
def compare_data(base_df, text):
    results = []
    text = re.sub(r'\s+', '', text)
    for index, row in base_df.iterrows():
        phone = clean_phone_number(row['Mobile No']) if pd.notnull(row['Mobile No']) else ''
        email = row['Mail ID'].replace(" ", "").lower() if pd.notnull(row['Mail ID']) else ''
        name = row['Passenger Name'].replace(" ", "").lower() if pd.notnull(row['Passenger Name']) else ''
        agency = row['Travel Agency'].replace(" ", "").lower() if pd.notnull(row['Travel Agency']) else ''

        phone_match = re.search(phone, text) if phone else None
        email_match = re.search(re.escape(email), text, re.IGNORECASE) if email else None
        name_match = re.search(re.escape(name), text, re.IGNORECASE) if name else None
        agency_match = re.search(re.escape(agency), text, re.IGNORECASE) if agency else None

        for match_type, match in [('Phone', phone_match), ('Email', email_match), ('Name', name_match), ('Agency', agency_match)]:
            if match:
                context_start = max(match.start() - 10, 0)
                context_end = match.end() + 10
                context = text[context_start:context_end]
                results.append([match_type, match.group(), context])

    return results

class TravelDataVerifier(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Travel Data Verifier")
        self.setGeometry(100, 100, 800, 600)
# Update the icon setting
        icon_path = self.resource_path('your_icon.ico')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        else:
            print(f"Warning: Icon file not found at {icon_path}")
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Base Excel File selection
        base_layout = QHBoxLayout()
        self.base_input = QLineEdit()
        base_button = QPushButton("Browse")
        base_button.clicked.connect(self.browse_base_file)
        base_layout.addWidget(QLabel("Base Excel File:"))
        base_layout.addWidget(self.base_input)
        base_layout.addWidget(base_button)
        layout.addLayout(base_layout)
        # Files to Compare selection
        compare_layout = QHBoxLayout()
        compare_label = QLabel("Files to Compare (supports(PDF (.pdf), Text (.txt), HTML (.htm, .html), Excel (.xls, .xlsx)):")
        compare_layout.addWidget(compare_label)
        layout.addLayout(compare_layout)

        self.compare_input = QLineEdit()
        compare_input_layout = QHBoxLayout()
        compare_input_layout.addWidget(self.compare_input)
        compare_button = QPushButton("Browse")
        compare_button.clicked.connect(self.browse_compare_files)
        compare_input_layout.addWidget(compare_button)
        layout.addLayout(compare_input_layout)

        # Compare button
        compare_button = QPushButton("Compare")
        compare_button.clicked.connect(self.compare_files)
        layout.addWidget(compare_button)

        # Results table
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(3)
        self.results_table.setHorizontalHeaderLabels(["Match Type", "Match String", "File Context"])
        self.results_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.results_table)

        # Add this method to handle resource paths
    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

    def browse_base_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Base Excel File", "", "Excel Files (*.xlsx *.xls)")
        if file_name:
            self.base_input.setText(file_name)

    def browse_compare_files(self):
        file_names, _ = QFileDialog.getOpenFileNames(self, "Select Files to Compare", "", "All Files (*.*)")
        if file_names:
            self.compare_input.setText(";".join(file_names))

    def compare_files(self):
        base_file = self.base_input.text()
        compare_files = self.compare_input.text().split(';')

        if not base_file or not compare_files:
            QMessageBox.warning(self, "Error", "Please select both base file and files to compare.")
            return

        try:
            base_df = pd.read_excel(base_file)
            base_df['Mobile No'] = base_df['Mobile No'].apply(clean_phone_number)
            base_df['Mail ID'] = base_df['Mail ID'].str.lower()
            base_df['Passenger Name'] = base_df['Passenger Name'].str.lower()
            base_df['Travel Agency'] = base_df['Travel Agency'].str.lower()

            all_compare_text = ""
            for compare_file in compare_files:
                file_extension = compare_file.split('.')[-1].lower()

                if file_extension == 'pdf':
                    compare_text = extract_text_from_pdf(compare_file)
                elif file_extension == 'txt':
                    compare_text = extract_text_from_txt(compare_file)
                elif file_extension in ['htm', 'html']:
                    compare_text = extract_text_from_html(compare_file)
                elif file_extension in ['xls', 'xlsx']:
                    compare_text = extract_text_from_excel(compare_file)
                    if compare_text == "encrypted":
                        password, ok = QInputDialog.getText(self, "Password Required", "Enter password for encrypted Excel file:", QLineEdit.Password)
                        if ok:
                            compare_text = extract_text_from_excel(compare_file, password)
                else:
                    QMessageBox.warning(self, "Error", f"Unsupported file type: {file_extension}")
                    continue

                if compare_text and compare_text != "encrypted":
                    all_compare_text += compare_text

            if all_compare_text:
                comparison_results = compare_data(base_df, all_compare_text)
                self.update_results_table(comparison_results)
            else:
                QMessageBox.warning(self, "Error", "No valid text found in comparison files.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred:Check the {str(e)} column")

    def update_results_table(self, results):
        self.results_table.setRowCount(len(results))
        for row, (match_type, match_string, context) in enumerate(results):
            self.results_table.setItem(row, 0, QTableWidgetItem(match_type))
            self.results_table.setItem(row, 1, QTableWidgetItem(match_string))
            self.results_table.setItem(row, 2, QTableWidgetItem(context))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = TravelDataVerifier()
    window.show()
    sys.exit(app.exec_())