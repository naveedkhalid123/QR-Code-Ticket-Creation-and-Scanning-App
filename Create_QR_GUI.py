import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QMessageBox
from PyQt5.QtCore import Qt
import qrcode
import pandas as pd
import os

class QRCodeCreationApp(QWidget):
    def __init__(self):
        super().__init__()

        self.students_data = self.load_entries_from_excel()  #Load existing entries
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Student QR Code Creation')
        self.setGeometry(100, 100, 400, 200)

        # Set font and styles
        font = self.font()
        font.setPointSize(12)

        # Title Label
        title_label = QLabel('QR Code Creation', self)
        title_label.setFont(self.font())
        title_label.setAlignment(Qt.AlignCenter)  # Center align the title
        title_label.setStyleSheet("font-weight: bold; font-size: 14pt;")  # Make it bold and slightly bigger

        # Labels and Edits
        self.name_label = QLabel('Enter student name:', self)
        self.name_label.setFont(font)
        self.name_edit = QLineEdit(self)
        self.name_edit.setStyleSheet("QLineEdit { color: white; }")

        self.section_label = QLabel('Enter student section:', self)
        self.section_label.setFont(font)
        self.section_edit = QLineEdit(self)
        self.section_edit.setStyleSheet("QLineEdit { color: white; }")

        self.payment_label = QLabel('Enter payment label:', self)
        self.payment_label.setFont(font)
        self.payment_edit = QLineEdit(self)
        self.payment_edit.setStyleSheet("QLineEdit { color: white; }")

        # Generate Button
        self.generate_button = QPushButton('Generate QR Code', self)
        self.generate_button.clicked.connect(self.generate_qr_code)
        self.generate_button.setFont(font)

        # Layout
        layout = QVBoxLayout(self)
        layout.addWidget(title_label)
        layout.addWidget(self.name_label)
        layout.addWidget(self.name_edit)
        layout.addWidget(self.section_label)
        layout.addWidget(self.section_edit)
        layout.addWidget(self.payment_label)
        layout.addWidget(self.payment_edit)
        layout.addWidget(self.generate_button)

        # Apply styles
        self.setStyleSheet(
            "QWidget { background-color: #333; border-radius: 10px; padding: 20px; margin: 10px; border: 2px solid #45a049; }"
            "QLabel { color: white; }"
            "QPushButton { background-color: #45a049; color: white; border: none; padding: 10px; border-radius: 5px; }"
            "QPushButton:hover { background-color: #4CAF50; }"
        )

        self.setLayout(layout)

    def generate_qr_code(self):
        student_name = self.name_edit.text().lower()
        student_section = self.section_edit.text().lower()
        payment_label = self.payment_edit.text()

        if self.check_entry_in_excel(student_name, student_section):
            self.show_message_box(f"Student already exists: {student_name}, Section: {student_section}")
            return

        student_data = {"Name": student_name, "Section": student_section, "PaymentLabel": payment_label}
        self.students_data.append(student_data)

        qr_code_filename = f"qr_codes/{student_name}_{student_section}_qr.png"
        self.create_qr_code(str(student_data), qr_code_filename)
        print(f"QR code generated for {student_name}, Section: {student_section}")

        # Save all entries to Excel file after generating each QR code
        self.save_entries_to_excel()

    def create_qr_code(self, data, filename):
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        qr.add_data(data)
        qr.make(fit=True)

        img = qr.make_image(fill_color="black", back_color="white")
        img.save(filename, format="PNG")

    def save_entries_to_excel(self):
        excel_filename = "students_data.xlsx"

        if not os.path.exists(excel_filename):
            # Create the file if it doesn't exist
            pd.DataFrame(columns=["Name", "Section", "PaymentLabel"]).to_excel(excel_filename, index=False)

        df = pd.DataFrame(self.students_data, columns=["Name", "Section", "PaymentLabel"])
        df.to_excel(excel_filename, index=False)
        print("Entries saved to Excel file.")

    def load_entries_from_excel(self):
        excel_filename = "students_data.xlsx"
        try:
            if not os.path.exists(excel_filename):
                # Create the file if it doesn't exist
                pd.DataFrame(columns=["Name", "Section", "PaymentLabel"]).to_excel(excel_filename, index=False)

            df = pd.read_excel(excel_filename)
            return df.to_dict(orient='records')
        except FileNotFoundError:
            print("No existing entries found.")
            return []

    def check_entry_in_excel(self, student_name, student_section):
        df = pd.read_excel("students_data.xlsx")
        matching_entry = df[(df['Name'].str.lower() == student_name) & (df['Section'].str.lower() == student_section)]
        return not matching_entry.empty

    def show_message_box(self, message):
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Registration Status")
        msg_box.setText(message)
        msg_box.addButton(QMessageBox.Ok)
        msg_box.exec_()

def run_cool_styled_qr_creation_app():
    app = QApplication(sys.argv)
    cool_styled_qr_creation_app = QRCodeCreationApp()
    cool_styled_qr_creation_app.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    run_cool_styled_qr_creation_app()