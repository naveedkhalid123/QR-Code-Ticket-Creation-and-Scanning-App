import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QVBoxLayout, QMessageBox
from PyQt5.QtGui import QImage, QPixmap, QFont
from PyQt5.QtCore import Qt, QTimer
import cv2
import pandas as pd
import json

class QRCodeScanningApp(QWidget):
    def __init__(self):
        super().__init__()

        # Initialize the Excel filename and a dictionary to track scanned students
        self.excel_filename = "students_data.xlsx"
        self.scanned_students = {}

        # Initialize the UI and load the camera
        self.init_ui()
        self.load_camera()

    def init_ui(self):
        # Set up the UI elements (labels, widgets, etc.)
        self.setWindowTitle('QR Code Scanning')
        self.setGeometry(100, 100, 800, 400)

        # Set font and styles
        font = QFont()
        font.setPointSize(12)

        # Title Label
        title_label = QLabel('QR Code Scanning', self)
        title_label.setFont(QFont('Arial', 20, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("QLabel { color: white; }")

        # Scanned Data Label
        self.scanned_data_label = QLabel('Scanned Data:', self)
        self.scanned_data_label.setFont(font)
        self.scanned_data_label.setAlignment(Qt.AlignCenter)
        self.scanned_data_label.setStyleSheet("QLabel { color: white; }")

        # Scanned Data Display
        self.scanned_data_display = QLabel(self)
        self.scanned_data_display.setFont(font)
        self.scanned_data_display.setStyleSheet(
            "QLabel { color : white; background-color: #333; border-radius: 5px; padding: 5px; }")
        self.scanned_data_display.setAlignment(Qt.AlignCenter)

        # CV Widget (Camera Screen)
        self.cv_widget = QLabel(self)
        self.cv_widget.setAlignment(Qt.AlignCenter)
        self.cv_widget.setStyleSheet(
            "QLabel { background-color: #4CAF50; border-radius: 10px; padding: 10px; border: 2px solid #45a049; }")

        # Layout
        layout = QVBoxLayout(self)
        layout.addWidget(title_label)
        layout.addWidget(self.scanned_data_label)
        layout.addWidget(self.scanned_data_display)
        layout.addWidget(self.cv_widget)

        # Apply styles
        self.setStyleSheet(
            "QWidget { background-color: #333; }"
        )

        self.setLayout(layout)

    def load_camera(self):
        # Open the camera and set up a timer to read QR codes
        self.cap = cv2.VideoCapture(0)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.read_qr_code_from_camera)
        self.timer.start(30)

    def read_qr_code_from_camera(self):
        # Read frames from the camera and detect QR codes
        ret, frame = self.cap.read()

        if not ret:
            print("Error reading frame from the camera.")
            return

        detector = cv2.QRCodeDetector()
        data, _, _ = detector.detectAndDecode(frame)

        if data:
            print(f"Scanned data: {data}")
            try:
                # Preprocess data to ensure it is valid JSON
                data = data.replace("'", "\"")  # Replace single quotes with double quotes
                student_info = json.loads(data)
                student_name = student_info.get("Name")
                student_section = student_info.get("Section")

                if student_name and student_section:
                    if self.check_entry_in_excel(student_name, student_section):
                        if (student_name, student_section) not in self.scanned_students:
                            self.scanned_students[(student_name, student_section)] = True
                            self.scanned_data_display.setText(
                                f"Student is registered: {student_name}, Section: {student_section}")
                            self.show_message_box(
                                f"Student is registered: {student_name}, Section: {student_section}")
                        else:
                            self.scanned_data_display.setText("Student already scanned.")
                            self.show_warning_message_box("Student already scanned.")
                    else:
                        self.scanned_data_display.setText("Student Not Registered.")

                    cv2.waitKey(2000)

            except Exception as e:
                print(f"Error processing scanned data: {e}")

        # Display the camera frame in the UI
        h, w, ch = frame.shape
        bytes_per_line = ch * w
        convert_to_qt_format = QImage(frame.data, w, h, bytes_per_line, QImage.Format_RGB888)
        p = QPixmap.fromImage(convert_to_qt_format)
        self.cv_widget.setPixmap(p)
    def show_warning_message_box(self, message):
        # Display a warning message box with an OK button
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Warning")
        msg_box.setIcon(QMessageBox.Warning)
        msg_box.setText(message)
        msg_box.addButton(QMessageBox.Ok)

        # Set the style sheet for the labels inside the QMessageBox
        for label in msg_box.findChildren(QLabel):
            label.setStyleSheet("QLabel { color: white; }")
        msg_box.exec_()

    def check_entry_in_excel(self, student_name, student_section):
        # Read Excel data and check if the entry exists
        df = pd.read_excel(self.excel_filename)
        matching_entry = df[(df['Name'] == student_name) & (df['Section'] == student_section)]
        return not matching_entry.empty

    def show_message_box(self, message):
        # Display a message box with the registration status
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Registration Status")
        msg_box.setText(message)
        msg_box.addButton(QMessageBox.Ok)

        # Set the style sheet for the labels inside the QMessageBox
        for label in msg_box.findChildren(QLabel):
            label.setStyleSheet("QLabel { color: white; }")
        msg_box.exec_()

    def closeEvent(self, event):
        # Release the camera when the application is closed
        self.cap.release()
        event.accept()

def run_cool_styled_qr_scanning_app():
    # Run the PyQt application
    app = QApplication(sys.argv)
    cool_styled_qr_scanning_app = QRCodeScanningApp()
    cool_styled_qr_scanning_app.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    run_cool_styled_qr_scanning_app()