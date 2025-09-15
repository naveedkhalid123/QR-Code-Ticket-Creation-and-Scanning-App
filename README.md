# QR-Code-Ticket-Creation-and-Scanning-App

This project consists of two main Python applications:

QR Code Generation App: Creates QR codes for student information and saves the data to an Excel file.
QR Code Scanning App: Scans QR codes using a webcam, checks student registration, and displays relevant messages.
Both applications feature sleek GUIs built with PyQt and handle student data through Excel files using Pandas.

Features
QR Code Generation
Input Student Information: Enter details like name, section, and payment label.
Generate QR Codes: Creates a QR code for each student.
Save Data: Stores student information in an Excel file to avoid duplication.
QR Code Scanning
Scan QR Codes: Uses a webcam to read QR codes.
Check Registration: Verifies if the student is registered.
Display Messages: Shows messages for registered or unregistered students.
Installation
Clone the Repository
git clone https://github.com/Fazeel-AIML/QR_Code_Ticket_Generator.git
cd QR_Code_Ticket_Generator
Libraries
PyQt5
qrcode
pandas
opencv-python
