# Event Management System

<<<<<<< HEAD
This repository contains a Python-based event management system that facilitates the generation of personalized event tickets in PDF format, sends them via email, and tracks attendance using QR codes.
=======
A Python-based QR code attendance tracking system with email confirmations. Built with Tkinter for GUI, OpenCV for QR code scanning for generating personalized event tickets.
>>>>>>> 5b705d3283c70b245948ce889a5ab7c0e12cc2a0

## Features

- **Generate PDF Tickets**: Create customized PDF tickets containing participant details and unique QR codes.
- **Email Functionality**: Send personalized tickets to participants' email addresses with customizable email subjects and bodies.
- **CSV Integration**: Import participant details from a CSV file.
- **Email Template Support**: Load email templates from text files to streamline the email sending process.
- **Preview Email**: Preview the email content before sending to ensure the correct information is included.
- **Scheduled Sending**: Schedule emails to be sent at a specified time in the future.
- **Attendance Tracking**: Scan QR codes at the event to mark attendance, storing the data for further analysis.
- **Progress Monitoring**: Monitor the email sending process with a visual progress bar.

## Requirements

- Python 3.x
- Required Python packages (can be installed via `pip`):
  - `qrcode`
  - `tkinter`
  - `reportlab`
  - `smtplib`
  - `pandas`
  - `dotenv`

## Installation

1. Clone this repository:
    ```bash
    git clone https://github.com/nairp126/qr-code-attendance-scanner.git
    cd event-management-system
    ```

2. Install the required dependencies:
    ```bash
    pip install -r requirements.txt
    ```

3. Set up the `.env` file for secure credentials:
    - Create a `.env` file in the root directory with the following content:
      ```bash
      EMAIL_USER=your-email@example.com
      EMAIL_PASS=your-email-password
      ```
    - Replace `your-email@example.com` and `your-email-password` with your email credentials.

4. Run the program:
    ```bash
    python attendance_scanner.py
    ```

## Usage

- Select the Excel file containing participant data (Name, Email, UniqueID).
- Press the **Start Scanning** button to begin scanning QR codes.
- The system will scan QR codes, match the scanned data with the participant list, and send confirmation emails.
- You can stop scanning at any time by pressing the **Stop Scanning** button.
- You can load email templates and customize the email subject and body before scanning.
- All scanned data will be saved in a CSV file, and the attendance status will be updated in the Excel file.
