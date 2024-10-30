
### Features and Functionality

```plaintext
# Features and Functionality

## Email Ticket Sender

### 1. PDF Ticket Generation
- Functionality Generates a personalized PDF ticket for each participant upon request.
- Details Included
  - Event title
  - Participant's name
  - Unique ID assigned to the participant
  - A QR code that contains the unique ID, allowing for quick scanning at the event.
- Dependencies Uses the `reportlab` library for PDF generation and `qrcode` library for QR code creation.

### 2. Email Sending
- Functionality Sends the generated PDF ticket as an email attachment.
- Email Customization
  - The subject and body of the email can be customized.
  - Supports placeholders in the email body for dynamic content (e.g., participant’s name).
- Email Validation Ensures that the email addresses in the CSV file are valid before sending.
- Dependencies Utilizes `smtplib` for email functionality and `email.mime` for creating email messages.

### 3. CSV File Support
- Functionality Reads participant information from a specified CSV file.
- Required Columns
  - Name
  - Email
  - Unique ID
- Validation Checks that all required columns are present and properly formatted.

### 4. Email Template Loading
- Functionality Loads pre-defined email templates from text files.
- Template Structure
  - The first line is treated as the email subject.
  - Subsequent lines comprise the body of the email.
- Usage Streamlines the email preparation process, especially for recurring events.

### 5. Email Preview
- Functionality Provides a preview of the email content before sending.
- User Interaction Asks the user to input a name for a personalized preview.
- Error Handling Checks for missing placeholders and notifies the user of any issues.

### 6. Scheduled Sending
- Functionality Allows users to schedule email sending for a future date and time.
- User Input Asks the user to input the desired send time, validated against the current time.
- Threading Uses a separate thread to handle the delay without freezing the GUI.

### 7. Progress Monitoring
- Functionality Displays a progress bar to indicate the sending status of emails.
- User Feedback Provides visual feedback during the sending process.

## QR Code Attendance Scanner

### 1. Attendance Tracking
- Functionality Scans QR codes from participant tickets at the event to record attendance.
- Data Handling Matches scanned IDs with the main attendance list to mark attendance efficiently.

### 2. Data Export
- Functionality Exports scanned attendance data into CSV or Excel formats.
- Record-Keeping Keeps a detailed log of attendance for event organizers.

### 3. User-Friendly Interface
- GUI Design Built using Tkinter for ease of use, allowing event organizers to interact with the system without technical knowledge.
- Accessibility Intuitive design makes it easy for volunteers or staff to use the system effectively during events.

By combining these features, the Event Management System provides a seamless experience for event organizers, enhancing both the management process and participant satisfaction.
