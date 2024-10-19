import cv2
from pyzbar.pyzbar import decode
import pandas as pd
import datetime
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import threading
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Initialize scanning flag
scanning = False

# Function to send confirmation email in a separate thread
def send_confirmation_email(name, email):
    from_email = os.getenv("EMAIL_USER")
    from_password = os.getenv("EMAIL_PASS")

    # Default subject and body
    default_subject = "Attendance Confirmation"
    default_body = f"Dear {name},\n\nThank you for attending the event! Your attendance has been recorded.\n\nBest regards,\n[Your Name]"

    subject = email_subject_var.get() or default_subject
    body = email_body_var.get("1.0", tk.END).strip() or default_body
    
    # Replace {Name} placeholder with the participant's name
    body = body.replace("{Name}", name)
    
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_email, from_password)
        server.sendmail(from_email, email, msg.as_string())
        server.quit()
        print(f"Email sent to {name} at {email}.")
    except Exception as e:
        print(f"Failed to send email to {name}: {e}")

# Function to scan QR codes and store unique IDs along with names
def scan_qr_code(main_sheet):
    global scanning
    cap = cv2.VideoCapture(0)

    if not cap.isOpened():
        messagebox.showerror("Error", "Could not open the camera.")
        return

    scanned_data = []
    last_scanned = set()  # Keep track of scanned IDs

    while scanning:
        ret, frame = cap.read()

        if not ret:
            messagebox.showerror("Error", "Failed to capture image from camera.")
            break

        # Decode the QR code
        barcodes = decode(frame)
        for barcode in barcodes:
            unique_id = barcode.data.decode('utf-8')
            if unique_id not in last_scanned:  # Process only if not already scanned
                last_scanned.add(unique_id)  # Mark as scanned

                # Check if the Unique ID is in the main sheet
                participant_data = main_sheet.loc[main_sheet['Unique ID'] == unique_id]
                if not participant_data.empty:
                    name = participant_data['Name'].values[0]
                    email = participant_data['Email'].values[0]

                    print(f"Scanned Unique ID: {unique_id}, Name: {name}")

                    scanned_data.append({
                        'Unique ID': unique_id,
                        'Name': name,
                        'Email': email,
                        'Timestamp': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    })

                    # Send confirmation email in a separate thread
                    threading.Thread(target=send_confirmation_email, args=(name, email)).start()

                    # Update the table with the scanned data
                    update_table(scanned_data)

        # Display the frame
        cv2.imshow('QR Code Scanner', frame)

        # Stop scanning if the user closes the window
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()

    # Save the scanned data to CSV
    save_to_csv(scanned_data)

    # Update attendance in the main sheet
    update_attendance(main_sheet, scanned_data)

def update_table(scanned_data):
    for i in tree.get_children():
        tree.delete(i)  # Clear previous entries

    for data in scanned_data:
        tree.insert("", tk.END, values=(data['Unique ID'], data['Name'], data['Timestamp']))

def save_to_csv(scanned_data):
    df = pd.DataFrame(scanned_data)
    df.to_csv('scanned_qr_data.csv', index=False)
    messagebox.showinfo("Info", "Scanned data has been saved to 'scanned_qr_data.csv'.")

def update_attendance(main_sheet, scanned_data):
    # Convert scanned data to DataFrame
    scanned_df = pd.DataFrame(scanned_data)

    # Merge both dataframes on the 'Unique ID' column
    merged_df = pd.merge(main_sheet, scanned_df[['Unique ID', 'Name', 'Email', 'Timestamp']], on='Unique ID', how='left')

    # Mark attendance based on whether a unique ID has been scanned
    merged_df['Attendance'] = merged_df['Timestamp'].notna().replace({True: 'Present', False: 'Absent'})

    # Save the updated attendance sheet
    merged_df.to_excel('updated_attendance_sheet.xlsx', index=False)
    messagebox.showinfo("Info", "Attendance has been updated in 'updated_attendance_sheet.xlsx'.")

def start_scanning():
    global scanning
    scanning = True
    main_sheet = load_main_sheet()
    if main_sheet is not None:
        scan_thread = threading.Thread(target=scan_qr_code, args=(main_sheet,))
        scan_thread.start()

def stop_scanning():
    global scanning
    scanning = False
    messagebox.showinfo("Info", "Scanning has been stopped.")

def load_main_sheet():
    try:
        return pd.read_excel(r"C:\Users\nairp\Downloads\IBM Organising Committee List Hypervision.xlsx")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load the main attendance sheet: {e}")
        return None

def load_email_template():
    # Open file dialog to select template file
    file_path = filedialog.askopenfilename(title="Select Email Template", filetypes=[("Text Files", "*.txt")])
    if file_path:
        try:
            with open(file_path, 'r') as file:
                content = file.read()
                # Split content into subject and body
                subject, body = content.split("\n\n", 1)  # Assuming the first part is the subject and the rest is the body
                email_subject_var.set(subject.strip())
                email_body_var.delete("1.0", tk.END)
                email_body_var.insert("1.0", body.strip())
                messagebox.showinfo("Info", "Template loaded successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load the email template: {e}")

# Create the main window
root = tk.Tk()
root.title("QR Code Attendance Scanner")

# Create and place widgets
scan_button = tk.Button(root, text="Start Scanning", command=start_scanning)
scan_button.pack(pady=20)

stop_button = tk.Button(root, text="Stop Scanning", command=stop_scanning)
stop_button.pack(pady=5)

# Create a table for displaying scanned data
columns = ('Unique ID', 'Name', 'Timestamp')
tree = ttk.Treeview(root, columns=columns, show='headings')
tree.heading('Unique ID', text='Unique ID')
tree.heading('Name', text='Name')
tree.heading('Timestamp', text='Timestamp')
tree.pack(pady=20)

# Email subject and body
tk.Label(root, text="Email Subject:").pack(pady=5)
email_subject_var = tk.StringVar()
email_subject_entry = tk.Entry(root, textvariable=email_subject_var, width=50)
email_subject_entry.pack(pady=5)

tk.Label(root, text="Email Body:").pack(pady=5)
email_body_var = tk.Text(root, height=10, width=60)
email_body_var.pack(pady=5)

# Load Template Button
load_template_button = tk.Button(root, text="Load Email Template", command=load_email_template)
load_template_button.pack(pady=5)

# Start the GUI loop
root.mainloop()
