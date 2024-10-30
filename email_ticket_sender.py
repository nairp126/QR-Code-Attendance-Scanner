import qrcode
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import pandas as pd
import threading
import time
from datetime import datetime
import tkinter.ttk as ttk
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Function to generate the PDF ticket
def generate_ticket_pdf(name, unique_id):
    file_name = f"{name}_Event_Ticket.pdf"
    c = canvas.Canvas(file_name, pagesize=letter)
    c.setFont("Helvetica-Bold", 24)
    c.drawCentredString(4.25 * inch, 10 * inch, "Event Ticket for IBM ICE Day")
    c.setFont("Helvetica", 18)
    c.drawString(1 * inch, 8.5 * inch, f"Name: {name}")
    c.drawString(1 * inch, 8 * inch, f"Unique ID: {unique_id}")
    qr_img = qrcode.make(unique_id, box_size=15)
    qr_img.save("qrcode.png")
    c.drawInlineImage("qrcode.png", 1.75 * inch, 2.75 * inch, 5 * inch, 5 * inch)
    c.drawCentredString(4.25 * inch, 2.75 * inch, "Thank you for registering!")
    c.save()
    return file_name

# Function to send email with the PDF ticket as attachment
def send_email(name, email, pdf_file, email_body, subject, attachments=[]):
    from_email = os.getenv("EMAIL_USER")
    from_password = os.getenv("EMAIL_PASS")

    if from_email is None or from_password is None:
        print("Email credentials are not set.")
        return

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = email
    msg['Subject'] = subject

    body = email_body.format(name=name)  # Customize body with participant's name
    msg.attach(MIMEText(body, 'plain'))

    # Attach PDF ticket
    with open(pdf_file, "rb") as f:
        mime_base = MIMEBase('application', 'octet-stream')
        mime_base.set_payload(f.read())
        encoders.encode_base64(mime_base)
        mime_base.add_header('Content-Disposition', f'attachment; filename={os.path.basename(pdf_file)}')
        msg.attach(mime_base)

    # Attach additional files
    for attachment in attachments:
        with open(attachment, "rb") as f:
            mime_base = MIMEBase('application', 'octet-stream')
            mime_base.set_payload(f.read())
            encoders.encode_base64(mime_base)
            mime_base.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment)}')
            msg.attach(mime_base)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    try:
        server.login(from_email, from_password)
        server.sendmail(from_email, email, msg.as_string())
    except Exception as e:
        print(f"Failed to send email: {e}")
        return "Failed"
    finally:
        server.quit()
    return "Sent"

# Function to generate and send personalized tickets
def send_personalized_tickets(file_path, email_body, subject, attachments):
    data = pd.read_csv(file_path)
    required_columns = ['Name', 'Email', 'UniqueID']
    if not all(column in data.columns for column in required_columns):
        raise ValueError(f"CSV file must contain the following columns: {', '.join(required_columns)}")

    email_log = pd.DataFrame(columns=["Name", "Email", "UniqueID", "Status"])
    total_sent = 0
    total_failed = 0

    for index, row in data.iterrows():
        name = row['Name']
        email = row['Email']
        unique_id = row['UniqueID']

        if '@' in email:  # Simple email validation
            pdf_file = generate_ticket_pdf(name, unique_id)
            status = send_email(name, email, pdf_file, email_body, subject, attachments)
            if status == "Sent":
                total_sent += 1
            else:
                total_failed += 1

            email_log = pd.concat([email_log, pd.DataFrame([{"Name": name, "Email": email, "UniqueID": unique_id, "Status": status}])], ignore_index=True)
            os.remove("qrcode.png")
            os.remove(pdf_file)

    email_log.to_csv("email_log.csv", index=False)
    print("Process completed and log saved.")
    print(f"Total Sent: {total_sent}, Total Failed: {total_failed}")

    return total_sent, total_failed

# GUI Functions
def browse_file():
    filename = filedialog.askopenfilename(title="Select CSV File", filetypes=(("CSV Files", "*.csv"), ("All Files", "*.*")))
    file_path_var.set(filename)

# Function to load email template
def load_template():
    template_file = filedialog.askopenfilename(title="Select Email Template", filetypes=(("Text Files", "*.txt"), ("All Files", "*.*")))
    if template_file:
        with open(template_file, 'r') as file:
            lines = file.readlines()
            if lines:
                subject = lines[0].strip()  # Read the first line as subject
                body = ''.join(lines[1:]).strip()  # Read the rest as body
                subject_var.set(subject)  # Set subject in the subject entry
                email_body_text.delete("1.0", tk.END)  # Clear existing content
                email_body_text.insert(tk.END, body)  # Load new body content

def preview_email():
    email_body = email_body_text.get("1.0", tk.END)
    subject = subject_var.get()
    name = simpledialog.askstring("Input", "Enter the name for preview:")
    if name:
        preview_window = tk.Toplevel(root)
        preview_window.title("Email Preview")
        
        # Use .format_map to prevent KeyError
        try:
            preview_text = f"Subject: {subject}\n\n" + email_body.format_map({"name": name})
        except KeyError as e:
            messagebox.showerror("Error", f"Placeholder not found: {e}")
            return
            
        preview_label = tk.Label(preview_window, text=preview_text, justify="left")
        preview_label.pack(padx=10, pady=10)


def schedule_send():
    schedule_time = simpledialog.askstring("Input", "Enter the time to send the emails (format: YYYY-MM-DD HH:MM):")
    if schedule_time:
        scheduled_time = datetime.strptime(schedule_time, "%Y-%m-%d %H:%M")
        now = datetime.now()
        delay = (scheduled_time - now).total_seconds()
        if delay < 0:
            messagebox.showerror("Error", "Scheduled time must be in the future.")
            return
        threading.Thread(target=delayed_send, args=(delay,)).start()

def delayed_send(delay):
    time.sleep(delay)
    send_tickets()

def send_tickets():
    file_path = file_path_var.get()
    if not file_path:
        messagebox.showerror("Error", "Please select a CSV file.")
        return

    email_body = email_body_text.get("1.0", tk.END)
    subject = subject_var.get()
    if not email_body.strip():
        messagebox.showerror("Error", "Please enter an email body.")
        return

    try:
        attachments = filedialog.askopenfilenames(title="Select Additional Attachments", filetypes=(("All Files", "*.*"),))
        total_sent, total_failed = send_personalized_tickets(file_path, email_body, subject, attachments)
        messagebox.showinfo("Success", f"Tickets sent successfully!\nTotal Sent: {total_sent}\nTotal Failed: {total_failed}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to send tickets: {e}")

# Create the main window
root = tk.Tk()
root.title("Event Ticket Sender")

# Variable to hold the file path
file_path_var = tk.StringVar()
subject_var = tk.StringVar()

# Create and place the components
tk.Label(root, text="Select CSV File:").pack(pady=10)
tk.Entry(root, textvariable=file_path_var, width=50).pack(pady=5)
tk.Button(root, text="Browse", command=browse_file).pack(pady=5)

tk.Label(root, text="Customize Email Subject:").pack(pady=10)
tk.Entry(root, textvariable=subject_var, width=50).pack(pady=5)

tk.Label(root, text="Customize Email Body:").pack(pady=10)
email_body_text = scrolledtext.ScrolledText(root, width=60, height=10)
email_body_text.pack(pady=5)

# Button to load email template
tk.Button(root, text="Load Email Template", command=load_template).pack(pady=5)
tk.Button(root, text="Preview Email", command=preview_email).pack(pady=5)
tk.Button(root, text="Schedule Email", command=schedule_send).pack(pady=5)
tk.Button(root, text="Send Tickets", command=send_tickets).pack(pady=20)

# Progress bar for sending
progress = ttk.Progressbar(root, length=300, mode='determinate')
progress.pack(pady=10)

# Start the Tkinter event loop
root.mainloop()
