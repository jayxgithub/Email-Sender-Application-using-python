import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import json
from datetime import datetime
import threading
import os

# Colors
BG_COLOR = "#f0f0f0"
BUTTON_COLOR = "#4CAF50"
BUTTON_COLOR_ALT = "#FF5722"
BUTTON_TEXT_COLOR = "white"
BUTTON_HOVER_COLOR = "#45a049"
BUTTON_ALT_COLOR = "#FF3D00"
BUTTON_ALT_HOVER_COLOR = "#E53935"


def send_email():
    sender_email = sender_entry.get()
    sender_password = password_entry.get()
    receiver_emails = [email.strip() for email in receiver_entry.get().split(',')]
    subject = subject_entry.get()
    body = body_text.get("1.0", tk.END).strip()
    files = attachments

    # Create the email message
    message = MIMEMultipart()
    message["From"] = sender_email
    message["Subject"] = subject

    # Add body to email
    message.attach(MIMEText(body, "plain"))

    # Add attachments
    for file in files:
        try:
            with open(file, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {os.path.basename(file)}",
                )
                message.attach(part)
        except Exception as e:
            messagebox.showwarning("Warning", f"Could not attach file {file}: {e}")

    # Connect to the SMTP server
    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.ehlo()
            server.starttls()
            server.login(sender_email, sender_password)

            # Send the email to multiple recipients
            for receiver_email in receiver_emails:
                message["To"] = receiver_email
                server.send_message(message)

                # Add to history
                add_to_history(sender_email, receiver_email, subject)

            messagebox.showinfo("Success", f"Email sent successfully to {len(receiver_emails)} recipient(s)")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


def schedule_email():
    try:
        send_time_str = schedule_entry.get()
        # Adjust the format based on the input
        if len(send_time_str) == 10:  # Format is YYYY-MM-DD
            send_time = datetime.strptime(send_time_str, "%Y-%m-%d")
        elif len(send_time_str) == 19:  # Format is YYYY-MM-DD HH:MM:SS
            send_time = datetime.strptime(send_time_str, "%Y-%m-%d %H:%M:%S")
        else:
            raise ValueError("Incorrect date format")

        # Check if the date is in the future
        if send_time <= datetime.now():
            messagebox.showwarning("Invalid Time", "The scheduled time must be in the future.")
            return

        delay = (send_time - datetime.now()).total_seconds()
        threading.Timer(delay, send_email).start()
        messagebox.showinfo("Scheduled", f"Email scheduled for {send_time_str}")
    except ValueError as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


def attach_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        attachments.append(file_path)
        attachments_listbox.insert(tk.END, os.path.basename(file_path))


def save_draft():
    draft = {
        "sender": sender_entry.get(),
        "password": password_entry.get(),
        "receivers": receiver_entry.get(),
        "subject": subject_entry.get(),
        "body": body_text.get("1.0", tk.END).strip(),
        "attachments": attachments
    }
    with open("draft.json", "w") as f:
        json.dump(draft, f)
    messagebox.showinfo("Draft Saved", "Draft saved successfully.")


def load_draft():
    try:
        with open("draft.json", "r") as f:
            draft = json.load(f)
            sender_entry.delete(0, tk.END)
            sender_entry.insert(0, draft["sender"])
            password_entry.delete(0, tk.END)
            password_entry.insert(0, draft["password"])
            receiver_entry.delete(0, tk.END)
            receiver_entry.insert(0, draft["receivers"])
            subject_entry.delete(0, tk.END)
            subject_entry.insert(0, draft["subject"])
            body_text.delete("1.0", tk.END)
            body_text.insert("1.0", draft["body"])

            attachments.clear()
            attachments_listbox.delete(0, tk.END)
            for file in draft["attachments"]:
                attachments.append(file)
                attachments_listbox.insert(tk.END, os.path.basename(file))

        messagebox.showinfo("Draft Loaded", "Draft loaded successfully.")
    except FileNotFoundError:
        messagebox.showerror("Error", "No draft found.")


def save_template():
    template_name = template_name_entry.get()
    if not template_name:
        messagebox.showwarning("Warning", "Template name cannot be empty.")
        return

    template = {
        "sender": sender_entry.get(),
        "password": password_entry.get(),
        "receivers": receiver_entry.get(),
        "subject": subject_entry.get(),
        "body": body_text.get("1.0", tk.END).strip(),
        "attachments": attachments
    }
    with open(f"template_{template_name}.json", "w") as f:
        json.dump(template, f)
    messagebox.showinfo("Template Saved", f"Template '{template_name}' saved successfully.")


def load_template():
    template_name = template_name_entry.get()
    if not template_name:
        messagebox.showwarning("Warning", "Template name cannot be empty.")
        return

    try:
        with open(f"template_{template_name}.json", "r") as f:
            template = json.load(f)
            sender_entry.delete(0, tk.END)
            sender_entry.insert(0, template["sender"])
            password_entry.delete(0, tk.END)
            password_entry.insert(0, template["password"])
            receiver_entry.delete(0, tk.END)
            receiver_entry.insert(0, template["receivers"])
            subject_entry.delete(0, tk.END)
            subject_entry.insert(0, template["subject"])
            body_text.delete("1.0", tk.END)
            body_text.insert("1.0", template["body"])

            attachments.clear()
            attachments_listbox.delete(0, tk.END)
            for file in template["attachments"]:
                attachments.append(file)
                attachments_listbox.insert(tk.END, os.path.basename(file))

        messagebox.showinfo("Template Loaded", f"Template '{template_name}' loaded successfully.")
    except FileNotFoundError:
        messagebox.showerror("Error", f"Template '{template_name}' not found.")


def add_to_history(sender, receiver, subject):
    try:
        with open("email_history.json", "r") as f:
            history = json.load(f)
    except FileNotFoundError:
        history = []

    history.append({
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "sender": sender,
        "receiver": receiver,
        "subject": subject
    })

    with open("email_history.json", "w") as f:
        json.dump(history, f)

    update_history_display()


def update_history_display():
    history_tree.delete(*history_tree.get_children())
    try:
        with open("email_history.json", "r") as f:
            history = json.load(f)
            for item in history:
                history_tree.insert("", "end",
                                    values=(item["timestamp"], item["sender"], item["receiver"], item["subject"]))
    except FileNotFoundError:
        pass


# Initialize attachments list
attachments = []

# Create the main window
root = tk.Tk()
root.title("Email Sender")
root.geometry("700x900")
root.configure(bg=BG_COLOR)

# Create and place widgets
frame = tk.Frame(root, bg=BG_COLOR)
frame.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)

tk.Label(frame, text="Sender Email:", bg=BG_COLOR, font=("Arial", 12)).grid(row=0, column=0, sticky="w")
sender_entry = tk.Entry(frame, width=50)
sender_entry.grid(row=0, column=1, pady=5)

tk.Label(frame, text="Password:", bg=BG_COLOR, font=("Arial", 12)).grid(row=1, column=0, sticky="w")
password_entry = tk.Entry(frame, show="*", width=50)
password_entry.grid(row=1, column=1, pady=5)

tk.Label(frame, text="Receiver Email(s) (comma-separated):", bg=BG_COLOR, font=("Arial", 12)).grid(row=2, column=0,
                                                                                                   sticky="w")
receiver_entry = tk.Entry(frame, width=50)
receiver_entry.grid(row=2, column=1, pady=5)

tk.Label(frame, text="Subject:", bg=BG_COLOR, font=("Arial", 12)).grid(row=3, column=0, sticky="w")
subject_entry = tk.Entry(frame, width=50)
subject_entry.grid(row=3, column=1, pady=5)

tk.Label(frame, text="Body:", bg=BG_COLOR, font=("Arial", 12)).grid(row=4, column=0, sticky="w")
body_text = tk.Text(frame, height=10, width=50)
body_text.grid(row=4, column=1, pady=5)

# Attachments
tk.Label(frame, text="Attachments:", bg=BG_COLOR, font=("Arial", 12)).grid(row=5, column=0, sticky="w")
attachments_listbox = tk.Listbox(frame, height=4, width=50)
attachments_listbox.grid(row=5, column=1, pady=5)
attach_button = tk.Button(frame, text="Attach File", command=attach_file, bg=BUTTON_COLOR_ALT, fg=BUTTON_TEXT_COLOR,
                          relief=tk.RAISED)
attach_button.grid(row=6, column=1, pady=5, sticky="w")

# Save Draft and Load Draft buttons
save_button = tk.Button(frame, text="Save Draft", command=save_draft, bg=BUTTON_COLOR, fg=BUTTON_TEXT_COLOR,
                        relief=tk.RAISED)
save_button.grid(row=6, column=0, pady=5, sticky="e")
load_button = tk.Button(frame, text="Load Draft", command=load_draft, bg=BUTTON_COLOR_ALT, fg=BUTTON_TEXT_COLOR,
                        relief=tk.RAISED)
load_button.grid(row=6, column=1, pady=5, sticky="e")

# Save Template and Load Template buttons
tk.Label(frame, text="Template Name:", bg=BG_COLOR, font=("Arial", 12)).grid(row=7, column=0, sticky="w")
template_name_entry = tk.Entry(frame, width=30)
template_name_entry.grid(row=7, column=1, pady=5, sticky="w")

save_template_button = tk.Button(frame, text="Save Template", command=save_template, bg=BUTTON_COLOR,
                                 fg=BUTTON_TEXT_COLOR, relief=tk.RAISED)
save_template_button.grid(row=8, column=0, pady=5, sticky="e")
load_template_button = tk.Button(frame, text="Load Template", command=load_template, bg=BUTTON_COLOR_ALT,
                                 fg=BUTTON_TEXT_COLOR, relief=tk.RAISED)
load_template_button.grid(row=8, column=1, pady=5, sticky="e")

# Schedule Email
tk.Label(frame, text="Schedule (YYYY-MM-DD or YYYY-MM-DD HH:MM:SS):", bg=BG_COLOR, font=("Arial", 12)).grid(row=9,
                                                                                                            column=0,
                                                                                                            sticky="w")
schedule_entry = tk.Entry(frame, width=50)
schedule_entry.grid(row=9, column=1, pady=5)

schedule_button = tk.Button(frame, text="Schedule Email", command=schedule_email, bg=BUTTON_COLOR_ALT,
                            fg=BUTTON_TEXT_COLOR, relief=tk.RAISED)
schedule_button.grid(row=10, column=1, pady=5, sticky="w")

# Send Email button
send_button = tk.Button(frame, text="Send Email", command=send_email, bg=BUTTON_COLOR, fg=BUTTON_TEXT_COLOR,
                        relief=tk.RAISED)
send_button.grid(row=11, column=1, pady=20, sticky="w")

# Email History
tk.Label(frame, text="Email History:", bg=BG_COLOR, font=("Arial", 12, "bold")).grid(row=12, column=0, columnspan=2,
                                                                                     sticky="w")
history_tree = ttk.Treeview(frame, columns=("Timestamp", "Sender", "Receiver", "Subject"), show="headings")
history_tree.grid(row=13, column=0, columnspan=2, pady=5, sticky="nsew")

history_tree.heading("Timestamp", text="Timestamp")
history_tree.heading("Sender", text="Sender")
history_tree.heading("Receiver", text="Receiver")
history_tree.heading("Subject", text="Subject")

# Configure grid weights
frame.grid_columnconfigure(1, weight=1)
frame.grid_rowconfigure(13, weight=1)

# Initialize the history display
update_history_display()

# Run the application
root.mainloop()
