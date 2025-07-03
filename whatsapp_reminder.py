import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import pywhatkit
import pyautogui
import time
import os
import threading
import logging
import re
from datetime import datetime

# --- Konfigurasi Logging ---
if not os.path.exists("logs"):
    os.makedirs("logs")

log_filename = f"logs/activity_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    filename=log_filename,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

# --- Tema CustomTkinter ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

# --- Event untuk pembatalan proses ---
cancel_event = threading.Event()

# --- Fungsi Utilitas ---

def get_excel_file():
    filepath = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
    )
    return filepath

def load_excel(filepath):
    return pd.read_excel(filepath, dtype="string")

def check_required_columns(df):
    required = ['Name', 'Phone', 'Alias', 'Nominal', 'Saving']
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"Missing columns: {', '.join(missing)}")

def load_template(path="reminder_text.txt"):
    if not os.path.exists(path):
        raise FileNotFoundError("Reminder text file not found.")
    with open(path, "r") as f:
        return f.read()

def validate_phone(phone):
    return bool(re.match(r"^\+\d{10,15}$", phone))

def personalize_message(template, row):
    return template.format(
        Name=row["Name"],
        Alias=row["Alias"],
        Nominal=row["Nominal"],
        Saving=row["Saving"]
    )

def send_whatsapp_message(phone, message):
    pywhatkit.sendwhatmsg_instantly(phone_no=phone, message=message)
    time.sleep(2)
    pyautogui.press("enter")
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'w')

# --- Fungsi Utama ---

def process_messages():
    cancel_event.clear()
    try:
        filepath = get_excel_file()
        if not filepath:
            status_label.configure(text="No file selected")
            logging.info("No file selected")
            return

        df = load_excel(filepath)
        check_required_columns(df)
        template = load_template()

        if not messagebox.askyesno("Confirmation", "Do you want to send the messages?"):
            logging.info("User cancelled message sending")
            return

        for _, row in df.iterrows():
            if cancel_event.is_set():
                status_label.configure(text="Process cancelled")
                logging.info("Process was cancelled by user")
                break

            name, phone = row["Name"], row["Phone"]
            if not validate_phone(phone):
                status_label.configure(text=f"Invalid phone number: {phone}")
                logging.warning(f"Invalid phone number: {phone}")
                continue

            try:
                message = personalize_message(template, row)
                status_label.configure(text=f"Sending to {name}...")
                window.update()

                send_whatsapp_message(phone, message)
                logging.info(f"Successfully sent to {name} ({phone})")
            except Exception as e:
                status_label.configure(text=f"Error sending to {name}")
                logging.error(f"Failed to send to {name}: {e}")

        if not cancel_event.is_set():
            messagebox.showinfo("Success", "All messages have been processed!")
            status_label.configure(text="Messages sent successfully")

    except Exception as e:
        logging.exception("Unexpected error occurred")
        messagebox.showerror("Error", f"An error occurred: {e}")
        status_label.configure(text="Error occurred during process")
    finally:
        logging.info("Activity log saved.")
        messagebox.showinfo("Log Saved", f"Activity log has been saved to '{log_filename}'")

def cancel_process():
    cancel_event.set()
    status_label.configure(text="Cancelling process...")
    logging.info("User requested process cancellation")

def toggle_mode():
    mode = mode_switch.get()
    if mode:
        ctk.set_appearance_mode("Light")
        mode_switch.configure(text="Dark Mode")
    else:
        ctk.set_appearance_mode("Dark")
        mode_switch.configure(text="Light Mode")

# --- GUI ---
window = ctk.CTk()
window.title("WhatsApp Reminder System")
window.geometry("500x300")

header_label = ctk.CTkLabel(window, text="WhatsApp Reminder System", font=("Arial", 20))
header_label.pack(pady=10)

content_frame = ctk.CTkFrame(window, corner_radius=10)
content_frame.pack(padx=20, pady=10, fill="both", expand=True)

instruction_label = ctk.CTkLabel(content_frame, text="Select an Excel file to send reminders:")
instruction_label.pack(pady=10)

send_button = ctk.CTkButton(
    content_frame, text="Select File & Send",
    command=lambda: threading.Thread(target=process_messages).start()
)
send_button.pack(pady=10)

cancel_button = ctk.CTkButton(
    content_frame,
    text="Cancel",
    command=cancel_process,
    fg_color="#CC2B52",
    hover_color="#AF1740"
)
cancel_button.pack(pady=10)

status_label = ctk.CTkLabel(content_frame, text="", font=("Arial", 12, "italic"))
status_label.pack(pady=(20, 0))

mode_switch = ctk.CTkSwitch(window, text="Light Mode", command=toggle_mode)
mode_switch.pack(side="bottom", pady=10)

# --- Run ---
window.mainloop()
