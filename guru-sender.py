import pandas as pd
import webbrowser
import time
import random
import urllib.parse
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import phonenumbers
import logging
import threading
import string  # Import string for formatting checks

# Configure the logger
logging.basicConfig(filename='guru_sender.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# List of forbidden keywords
forbidden_words = [
    'drogas', 'maconha', 'cigarro', 'tabaco', 'álcool', 'esteroides', 'armas',
    'munição', 'explosivos', 'animais', 'sexo', 'pornografia', 'jogos de azar',
    'encontros', 'pirataria', 'moeda falsa'
]

# Function to validate phone numbers
def validate_phone(number):
    try:
        parsed_number = phonenumbers.parse(number, "BR")  # Assuming Brazil as default
        return phonenumbers.is_valid_number(parsed_number)
    except phonenumbers.NumberParseException:
        return False

# Function to read the Excel file
def read_excel(file_path):
    df = pd.read_excel(file_path)
    return df

# Function to check compliance with the commercial policy
def check_compliance(message):
    for word in forbidden_words:
        if word in message.lower():
            return False, f"Message contains forbidden content: {word}"
    return True, ""

# Function to send a message via WhatsApp
def send_message(number, message):
    try:
        url = f"https://wa.me/{number}?text={urllib.parse.quote(message)}"
        webbrowser.open(url)
        return True
    except Exception as e:
        logging.error(f"Error sending message to {number}: {e}")
        return False

# Function to choose an Excel file
def choose_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    entry_file.delete(0, tk.END)
    entry_file.insert(0, file_path)

# Function to start sending messages
def start_sending():
    btn_start.config(state=tk.DISABLED)
    file_path = entry_file.get()
    base_message = text_message.get("1.0", tk.END).strip()

    if not file_path or not base_message:
        messagebox.showerror("Error", "Please provide the file path and the message.")
        return

    # Check if all formatting keys are present
    keys = [part[1] for part in string.Formatter().parse(base_message) if part[1] is not None]
    if 'nome' not in keys:
        messagebox.showerror("Error", "The message must contain the key {nome}.")
        return

    try:
        df = read_excel(file_path)
    except Exception as e:
        messagebox.showerror("Error", f"Error reading the Excel file: {e}")
        return

    total_messages = len(df)
    successful_sends = 0
    sending_errors = 0

    for index, row in df.iterrows():
        number = str(row['telefone'])
        name = row['nome']
        message = base_message.format(nome=name)

        # Validate phone number
        if not validate_phone(number):
            log_message(f"Invalid phone number: {number}")
            sending_errors += 1
            continue

        # Check compliance with the commercial policy
        compliant, error_message = check_compliance(message)
        if not compliant:
            log_message(f"Error: {error_message} for message to {name}")
            sending_errors += 1
            continue

        # Simulate sending and check success
        if send_message(number, message):
            # Assume sending was successful if we can open the link
            successful_sends += 1
            log_message(f"Message sent successfully to {name}")
        else:
            sending_errors += 1

        # Update progress bar
        progress.set((successful_sends + sending_errors) / total_messages * 100)

        # Update log in the GUI
        log_message(f"Progress: {successful_sends}/{total_messages} sent")

        # Random wait between 30 to 90 seconds
        time.sleep(random.randint(30, 90))

    messagebox.showinfo("Completed", f"Message sending completed.\nSuccessfully sent: {successful_sends}\nErrors: {sending_errors}")
    log_message("Message sending completed.")
    btn_start.config(state=tk.NORMAL)

# Function to log messages both in the file and in the GUI
def log_message(message):
    logging.info(message)
    log_text.insert(tk.END, message + '\n')
    log_text.see(tk.END)  # Scroll to the end

# Create the settings window
def open_settings():
    settings_window = tk.Toplevel(root)
    settings_window.title("Settings")

    tk.Label(settings_window, text="Forbidden Words (comma-separated):").pack(padx=10, pady=10)
    forbidden_words_entry = tk.Entry(settings_window, width=50)
    forbidden_words_entry.insert(0, ', '.join(forbidden_words))
    forbidden_words_entry.pack(padx=10, pady=10)

    tk.Label(settings_window, text="Wait Interval (seconds, e.g., 30-90):").pack(padx=10, pady=10)
    interval_entry = tk.Entry(settings_window, width=20)
    interval_entry.insert(0, "30-90")
    interval_entry.pack(padx=10, pady=10)

    def save_settings():
        global forbidden_words
        forbidden_words = [word.strip() for word in forbidden_words_entry.get().split(',')]
        messagebox.showinfo("Settings", "Settings saved successfully.")
        settings_window.destroy()

    tk.Button(settings_window, text="Save Settings", command=save_settings).pack(pady=20)

# Create the main interface
root = tk.Tk()
root.title("guru-sender")

tk.Label(root, text="Excel File:").grid(row=0, column=0, padx=10, pady=10)
entry_file = tk.Entry(root, width=50)
entry_file.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Choose File", command=choose_file).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Message:").grid(row=1, column=0, padx=10, pady=10)
text_message = tk.Text(root, height=10, width=50)
text_message.grid(row=1, column=1, padx=10, pady=10, columnspan=2)

# Progress bar
progress = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress, maximum=100)
progress_bar.grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky='ew')

# Log text area
log_frame = tk.Frame(root)
log_frame.grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky='ew')

log_text = tk.Text(log_frame, height=10, width=80, state='normal')
log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

log_scrollbar = tk.Scrollbar(log_frame, command=log_text.yview)
log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

log_text.config(yscrollcommand=log_scrollbar.set)

# Buttons
btn_start = tk.Button(root, text="Start Sending", command=lambda: threading.Thread(target=start_sending).start())
btn_start.grid(row=4, column=1, padx=10, pady=10)

tk.Button(root, text="Settings", command=open_settings).grid(row=4, column=2, padx=10, pady=10)

# Copyright footer
footer = tk.Label(root, text="Copyright 2024 by Codever, All rights reserved. For domestic and academic use only. No commercial use authorized.", font=("Arial", 8))
footer.grid(row=5, column=0, columnspan=3, pady=(10, 0))

root.mainloop()