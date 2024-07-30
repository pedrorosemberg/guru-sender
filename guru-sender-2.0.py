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
import string

# Configure o logger para escrever em um arquivo .doc
logging.basicConfig(filename='guru_sender.doc', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Palavras proibidas para validação de mensagens
forbidden_words = [
    'drogas', 'maconha', 'cigarro', 'tabaco', 'álcool', 'esteroides', 'armas',
    'munição', 'explosivos', 'animais', 'sexo', 'pornografia', 'jogos de azar',
    'encontros', 'pirataria', 'moeda falsa'
]

# Função para validar números de telefone
def validate_phone(number):
    try:
        parsed_number = phonenumbers.parse(number, "BR")  # Assumindo Brasil como padrão
        return phonenumbers.is_valid_number(parsed_number)
    except phonenumbers.NumberParseException:
        return False

# Função para ler o arquivo Excel
def read_excel(file_path):
    df = pd.read_excel(file_path)
    return df

# Função para verificar a conformidade com a política comercial
def check_compliance(message):
    for word in forbidden_words:
        if word in message.lower():
            return False, f"Message contains forbidden content: {word}"
    return True, ""

# Função para enviar mensagem via WhatsApp
def send_message(number, message):
    try:
        url = f"https://wa.me/{number}?text={urllib.parse.quote(message)}"
        webbrowser.open(url)
        return True
    except Exception as e:
        logging.error(f"Error sending message to {number}: {e}")
        return False

# Função para escolher um arquivo Excel
def choose_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    entry_file.delete(0, tk.END)
    entry_file.insert(0, file_path)

# Função para alternar entre modos claro e escuro
def toggle_theme():
    global dark_mode
    dark_mode = not dark_mode

    bg_color = "black" if dark_mode else "white"
    fg_color = "white" if dark_mode else "black"
    entry_bg = "gray20" if dark_mode else "white"
    entry_fg = "white" if dark_mode else "black"

    # Atualiza o estilo para ttk
    style.configure("TLabel", background=bg_color, foreground=fg_color)
    style.configure("TButton", background=bg_color, foreground=fg_color)
    style.configure("TEntry", fieldbackground=entry_bg, foreground=entry_fg)
    style.configure("TText", fieldbackground=entry_bg, foreground=entry_fg)
    style.configure("TFrame", background=bg_color)

    # Atualiza widgets do tkinter
    root.configure(bg=bg_color)
    logo_label.configure(bg=bg_color, fg=fg_color)
    footer.configure(bg=bg_color, fg=fg_color)
    
    # Atualiza cores dos widgets não-ttk
    for widget in [text_message, text_numbers, log_text]:
        widget.configure(bg=entry_bg, fg=entry_fg)
    
    # Atualiza log_frame background
    log_frame.configure(bg=bg_color)

# Função para iniciar o envio de mensagens
def start_sending():
    btn_start.config(state=tk.DISABLED)
    file_path = entry_file.get()
    base_message = text_message.get("1.0", tk.END).strip()
    manual_numbers = text_numbers.get("1.0", tk.END).strip().split('\n')

    if not (file_path or manual_numbers) or not base_message:
        messagebox.showerror("Error", "Please provide the file path, manual numbers, or the message.")
        return

    # Verifica se todas as chaves de formatação estão presentes
    keys = [part[1] for part in string.Formatter().parse(base_message) if part[1] is not None]
    if 'nome' not in keys:
        messagebox.showerror("Error", "The message must contain the key {nome}.")
        return

    # Lista para armazenar dados dos destinatários
    recipients = []

    # Adiciona números e nomes da planilha, se disponível
    if file_path:
        try:
            df = read_excel(file_path)
            recipients.extend([(str(row['telefone']), row['nome']) for _, row in df.iterrows()])
        except Exception as e:
            messagebox.showerror("Error", f"Error reading the Excel file: {e}")
            return

    # Adiciona números e nomes manuais
    for number in manual_numbers:
        if number.strip():
            recipients.append((number.strip(), "Cliente"))

    total_messages = len(recipients)
    successful_sends = 0
    sending_errors = 0

    for number, name in recipients:
        message = base_message.format(nome=name)

        # Valida número de telefone
        if not validate_phone(number):
            log_message(f"Invalid phone number: {number}", success=False)
            sending_errors += 1
            continue

        # Verifica conformidade com a política comercial
        compliant, error_message = check_compliance(message)
        if not compliant:
            log_message(f"Error: {error_message} for message to {name}", success=False)
            sending_errors += 1
            continue

        # Simula envio e verifica sucesso
        if send_message(number, message):
            successful_sends += 1
            log_message(f"Message sent successfully to {name}", success=True)
        else:
            sending_errors += 1

        # Atualiza a barra de progresso
        progress.set((successful_sends + sending_errors) / total_messages * 100)

        # Atualiza o log na interface gráfica
        log_message(f"Progress: {successful_sends}/{total_messages} sent", success=True)

        # Espera aleatória entre 30 e 90 segundos
        time.sleep(random.randint(30, 90))

    messagebox.showinfo("Completed", f"Message sending completed.\nSuccessfully sent: {successful_sends}\nErrors: {sending_errors}")
    log_message("Message sending completed.", success=True)
    btn_start.config(state=tk.NORMAL)

# Função para registrar mensagens tanto no arquivo quanto na interface gráfica
def log_message(message, success=True):
    logging.info(message)
    log_text.config(state=tk.NORMAL)
    color = "blue" if success else "red"
    log_text.insert('1.0', message + '\n', color)
    log_text.config(state=tk.DISABLED)

# Criação da janela de configurações
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

# Criação da interface principal
root = tk.Tk()
root.title("guru-sender")

# Variável para controlar o modo escuro
dark_mode = True

# Estilo para temas
style = ttk.Style(root)

# Logotipo no canto superior esquerdo
logo_label = tk.Label(root, text="CODEVER", font=("Arial", 10, "bold"))
logo_label.grid(row=0, column=0, padx=10, pady=10, sticky='nw')

tk.Label(root, text="Excel File:").grid(row=1, column=0, padx=10, pady=10)
entry_file = tk.Entry(root, width=50)
entry_file.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Choose File", command=choose_file).grid(row=1, column=2, padx=10, pady=10)

tk.Label(root, text="Manual Numbers (one per line):").grid(row=2, column=0, padx=10, pady=10)
text_numbers = tk.Text(root, height=5, width=50)
text_numbers.grid(row=2, column=1, padx=10,pady=10, columnspan=2)

tk.Label(root, text="Message:").grid(row=3, column=0, padx=10, pady=10)
text_message = tk.Text(root, height=10, width=50)
text_message.grid(row=3, column=1, padx=10, pady=10, columnspan=2)

# Barra de progresso
progress = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress, maximum=100)
progress_bar.grid(row=4, column=0, columnspan=3, padx=10, pady=10, sticky='ew')

# Área de texto para o log
log_frame = tk.Frame(root)
log_frame.grid(row=5, column=0, columnspan=3, padx=10, pady=10, sticky='ew')

log_text = tk.Text(log_frame, height=10, width=80, state='disabled')
log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Definindo tags para colorir o texto no log
log_text.tag_configure("blue", foreground="blue")
log_text.tag_configure("red", foreground="red")

log_scrollbar = tk.Scrollbar(log_frame, command=log_text.yview)
log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

log_text.config(yscrollcommand=log_scrollbar.set)

# Botões
btn_start = tk.Button(root, text="Start Sending", command=lambda: threading.Thread(target=start_sending).start())
btn_start.grid(row=6, column=1, padx=10, pady=10)

tk.Button(root, text="Settings", command=open_settings).grid(row=6, column=2, padx=10, pady=10)
tk.Button(root, text="Toggle Theme", command=toggle_theme).grid(row=6, column=0, padx=10, pady=10)

# Rodapé de direitos autorais
footer = tk.Label(root, text="Copyright 2024 by Codever - Pedro Rosemberg. All rights reserved. Just for domestic or academic use only, no commercial use authorized.", font=("Montserrat", 8))
footer.grid(row=7, column=0, columnspan=3, pady=(10, 0))

# Inicializa o tema
toggle_theme()

root.mainloop()
