import pandas as pd
import webbrowser
import time
import random
import urllib.parse
import tkinter as tk
from tkinter import filedialog, messagebox

# Lista de palavras-chave proibidas
palavras_proibidas = [
    'drogas', 'maconha', 'cigarro', 'tabaco', 'álcool', 'esteroides', 'armas', 
    'munição', 'explosivos', 'animais', 'sexo', 'pornografia', 'jogos de azar', 
    'encontros', 'pirataria', 'moeda falsa'
]

# Leitura do arquivo Excel
def ler_excel(file_path):
    df = pd.read_excel(file_path)
    return df

# Verificação de conformidade com a política comercial
def verificar_politica_comercial(mensagem):
    for palavra in palavras_proibidas:
        if palavra in mensagem.lower():
            return False, f"Mensagem contém conteúdo proibido: {palavra}"
    return True, ""

# Enviar mensagem via WhatsApp
def enviar_mensagem(numero, mensagem):
    try:
        url = f"https://wa.me/{numero}?text={urllib.parse.quote(mensagem)}"
        webbrowser.open(url)
        return True
    except Exception as e:
        print(f"Erro ao enviar mensagem para {numero}: {e}")
        return False

# Função para escolher o arquivo Excel
def escolher_arquivo():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    entry_arquivo.delete(0, tk.END)
    entry_arquivo.insert(0, file_path)

# Função para iniciar o envio das mensagens
def iniciar_envio():
    file_path = entry_arquivo.get()
    mensagem_base = text_mensagem.get("1.0", tk.END).strip()

    if not file_path or not mensagem_base:
        messagebox.showerror("Erro", "Por favor, forneça o caminho do arquivo e a mensagem.")
        return

    try:
        df = ler_excel(file_path)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o arquivo Excel: {e}")
        return

    total_mensagens = len(df)
    enviadas_com_sucesso = 0
    erros_no_envio = 0

    for index, row in df.iterrows():
        numero = row['telefone']
        nome = row['nome']
        mensagem = mensagem_base.format(nome=nome)
        
        # Verificar conformidade com a política comercial
        conforme, erro = verificar_politica_comercial(mensagem)
        if not conforme:
            print(f"Erro: {erro} na mensagem para {nome}")
            erros_no_envio += 1
            continue

        if enviar_mensagem(numero, mensagem):
            enviadas_com_sucesso += 1
        else:
            erros_no_envio += 1
        
        # Espera aleatória entre 30 a 90 segundos
        time.sleep(random.randint(30, 90))

    messagebox.showinfo("Concluído", f"Envio de mensagens concluído.\nEnviadas com sucesso: {enviadas_com_sucesso}\nErros no envio: {erros_no_envio}")
    print("Envio de mensagens concluído.")

# Criação da interface gráfica
root = tk.Tk()
root.title("guru-sender")

tk.Label(root, text="Arquivo Excel:").grid(row=0, column=0, padx=10, pady=10)
entry_arquivo = tk.Entry(root, width=50)
entry_arquivo.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Escolher Arquivo", command=escolher_arquivo).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Mensagem:").grid(row=1, column=0, padx=10, pady=10)
text_mensagem = tk.Text(root, height=10, width=50)
text_mensagem.grid(row=1, column=1, padx=10, pady=10, columnspan=2)

tk.Button(root, text="Iniciar Envio", command=iniciar_envio).grid(row=2, column=1, padx=10, pady=10, columnspan=2)

root.mainloop()