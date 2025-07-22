import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
import threading
import os

def coletar_fiis():
    try:
        # Caminho para salvar na pasta DOWNLOAD dentro do diretório atual
        pasta_download = os.path.join(os.path.dirname(__file__), "DOWNLOAD")
        os.makedirs(pasta_download, exist_ok=True)
        caminho_excel = os.path.join(pasta_download, "fundos_imobiliarios_explorer.xlsx")

        # Configurações do navegador
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")

        driver = webdriver.Chrome(options=chrome_options)
        driver.get("https://www.fundsexplorer.com.br/ranking")

        status_label.config(text="🔄 Coletando dados...")
        time.sleep(15)  # Aguarda o carregamento da tabela

        tabelas = pd.read_html(driver.page_source, decimal=",", thousands=".")
        driver.quit()

        if not tabelas:
            status_label.config(text="❌ Nenhuma tabela encontrada.")
            return

        df = tabelas[0]
        df.to_excel(caminho_excel, index=False)

        status_label.config(text="✅ Dados atualizados com sucesso!")
        messagebox.showinfo("Sucesso", f"Planilha salva em: {caminho_excel}")

    except Exception as e:
        status_label.config(text="❌ Erro na atualização.")
        messagebox.showerror("Erro", f"Erro ao coletar dados:\n{e}")

def iniciar_coleta():
    threading.Thread(target=coletar_fiis).start()

# Interface Tkinter
janela = tk.Tk()
janela.title("Atualizador de FIIs")
janela.geometry("400x180")

titulo = tk.Label(janela, text="Atualizar Planilha de FIIs", font=("Arial", 16))
titulo.pack(pady=10)

btn = tk.Button(janela, text="🔄 Atualizar Dados", font=("Arial", 12), command=iniciar_coleta)
btn.pack(pady=10)

status_label = tk.Label(janela, text="", font=("Arial", 10))
status_label.pack(pady=5)

janela.mainloop()
