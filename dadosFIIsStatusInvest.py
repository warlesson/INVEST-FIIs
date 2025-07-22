import time
import os
import glob
import pandas as pd
import threading
import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def coletar_dados():
    try:
        # Caminho da pasta do script
        script_dir = os.path.dirname(os.path.abspath(__file__))

        # Pasta DOWNLOAD dentro da pasta do script
        download_dir = os.path.join(script_dir, "DOWNLOAD")
        os.makedirs(download_dir, exist_ok=True)

        # Configurações do Chrome
        chrome_options = Options()
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })
        chrome_options.add_argument("--window-position=-32000,-32000")  # minimiza fora da tela
        chrome_options.add_argument("--window-size=800,600")

        driver = webdriver.Chrome(options=chrome_options)
        driver.get("https://statusinvest.com.br/fundos-imobiliarios/busca-avancada")

        status_label.config(text="🔍 Clicando em 'Buscar'...")
        wait = WebDriverWait(driver, 20)
        buscar_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.find")))
        buscar_btn.click()

        status_label.config(text="⌛ Aguardando tabela e botão de download...")
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr")))
        time.sleep(2)

        download_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='DOWNLOAD']")))
        download_btn.click()

        status_label.config(text="⬇️ Aguardando download do arquivo CSV...")
        downloaded = False
        for _ in range(20):
            csv_files = glob.glob(os.path.join(download_dir, "*.csv"))
            if csv_files:
                downloaded = True
                break
            time.sleep(1)

        if not downloaded:
            raise Exception("❌ CSV não foi baixado dentro do tempo esperado.")

        csv_path = csv_files[0]
        excel_path = os.path.join(download_dir, "fundos_imobiliarios_statusinvest.xlsx")

        df = pd.read_csv(csv_path, sep=";", encoding="utf-8")
        df.to_excel(excel_path, index=False)

        os.remove(csv_path)
        driver.quit()

        status_label.config(text="✅ Planilha gerada com sucesso!")
        messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{excel_path}")

    except Exception as e:
        status_label.config(text="❌ Erro ao coletar dados.")
        messagebox.showerror("Erro", f"Erro ao coletar dados:\n{e}")

def iniciar_thread():
    threading.Thread(target=coletar_dados).start()

# Interface Tkinter
janela = tk.Tk()
janela.title("Coletor de FIIs - StatusInvest")
janela.geometry("420x180")

titulo = tk.Label(janela, text="Atualizar FIIs do StatusInvest", font=("Arial", 16))
titulo.pack(pady=10)

botao = tk.Button(janela, text="🔄 Atualizar Dados", font=("Arial", 12), command=iniciar_thread)
botao.pack(pady=10)

status_label = tk.Label(janela, text="", font=("Arial", 10))
status_label.pack(pady=5)

janela.mainloop()
