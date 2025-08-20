# main.py
import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
from openpyxl import Workbook
from datetime import datetime
import os
import time

class WeatherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Captador de Temperatura de São Paulo")
        self.root.geometry("400x200")

        self.create_widgets()

    def create_widgets(self):
        # Frame principal
        main_frame = tk.Frame(self.root)
        main_frame.pack(expand=True)

        tk.Label(main_frame, text="Previsão do tempo de São Paulo", font=("Arial", 16)).pack(pady=10)

        # Botão para buscar a previsão
        tk.Button(main_frame, text="Buscar Previsão", command=self.get_weather_data, font=("Arial", 12)).pack(pady=10)

        # Rótulo de status
        self.status_label = tk.Label(main_frame, text="", font=("Arial", 10))
        self.status_label.pack(pady=5)

    def get_weather_data(self):
        self.status_label.config(text="Buscando informações. Aguarde...")
        self.root.update_idletasks()

        try:
            driver = webdriver.Chrome()
            driver.get("https://www.google.com/search?q=previsão+do+tempo+são+paulo")

            wait = WebDriverWait(driver, 10)
            
            # Tentativa 1: Seletores robustos
            try:
                temperature_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".wob_t.q8U8x")))
                humidity_element = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='wob_hm']")))
                temperature = temperature_element.text
                humidity = humidity_element.text
            except Exception:
                # Tentativa 2: Seletores alternativos
                try:
                    temperature_element = driver.find_element(By.CSS_SELECTOR, ".wob_t")
                    humidity_element = driver.find_element(By.XPATH, "//*[contains(text(), 'Umidade:')]")
                    temperature = temperature_element.text
                    humidity = humidity_element.text.replace("Umidade: ", "")
                except Exception:
                    raise Exception("Não foi possível encontrar os elementos de temperatura e umidade na página.")

            driver.quit()
            self.save_to_excel(temperature, humidity)
            self.status_label.config(text="Dados capturados e salvos com sucesso!")
            messagebox.showinfo("Sucesso", "Dados capturados e salvos na planilha!")

        except Exception as e:
            self.status_label.config(text="Erro ao buscar ou salvar os dados.")
            messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

    def save_to_excel(self, temperature, humidity):
        file_name = "dados_previsao.xlsx"
        
        if os.path.exists(file_name):
            workbook = openpyxl.load_workbook(file_name)
            sheet = workbook.active
        else:
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Data/Hora", "Temperatura", "Status da Umidade do Ar"])

        now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        sheet.append([now, temperature, humidity])
        
        workbook.save(file_name)

def main():
    root = tk.Tk()
    app = WeatherApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
