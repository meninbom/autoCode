import customtkinter as ctk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
import pandas as pd
import json
import os
import time
import threading
import re
import logging
import requests
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Border, Side

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

class App(ctk.CTk):
    responsible_list = [
        "Cintia Aparecida Zeinade",
        "Diego Junior Emerenciano",
        "Jean Michel de Oliveira",
        "Lucas Dessunti",
        "Rafael de Araujo Vasilio",
        "Fabio Augusto Souza Munhoz",
        "Erik Rubens Bortotto Chaiben",
        "Carlos Henrique Moreira Diniz",
        "Paulo Victor Arroteia de Albuquerque",
        "Marcos Vinicius Orias Costa",
        "Kauã de Melo Silva"
    ]

    def __init__(self):
        super().__init__()
        self.title("Gerador de Links SGD")
        self.geometry("500x500")
        self.minsize(400, 750)
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self.config = self.load_config()
        
        self.create_widgets()
        self.populate_fields()
    
    def create_widgets(self):
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        
        self.lbl_username = ctk.CTkLabel(self.main_frame, text="Usuário:")
        self.lbl_username.pack(fill="x", padx=10, pady=5)
        self.entry_username = ctk.CTkEntry(self.main_frame)
        self.entry_username.pack(fill="x", padx=10, pady=5)
        
        self.lbl_password = ctk.CTkLabel(self.main_frame, text="Senha:")
        self.lbl_password.pack(fill="x", padx=10, pady=5)
        self.entry_password = ctk.CTkEntry(self.main_frame, show="*")
        self.entry_password.pack(fill="x", padx=10, pady=5)
        
        self.lbl_directory = ctk.CTkLabel(self.main_frame, text="Diretório de Download:")
        self.lbl_directory.pack(fill="x", padx=10, pady=5)
        self.entry_directory = ctk.CTkEntry(self.main_frame)
        self.entry_directory.pack(fill="x", padx=10, pady=5)
        self.btn_browse = ctk.CTkButton(self.main_frame, text="Procurar", command=self.browse_directory)
        self.btn_browse.pack(fill="x", padx=10, pady=5)
        
        self.lbl_responsible = ctk.CTkLabel(self.main_frame, text="Responsável:")
        self.lbl_responsible.pack(fill="x", padx=10, pady=5)
        self.dropdown_responsible = ctk.CTkComboBox(self.main_frame, values=self.responsible_list)
        self.dropdown_responsible.pack(fill="x", padx=10, pady=5)
        
        self.lbl_start_date = ctk.CTkLabel(self.main_frame, text="Data Inicial (dd/mm/aa):")
        self.lbl_start_date.pack(fill="x", padx=10, pady=5)
        self.entry_start_date = ctk.CTkEntry(self.main_frame)
        self.entry_start_date.pack(fill="x", padx=10, pady=5)
        
        self.lbl_end_date = ctk.CTkLabel(self.main_frame, text="Data Final (dd/mm/aa):")
        self.lbl_end_date.pack(fill="x", padx=10, pady=5)
        self.entry_end_date = ctk.CTkEntry(self.main_frame)
        self.entry_end_date.pack(fill="x", padx=10, pady=5)
        
        self.progress_label = ctk.CTkLabel(self.main_frame, text="Pronto para iniciar")
        self.progress_label.pack(fill="x", padx=10, pady=5)
        
        self.progress_bar = ctk.CTkProgressBar(self.main_frame)
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", padx=10, pady=5)
        
        self.btn_execute = ctk.CTkButton(self.main_frame, text="Executar", command=self.execute)
        self.btn_execute.pack(fill="x", padx=10, pady=10)
        
        self.btn_save_config = ctk.CTkButton(self.main_frame, text="Salvar Configuração", command=self.save_config)
        self.btn_save_config.pack(fill="x", padx=10, pady=10)
        
        self.btn_exit = ctk.CTkButton(self.main_frame, text="Sair", command=self.quit)
        self.btn_exit.pack(fill="x", padx=10, pady=10)
    
    def browse_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.entry_directory.delete(0, ctk.END)
            self.entry_directory.insert(0, directory)
    
    def load_config(self):
        try:
            with open('config.json', 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            default_config = {
                "usuario": "",
                "senha": "",
                "diretorio_download": "",
                "responsavel": self.responsible_list[0],
                "data_inicial": "",
                "data_final": ""
            }
            with open('config.json', 'w') as f:
                json.dump(default_config, f)
            return default_config
    
    def populate_fields(self):
        if 'usuario' in self.config:
            self.entry_username.insert(0, self.config['usuario'])
        if 'senha' in self.config:
            self.entry_password.insert(0, self.config['senha'])
        if 'diretorio_download' in self.config:
            self.entry_directory.insert(0, self.config['diretorio_download'])
        if 'responsavel' in self.config and self.config['responsavel'] in self.responsible_list:
            self.dropdown_responsible.set(self.config['responsavel'])
        else:
            self.dropdown_responsible.set(self.responsible_list[0])
        if 'data_inicial' in self.config:
            self.entry_start_date.insert(0, self.config['data_inicial'])
        if 'data_final' in self.config:
            self.entry_end_date.insert(0, self.config['data_final'])
    
    def save_config(self):
        config = {
            'usuario': self.entry_username.get(),
            'senha': self.entry_password.get(),
            'diretorio_download': self.entry_directory.get(),
            'responsavel': self.dropdown_responsible.get(),
            'data_inicial': self.entry_start_date.get(),
            'data_final': self.entry_end_date.get()
        }
        with open('config.json', 'w') as f:
            json.dump(config, f)
        messagebox.showinfo("Configuração", "Configuração salva com sucesso!")
    
    def validate_date(self, date_str):
        pattern = r"^\d{2}/\d{2}/\d{2}$"
        return bool(re.match(pattern, date_str))
    
    def get_unique_filename(self, directory, base_name):
        file_path = os.path.join(directory, base_name)
        counter = 1
        while os.path.exists(file_path):
            name, ext = os.path.splitext(base_name)
            new_name = f"{name}_{counter}{ext}"
            file_path = os.path.join(directory, new_name)
            counter += 1
        return file_path
    
    def update_progress(self, progress, message):
        self.progress_bar.set(progress / 100)
        self.progress_label.configure(text=message)
        self.update_idletasks()
    
    def thread_safe_update(self, progress, message):
        self.after(0, self.update_progress, progress, message)
    
    def execute(self):
        self.btn_execute.configure(state="disabled")
        threading.Thread(target=self._execute, daemon=True).start()
    
    def _execute(self):
        username = self.entry_username.get()
        password = self.entry_password.get()
        directory = self.entry_directory.get()
        responsible = self.dropdown_responsible.get()
        start_date = self.entry_start_date.get()
        end_date = self.entry_end_date.get()
        
        if not all([username, password, directory, responsible, start_date, end_date]):
            self.thread_safe_update(0, "Erro")
            messagebox.showerror("Erro", "Todos os campos são obrigatórios.")
            self.after(0, lambda: self.btn_execute.configure(state="normal"))
            return
        
        if not (self.validate_date(start_date) and self.validate_date(end_date)):
            self.thread_safe_update(0, "Erro")
            messagebox.showerror("Erro", "Datas devem estar no formato dd/mm/aa.")
            self.after(0, lambda: self.btn_execute.configure(state="normal"))
            return
        
        options = webdriver.ChromeOptions()
        options.add_argument("--start-minimized")
        options.add_experimental_option("prefs", {
            "download.default_directory": directory,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })
        
        try:
            logger.info("Inicializando ChromeDriver...")
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=options)
            logger.info("ChromeDriver inicializado com sucesso.")
            driver.minimize_window()
            
            self.thread_safe_update(20, "Acessando sistema...")
            driver.get("https://sgd.dominiosistemas.com.br")
            
            logger.info("Aguardando campos de login...")
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/form/input[2]')))
            user_input = driver.find_element(By.XPATH, '/html/body/div/form/input[2]')
            pwd_input = driver.find_element(By.XPATH, '/html/body/div/form/input[3]')
            login_btn = driver.find_element(By.XPATH, '/html/body/div/form/input[4]')
            
            logger.info("Preenchendo credenciais...")
            user_input.send_keys(username)
            pwd_input.send_keys(password)
            login_btn.click()
            
            WebDriverWait(driver, 10).until(EC.url_changes("https://sgd.dominiosistemas.com.br"))
            
            self.thread_safe_update(30, "Acessando página de programações...")
            url_programacoes = "https://sgd.dominiosistemas.com.br/sgsc/faces/programacoes.html"
            logger.info(f"Acessando URL de programações: {url_programacoes}")
            driver.get(url_programacoes)
            
            self.thread_safe_update(35, "Configurando datas...")
            logger.info("Preenchendo data inicial...")
            data_inicio_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="programacoesForm:agendadoDe"]'))
            )
            data_inicio_input.clear()
            data_inicio_input.click()
            data_inicio_input.send_keys(Keys.CONTROL + "a")
            data_inicio_input.send_keys(Keys.DELETE)
            data_inicio_input.send_keys(start_date)
            data_inicio_input.send_keys(Keys.TAB)
            
            logger.info("Preenchendo data final...")
            data_fim_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="programacoesForm:agendadoAte"]'))
            )
            data_fim_input.clear()
            data_fim_input.click()
            data_fim_input.send_keys(Keys.CONTROL + "a")
            data_fim_input.send_keys(Keys.DELETE)
            data_fim_input.send_keys(end_date)
            data_fim_input.send_keys(Keys.TAB)
            
            self.thread_safe_update(40, "Salvando configurações de datas...")
            logger.info("Clicando no botão de salvar...")
            salvar_btn = driver.find_element(By.XPATH, '//*[@id="programacoesForm:atualizarBtn"]')
            salvar_btn.click()
            time.sleep(5)
            
            self.thread_safe_update(45, "Gerando relatório...")
            logger.info("Clicando no botão de gerar relatório...")
            gerarelatorio_btn = driver.find_element(By.XPATH, '//*[@id="programacoesForm:gerarRelatorio"]')
            gerarelatorio_btn.click()
            time.sleep(5)
            
            self.thread_safe_update(50, "Aguardando link de download...")
            logger.info("Aguardando link de download...")
            download_link = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="downloadBoxID"]/table/tbody/tr[5]/td/a'))
            )
            
            download_url = download_link.get_attribute('href')
            logger.info(f"URL de download obtido: {download_url}")
            cookies = {cookie['name']: cookie['value'] for cookie in driver.get_cookies()}
            
            self.thread_safe_update(60, "Baixando arquivo...")
            logger.info("Baixando arquivo via requests...")
            response = requests.get(download_url, cookies=cookies, stream=True)
            
            if response.status_code != 200:
                raise Exception(f"Falha ao baixar o arquivo: Status {response.status_code}")
            
            # Extract filename from URL or use a default
            filename = download_url.split('/')[-1] if '/' in download_url else "relatorio.xlsx"
            raw_file_path = os.path.join(directory, filename)
            
            # Save the downloaded file temporarily
            with open(raw_file_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            logger.info(f"Arquivo baixado: {raw_file_path}")
            
            # Rename the file to YYYY-MM-DD.xlsx with suffix if needed
            current_date = datetime.now().strftime("%Y-%m-%d")
            base_filename = f"{current_date}.xlsx"
            processed_file = self.get_unique_filename(directory, base_filename)
            os.rename(raw_file_path, processed_file)
            logger.info(f"Arquivo renomeado para: {processed_file}")
            
            self.thread_safe_update(70, "Processando arquivo...")
            logger.info("Processando planilha baixada...")
            df = pd.read_excel(processed_file)
            
            # Filter by responsible
            df = df[df['Responsável'] == responsible]
            
            # Remove specified columns
            columns_to_remove = [
                "Data de entrada",
                "Unidade",
                "Estado",
                "Cidade",
                "Cliente Referencial",
                "Segmento",
                "Produto",
                "Agendado",
                "Tempo Realizado",
                "Local do Treinamento"
            ]
            # Only drop columns that exist in the DataFrame to avoid KeyError
            columns_to_drop = [col for col in columns_to_remove if col in df.columns]
            df.drop(columns=columns_to_drop, inplace=True)
            logger.info(f"Colunas removidas: {columns_to_drop}")
            
            # Generate links
            link_base = "https://suporte.dominioatendimento.com/sgsc/faces/externo.html?externo={codigo}&externoPendente=true"
            df['link'] = df['Número'].apply(lambda x: link_base.format(codigo=x))
            
            # Overwrite the file with the processed data
            df.to_excel(processed_file, index=False)
            logger.info(f"Arquivo processado e sobrescrito: {processed_file}")
            
            # Apply formatting using openpyxl
            self.thread_safe_update(80, "Formatando planilha...")
            logger.info("Aplicando formatação ao arquivo...")
            workbook = load_workbook(processed_file)
            worksheet = workbook.active
            
            # Auto-adjust column widths
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter  # Get the column letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = max_length + 2  # Add padding
                worksheet.column_dimensions[column].width = adjusted_width
            
            # Apply borders to all cells
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            for row in worksheet.rows:
                for cell in row:
                    cell.border = thin_border
            
            # Save the formatted file
            workbook.save(processed_file)
            logger.info(f"Formatação aplicada e arquivo salvo: {processed_file}")
            
            self.thread_safe_update(90, "Atualizando configuração...")
            self.config = {
                'usuario': username,
                'senha': password,
                'diretorio_download': directory,
                'responsavel': responsible,
                'data_inicial': start_date,
                'data_final': end_date
            }
            with open('config.json', 'w') as f:
                json.dump(self.config, f)
            
            self.thread_safe_update(100, "Processo concluído!")
            messagebox.showinfo("Sucesso", "Processo concluído com sucesso!")
        
        except Exception as e:
            logger.error(f"Erro durante a execução: {str(e)}")
            self.thread_safe_update(0, "Erro")
            messagebox.showerror("Erro", f"Erro durante a execução: {str(e)}")
        
        finally:
            if 'driver' in locals():
                driver.quit()
                logger.info("ChromeDriver finalizado.")
            self.after(0, lambda: self.btn_execute.configure(state="normal"))

if __name__ == "__main__":
    try:
        app = App()
        app.mainloop()
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao iniciar a aplicação: {str(e)}")