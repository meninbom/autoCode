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
from cryptography.fernet import Fernet
from datetime import datetime as dt

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
ctk.set_default_color_theme("green")

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
        self.geometry("500x600")
        self.minsize(400, 800)
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self.config = self.load_config()
        self.encryption_key = self.load_or_generate_key()
        self.cipher = Fernet(self.encryption_key)
        self.responsible_selections = {}  # To track checkbox states
        
        self.create_widgets()
        self.populate_fields()
    
    def load_or_generate_key(self):
        key_file = 'key.key'
        if os.path.exists(key_file):
            with open(key_file, 'rb') as f:
                return f.read()
        else:
            key = Fernet.generate_key()
            with open(key_file, 'wb') as f:
                f.write(key)
            return key
    
    def encrypt_password(self, password):
        if not password:
            return ""
        return self.cipher.encrypt(password.encode()).decode()
    
    def decrypt_password(self, encrypted_password):
        if not encrypted_password:
            return ""
        try:
            return self.cipher.decrypt(encrypted_password.encode()).decode()
        except:
            return ""
    
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
        
        self.lbl_filename_prefix = ctk.CTkLabel(self.main_frame, text="Prefixo do Nome do Arquivo:")
        self.lbl_filename_prefix.pack(fill="x", padx=10, pady=5)
        self.entry_filename_prefix = ctk.CTkEntry(self.main_frame, placeholder_text="Ex: relatorio")
        self.entry_filename_prefix.pack(fill="x", padx=10, pady=5)
        
        self.lbl_responsible = ctk.CTkLabel(self.main_frame, text="Responsáveis (selecione um ou mais):")
        self.lbl_responsible.pack(fill="x", padx=10, pady=5)
        self.responsible_frame = ctk.CTkFrame(self.main_frame)
        self.responsible_frame.pack(fill="x", padx=10, pady=5)
        for responsible in self.responsible_list:
            var = ctk.BooleanVar(value=False)
            chk = ctk.CTkCheckBox(self.responsible_frame, text=responsible, variable=var)
            chk.pack(anchor="w", padx=5, pady=2)
            self.responsible_selections[responsible] = var
        
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
                "responsaveis": [],
                "data_inicial": "",
                "data_final": "",
                "filename_prefix": ""
            }
            with open('config.json', 'w') as f:
                json.dump(default_config, f)
            return default_config
    
    def populate_fields(self):
        if 'usuario' in self.config:
            self.entry_username.insert(0, self.config['usuario'])
        if 'senha' in self.config:
            decrypted_password = self.decrypt_password(self.config['senha'])
            self.entry_password.insert(0, decrypted_password)
        if 'diretorio_download' in self.config:
            self.entry_directory.insert(0, self.config['diretorio_download'])
        if 'responsaveis' in self.config:
            for responsible in self.config['responsaveis']:
                if responsible in self.responsible_selections:
                    self.responsible_selections[responsible].set(True)
        if 'data_inicial' in self.config:
            self.entry_start_date.insert(0, self.config['data_inicial'])
        if 'data_final' in self.config:
            self.entry_end_date.insert(0, self.config['data_final'])
        if 'filename_prefix' in self.config:
            self.entry_filename_prefix.insert(0, self.config['filename_prefix'])
    
    def save_config(self):
        selected_responsibles = [r for r, var in self.responsible_selections.items() if var.get()]
        config = {
            'usuario': self.entry_username.get(),
            'senha': self.encrypt_password(self.entry_password.get()),
            'diretorio_download': self.entry_directory.get(),
            'responsaveis': selected_responsibles,
            'data_inicial': self.entry_start_date.get(),
            'data_final': self.entry_end_date.get(),
            'filename_prefix': self.entry_filename_prefix.get()
        }
        with open('config.json', 'w') as f:
            json.dump(config, f)
        messagebox.showinfo("Configuração", "Configuração salva com sucesso!")
    
    def validate_date(self, date_str):
        pattern = r"^\d{2}/\d{2}/\d{2}$"
        if not re.match(pattern, date_str):
            return False
        try:
            dt.strptime(date_str, "%d/%m/%y")
            return True
        except ValueError:
            return False
    
    def validate_date_range(self, start_date, end_date):
        try:
            start = dt.strptime(start_date, "%d/%m/%y")
            end = dt.strptime(end_date, "%d/%m/%y")
            return end >= start
        except ValueError:
            return False
    
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
    
    def try_find_element(self, driver, xpaths, description):
        for xpath in xpaths:
            try:
                element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, xpath))
                )
                logger.info(f"Elemento encontrado: {description} com XPath {xpath}")
                return element
            except:
                logger.warning(f"Falha ao encontrar {description} com XPath {xpath}")
        raise Exception(f"Não foi possível encontrar {description} com os XPaths fornecidos")
    
    def _execute(self):
        username = self.entry_username.get()
        password = self.entry_password.get()
        directory = self.entry_directory.get()
        selected_responsibles = [r for r, var in self.responsible_selections.items() if var.get()]
        start_date = self.entry_start_date.get()
        end_date = self.entry_end_date.get()
        filename_prefix = self.entry_filename_prefix.get() or "relatorio"
        
        if not all([username, password, directory, selected_responsibles, start_date, end_date]):
            self.thread_safe_update(0, "Erro")
            messagebox.showerror("Erro", "Todos os campos são obrigatórios, incluindo pelo menos um responsável.")
            self.after(0, lambda: self.btn_execute.configure(state="normal"))
            return
        
        if not (self.validate_date(start_date) and self.validate_date(end_date)):
            self.thread_safe_update(0, "Erro")
            messagebox.showerror("Erro", "Datas devem estar no formato dd/mm/aa.")
            self.after(0, lambda: self.btn_execute.configure(state="normal"))
            return
        
        if not self.validate_date_range(start_date, end_date):
            self.thread_safe_update(0, "Erro")
            messagebox.showerror("Erro", "A data final não pode ser anterior à data inicial.")
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
            login_xpaths = {
                'user_input': [
                    '/html/body/div/form/input[2]',
                    '//input[@name="username"]'
                ],
                'pwd_input': [
                    '/html/body/div/form/input[3]',
                    '//input[@name="password"]'
                ],
                'login_btn': [
                    '/html/body/div/form/input[4]',
                    '//input[@type="submit"]'
                ]
            }
            
            user_input = self.try_find_element(driver, login_xpaths['user_input'], "campo de usuário")
            pwd_input = self.try_find_element(driver, login_xpaths['pwd_input'], "campo de senha")
            login_btn = self.try_find_element(driver, login_xpaths['login_btn'], "botão de login")
            
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
            date_xpaths = {
                'data_inicio': [
                    '//*[@id="programacoesForm:agendadoDe"]',
                    '//input[contains(@id, "agendadoDe")]'
                ],
                'data_fim': [
                    '//*[@id="programacoesForm:agendadoAte"]',
                    '//input[contains(@id, "agendadoAte")]'
                ]
            }
            
            logger.info("Preenchendo data inicial...")
            data_inicio_input = self.try_find_element(driver, date_xpaths['data_inicio'], "campo de data inicial")
            data_inicio_input.clear()
            data_inicio_input.click()
            data_inicio_input.send_keys(Keys.CONTROL + "a")
            data_inicio_input.send_keys(Keys.DELETE)
            data_inicio_input.send_keys(start_date)
            data_inicio_input.send_keys(Keys.TAB)
            
            logger.info("Preenchendo data final...")
            data_fim_input = self.try_find_element(driver, date_xpaths['data_fim'], "campo de data final")
            data_fim_input.clear()
            data_fim_input.click()
            data_fim_input.send_keys(Keys.CONTROL + "a")
            data_fim_input.send_keys(Keys.DELETE)
            data_fim_input.send_keys(end_date)
            data_fim_input.send_keys(Keys.TAB)
            
            self.thread_safe_update(40, "Salvando configurações de datas...")
            logger.info("Clicando no botão de salvar...")
            salvar_btn = self.try_find_element(driver, [
                '//*[@id="programacoesForm:atualizarBtn"]',
                '//button[contains(@id, "atualizarBtn")]'
            ], "botão de salvar")
            salvar_btn.click()
            time.sleep(5)
            
            self.thread_safe_update(45, "Gerando relatório...")
            logger.info("Clicando no botão de gerar relatório...")
            gerarelatorio_btn = self.try_find_element(driver, [
                '//*[@id="programacoesForm:gerarRelatorio"]',
                '//button[contains(@id, "gerarRelatorio")]'
            ], "botão de gerar relatório")
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
            
            filename = download_url.split('/')[-1] if '/' in download_url else "relatorio.xlsx"
            raw_file_path = os.path.join(directory, filename)
            
            with open(raw_file_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            logger.info(f"Arquivo baixado: {raw_file_path}")
            
            current_date = datetime.now().strftime("%Y-%m-%d")
            base_filename = f"{filename_prefix}_{current_date}.xlsx"
            processed_file = self.get_unique_filename(directory, base_filename)
            os.rename(raw_file_path, processed_file)
            logger.info(f"Arquivo renomeado para: {processed_file}")
            
            self.thread_safe_update(70, "Processando arquivo...")
            logger.info("Processando planilha baixada...")
            df = pd.read_excel(processed_file)
            
            if selected_responsibles:
                df = df[df['Responsável'].isin(selected_responsibles)]
            
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
            columns_to_drop = [col for col in columns_to_remove if col in df.columns]
            df.drop(columns=columns_to_drop, inplace=True)
            logger.info(f"Colunas removidas: {columns_to_drop}")
            
            link_base = "https://suporte.dominioatendimento.com/sgsc/faces/externo.html?externo={codigo}&externoPendente=true"
            df['link'] = df['Número'].apply(lambda x: link_base.format(codigo=x))
            
            df.to_excel(processed_file, index=False)
            logger.info(f"Arquivo processado e sobrescrito: {processed_file}")
            
            self.thread_safe_update(80, "Formatando planilha...")
            logger.info("Aplicando formatação ao arquivo...")
            workbook = load_workbook(processed_file)
            worksheet = workbook.active
            
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = max_length + 2
                worksheet.column_dimensions[column].width = adjusted_width
            
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            for row in worksheet.rows:
                for cell in row:
                    cell.border = thin_border
            
            workbook.save(processed_file)
            logger.info(f"Formatação aplicada e arquivo salvo: {processed_file}")
            
            self.thread_safe_update(90, "Atualizando configuração...")
            self.config = {
                'usuario': username,
                'senha': self.encrypt_password(password),
                'diretorio_download': directory,
                'responsaveis': selected_responsibles,
                'data_inicial': start_date,
                'data_final': end_date,
                'filename_prefix': filename_prefix
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