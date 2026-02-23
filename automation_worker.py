import os
import time
import datetime
import shutil
import zipfile
import threading
import schedule
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from selenium import webdriver
import sys
from PIL import Image, ImageDraw, ImageFont
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
# Matplotlib will be imported locally to avoid build/startup issues
# import matplotlib
# matplotlib.use('Agg') 
# import matplotlib.pyplot as plt
import io

class AutomationWorker(threading.Thread):
    def __init__(self, config, logger_callback=None, status_callback=None, clear_log_callback=None, update_stats_callback=None):
        super().__init__()
        self.config = config
        self.log_callback = logger_callback
        self.status_callback = status_callback
        self.clear_log_callback = clear_log_callback
        self.update_stats_callback = update_stats_callback
        self.success_count = 0
        self.fail_count = 0
        self.stop_flag = False
        self.daemon = True # Mata thread se app fechar

    def log(self, message):
        if self.log_callback:
            self.log_callback(message)
        else:
            print(message)

    def status(self, message):
        if self.status_callback:
            self.status_callback(message)

    def setup_driver(self):
        chrome_options = Options()
        chrome_options.add_argument("--ignore-certificate-errors")
        chrome_options.add_argument("--start-maximized")
        
        if self.config.get("headless", False):
            chrome_options.add_argument("--headless")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-software-rasterizer")
            chrome_options.add_argument("window-size=1920,1080")
            chrome_options.add_argument("--log-level=3")

        # Configurar download automático
        raw_download_dir = self.config.get("download_path", os.getcwd())
        download_dir = os.path.abspath(raw_download_dir) # Garante caminho absoluto
        
        if not os.path.exists(download_dir):
            try:
                os.makedirs(download_dir)
                self.log(f"Pasta criada: {download_dir}")
            except Exception as e:
                self.log(f"Erro ao criar pasta: {e}")

        prefs = {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "profile.default_content_settings.popups": 0
        }
        chrome_options.add_experimental_option("prefs", prefs)

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        return driver, download_dir # Retorna também o dir configurado

    def safe_click(self, driver, element):
        """Tenta clique normal, se falhar usa JS."""
        try:
            # Tenta scroll para o centro primeiro
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
            time.sleep(0.5)
            element.click()
        except Exception as e:
            self.log(f"Clique normal falhou ({str(e)[:50]}...), tentando JS...")
            driver.execute_script("arguments[0].click();", element)

    def _login(self, driver, wait):
        self.status("Realizando login...")
        try:
            if driver.find_elements(By.CSS_SELECTOR, "input[type='password']"):
                self.log("Inserindo credenciais...")
                user_input = driver.find_element(By.CSS_SELECTOR, "input[type='text'], input[name='username']")
                user_input.clear()
                user_input.send_keys(self.config["incontrol_user"])
                
                pass_input = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
                pass_input.clear()
                pass_input.send_keys(self.config["incontrol_password"])
                
                login_btn = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
                self.safe_click(driver, login_btn)
                time.sleep(8)
        except Exception as e:
            self.log(f"Login info/pulo: {e}")

    def _apply_date_filters(self, driver, wait):
        self.log("Aplicando filtro de data de HOJE...")
        try:
            # 0. Abrir Painel de Filtros se necessário
            try:
                btn_abrir = wait.until(EC.element_to_be_clickable((By.ID, "bnt_filtro_v3_filtros")))
                self.safe_click(driver, btn_abrir)
                time.sleep(1)
            except:
                pass # Pode já estar aberto

            now = datetime.datetime.now()
            start_str = now.strftime("%d/%m/%Y 00:00")
            end_str = now.strftime("%d/%m/%Y %H:%M")

            # Helper para preencher data
            def set_date(placeholder, value):
                inp = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f"input[placeholder='{placeholder}']")))
                self.safe_click(driver, inp)
                driver.execute_script("arguments[0].value = '';", inp)
                inp.send_keys(Keys.CONTROL + "a")
                inp.send_keys(Keys.BACKSPACE)
                inp.send_keys(value)
                inp.send_keys(Keys.TAB)

            set_date("Data inicial dos eventos", start_str)
            set_date("Data final dos eventos", end_str)
            
            # Botão Filtrar
            btn_filtrar = driver.find_element(By.ID, "bnt_filtro_v3_filtrar")
            self.safe_click(driver, btn_filtrar)
            self.log("Filtro de data submetido.")
            time.sleep(5)

        except Exception as e:
            self.log(f"Erro ao aplicar datas: {e}")

    def _trigger_export(self, driver, wait):
        self.log("Acionando menu de exportação...")
        # Tenta clicar no botão split ou na seta
        try:
            # Estrategia 1: Seta do SplitButton
            dropdown_arrow = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'ui-splitbutton-menubutton')]")))
            self.safe_click(driver, dropdown_arrow)
            time.sleep(1)
            
            # Estrategia 2: Item CSV
            csv_opt = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'CSV')]")))
            self.safe_click(driver, csv_opt)
            self.log("Opção CSV selecionada.")
            return True
        except Exception as e:
            self.log(f"Erro ao clicar exportar: {e}")
            return False

    def login_and_export_report(self):
        self.log("Iniciando extração...")
        self.status("Extraindo dados...")
        
        driver = None
        current_download_dir = self.config.get("download_path", os.getcwd())
        if not os.path.exists(current_download_dir):
            os.makedirs(current_download_dir)

        try:
            driver, _ = self.setup_driver() # setup_driver já lida com diretório se precisar
            driver.get(self.config["incontrol_url"])
            wait = WebDriverWait(driver, 30)

            self._login(driver, wait)

            if "eventos-usuario" not in driver.current_url:
                driver.get(self.config["incontrol_url"])
                time.sleep(5)

            # Filtra por data para evitar download gigante
            self._apply_date_filters(driver, wait)

            # Marca tempo
            start_export_time = time.time()
            
            if self._trigger_export(driver, wait):
                 # Monitorar
                 self.log("Aguardando download...")
                 return self._wait_for_download(current_download_dir, start_export_time)
            
            return None

        except Exception as e:
            self.log(f"Erro no fluxo Selenium: {e}")
            return None
        finally:
            if driver: driver.quit()

    def _wait_for_download(self, download_dir, start_time, timeout=90):
        loop_start = time.time()
        while time.time() - loop_start < timeout:
            if self.stop_flag: return None
            try:
                # Lista arquivos modificados APÓS o clique
                files = os.listdir(download_dir)
                candidates = []
                for f in files:
                    full_path = os.path.join(download_dir, f)
                    if f.endswith('.csv') and os.path.getmtime(full_path) > (start_time - 10):
                        candidates.append(full_path)
                
                # Se achou e não tem .crdownload
                if candidates and not any(f.endswith('.crdownload') for f in files):
                    latest = max(candidates, key=os.path.getmtime)
                    self.log(f"Download concluído: {os.path.basename(latest)}")
                    return self.organize_files(latest)
            except:
                pass
            time.sleep(1)
        self.log("Timeout esperando download.")
        return None

    def organize_files(self, downloaded_file_path):
        """Padroniza nomes e arquiva antigos."""
        dir_path = os.path.dirname(downloaded_file_path)
        timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        new_name = f"report-{timestamp}.csv"
        new_path = os.path.join(dir_path, new_name)
        
        try:
            time.sleep(1) 
            os.rename(downloaded_file_path, new_path)
        except:
            new_path = downloaded_file_path

        # Cria/Atualiza report-atual.csv
        atual_path = os.path.join(dir_path, "report-atual.csv")
        try:
            shutil.copy2(new_path, atual_path)
        except:
            pass
        
        # Arquiva TUDO menos o atual e o timestamp recém criado
        self.archive_old_reports(dir_path, current_timestamp_file=new_name)

        return atual_path

    def archive_old_reports(self, dir_path, current_timestamp_file):
        zip_path = os.path.join(dir_path, "Backups.zip")
        files_to_zip = []
        for f in os.listdir(dir_path):
            if f.endswith(".csv"):
                if f != "report-atual.csv" and f != current_timestamp_file:
                    files_to_zip.append(f)

        if files_to_zip:
            self.log(f"Arquivando {len(files_to_zip)} arquivos antigos...")
            try:
                with zipfile.ZipFile(zip_path, 'a', zipfile.ZIP_DEFLATED) as zipf:
                    for file in files_to_zip:
                        zipf.write(os.path.join(dir_path, file), file)
                for file in files_to_zip:
                    os.remove(os.path.join(dir_path, file))
            except Exception as e:
                self.log(f"Erro zip: {e}")

    def generate_graphs(self, df, df_inside):
        """Gera gráficos usando Matplotlib (import lazy)."""
        try:
            import matplotlib
            matplotlib.use('Agg')
            import matplotlib.pyplot as plt
        except ImportError:
            self.log("Matplotlib não instalado/encontrado. Gráficos ignorados.")
            return {}

        import io
        
        images = {}
        
        try:
            # Config Style
            plt.style.use('ggplot')
            
            # 1. Gráfico de Barras REMOVIDO conforme solicitação
            # Apenas Gráfico de Linha (Fluxo) mantido
            
            # 2. Gráfico de Fluxo por Hora (Linha)
            if not df.empty:
                plt.figure(figsize=(8, 4))
                df['Hora'] = df['Data Evento'].dt.hour
                hourly = df.groupby('Hora').size()
                # Config cor da linha para Verde Moderno
                hourly.plot(kind='line', marker='o', color='#198754', linestyle='-', linewidth=2)
                plt.title('Fluxo de Acessos por Hora', fontsize=12, color='#333')
                plt.xlabel('Hora do Dia')
                plt.ylabel('Movimentos')
                plt.grid(True, linestyle='--', alpha=0.6)
                plt.tight_layout()
                
                buf_line = io.BytesIO()
                plt.savefig(buf_line, format='png')
                buf_line.seek(0)
                images['line_chart'] = buf_line.read()
                plt.close()
                
        except Exception as e:
            self.log(f"Erro ao gerar gráficos: {e}")
            
        return images

    def process_data(self, csv_path):
        self.status("Analisando Dados Avançados...")
        try:
            # 1. Carregar Dados
            try:
                # Tenta UTF-8 primeiro (padrão moderno que corrige caracteres estranhos)
                df = pd.read_csv(csv_path, sep=';', encoding='utf-8', on_bad_lines='skip')
            except UnicodeDecodeError:
                # Se falhar, tenta Latin-1 (padrão legado Excel BR)
                self.log("Falha na leitura UTF-8. Tentando Latin-1...")
                df = pd.read_csv(csv_path, sep=';', encoding='latin1', on_bad_lines='skip')
            
            df.columns = df.columns.str.strip()
            
            # Normalização de Colunas (Evitar Duplicatas)
            df_cols_lower = [c.lower() for c in df.columns]
            
            col_mapping = {}
            # Prioridade de busca
            targets = {
                'Nome': ['nome do usuário', 'nome', 'usuario', 'colaborador'],
                'Departamento': ['departamento', 'setor', 'area'],
                'Status': ['status', 'situacao'],
                'Direcao': ['acesso', 'tipo de evento', 'direcao', 'tipo'], # 'Acesso' column has 'Entrada'/'Saída'
                'Local': ['ponto de acesso', 'local', 'porta'],
                'Data': ['data evento', 'data do evento', 'data', 'horario'],
                'CPF': ['cpf', 'documento'],
                'Matricula': ['matricula', 'cracha']
            }
            
            # Encontra a melhor correspondência para cada target
            used_indices = set()
            
            for target_key, candidate_list in targets.items():
                found_col_name = None
                # Tenta match exato primeiro na lista de prioridade
                for cand in candidate_list:
                    for i, col_real in enumerate(df.columns):
                        if i in used_indices: continue
                        if cand in col_real.lower():
                            found_col_name = col_real
                            used_indices.add(i)
                            break
                    if found_col_name: break
                
                if found_col_name:
                    col_mapping[found_col_name] = target_key
            
            # Renomeia apenas as colunas encontradas e descarta o resto para evitar confusão
            df = df.rename(columns=col_mapping)

            # --- FILTRAGEM DE SEGURANÇA (PANDAS) ---
            # Garante que só fique 'Acesso Liberado' se o driver baixou errado.
            # --- FILTRAGEM DE SEGURANÇA (PANDAS) ---
            # Garante que só fique 'Acesso Liberado' se o driver baixou errado.
            if 'Status' in df.columns:
                df = df.dropna(subset=['Status'])
                initial_count = len(df)
                df = df[df['Status'].astype(str).str.contains("Liberado", case=False, na=False)]
                
                # Filtro de Usuários Desconhecidos/Inválidos
                if 'Nome' in df.columns:
                     df = df[~df['Nome'].astype(str).str.contains("Desconhecido|Visitante|N/A", case=False, na=False)]
                     
                final_count = len(df)
                self.log(f"Filtro Excel aplicado: {initial_count} -> {final_count} ('Acesso Liberado' e 'Usuários Válidos').")
            else:
                self.log("AVISO: Coluna 'Status' não encontrada. Filtro Excel ignorado.")
            
            # Garante que temos as colunas essenciais E opcionais
            # Se não achou Matricula ou CPF, cria com N/A
            for req in ['Nome', 'Departamento', 'Status', 'Data', 'CPF', 'Matricula']:
                if req not in df.columns:
                     df[req] = 'N/A'
            
            # Filtra apenas colunas normalizadas para limpar o DF
            # Filtra apenas colunas normalizadas para limpar o DF
            final_cols = [c for c in df.columns if c in ['Nome', 'Departamento', 'Status', 'Direcao', 'Local', 'Data', 'CPF', 'Matricula']]
            df = df[final_cols]
            
            # Remove duplicatas de colunas se houver (safety net)
            df = df.loc[:, ~df.columns.duplicated()]

            # Conversão de Data
            df['Data Evento'] = pd.to_datetime(df['Data'], errors='coerce')
            df = df.dropna(subset=['Data Evento']) # Remove datas inválidas
            
            # Filtro Hoje
            hoje = datetime.datetime.now().date()
            df = df[df['Data Evento'].dt.date == hoje]
            
            # --- LÓGICA DE PRESENÇA (Quem está na obra?) ---
            # 1. Deduplicação Temporal: Remove eventos duplicados da mesma pessoa em < 60 segundos
            df = df.sort_values(by=['Nome', 'Data Evento'])
            df['TimeDiff'] = df.groupby('Nome')['Data Evento'].diff()
            # Mantém se for primeira ocorrência ou diferença > 60s
            df = df[(df['TimeDiff'].isnull()) | (df['TimeDiff'] > pd.Timedelta(seconds=60))]
            
            # 2. Agrupa por Identidade Única (Nome + CPF + Matricula)
            # Ordena com segurança para garantir consistência
            df_sorted = df.sort_values(by=['Nome', 'CPF', 'Matricula', 'Data Evento'])
            
            # Pega o último status de cada pessoa
            df_last_status = df_sorted.groupby(['Nome', 'CPF', 'Matricula'], dropna=False).last().reset_index()
            

            
            # Quem está dentro? (Último evento == Entrada) e (Status != Saída)
            # --- LÓGICA DE PRESENÇA (QUEM ESTÁ DENTRO?) ---
            
            # 1. Normalizar o texto da coluna Direcao para evitar erros de acento (SaÃda -> Saida)
            # Remove acentos e joga tudo para minúsculo
            if 'Direcao' in df_last_status.columns:
                # Cria uma coluna temporária normalizada
                df_last_status['Direcao_Norm'] = df_last_status['Direcao'].astype(str).str.lower().str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')
                
                # Agora 'entrada' vira 'entrada' e 'SaÃda' (ou 'Saída') vira 'saida'
                
                # Quem está dentro? 
                # Último evento contém 'entra' (pega 'entrada', 'entrando', etc)
                # E NÃO contém 'sai' (pega 'saida', 'saída', 'saiu')
                df_inside = df_last_status[
                    (df_last_status['Direcao_Norm'].str.contains('entra', na=False)) & 
                    (~df_last_status['Direcao_Norm'].str.contains('sai', na=False))
                ]
                
                # Log para debug (opcional, ajuda a ver se funcionou)
                self.log(f"Debug Presença: Encontrados {len(df_inside)} registros de entrada.")
                
            else:
                self.log("ERRO CRÍTICO: Coluna 'acesso/Direcao' não encontrada mesmo após mapeamento.")
                df_inside = pd.DataFrame()
            
            # Estatísticas
            total_presentes = len(df_inside)
            total_movimentos = len(df)
            
            # Tabela de Presença por Dept
            dept_stats = df_inside.groupby('Departamento').size().reset_index(name='Qtd')
            dept_stats = dept_stats.sort_values('Qtd', ascending=False)
            
            # Calcular Porcentagem
            total_dept = dept_stats['Qtd'].sum()
            dept_stats['Pct'] = (dept_stats['Qtd'] / total_dept * 100).fillna(0)
            
            # html Table rows (3 colunas: Dept, Qtd, Pct)
            rows_html = ""
            for _, row in dept_stats.iterrows():
                pct_val = row['Pct']
                rows_html += f"""
                <tr style="border-bottom: 1px solid #f0f0f0;">
                    <td style="padding: 15px 20px; font-size: 14px; color: #333;">{row['Departamento']}</td>
                    <td style="padding: 15px 20px; text-align: right; font-weight: bold; color: #333; font-size: 14px;">{row['Qtd']}</td>
                    <td style="padding: 15px 20px; text-align: right;">
                        <span style="background-color: #e8f5e9; color: #198754; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 600;">{pct_val:.1f}%</span>
                    </td>
                </tr>
                """
            
            # Gerar Gráficos com Tratamento de Erro
            graph_imgs = {}
            try:
                graph_imgs = self.generate_graphs(df, df_inside)
            except Exception as e:
                self.log(f"ERRO CRÍTICO AO GERAR GRÁFICOS: {e}. Enviando sem gráficos.")
            
            # Conteúdo condicional de gráficos
            graphs_html = ""
            if 'line_chart' in graph_imgs:
                graphs_html = """
                     <h3 style="color: #198754; font-size: 16px; margin-top: 30px; border-left: 4px solid #ffc107; padding-left: 10px;">Fluxo Horário</h3>
                    <div style="text-align: center; margin-top: 15px;">
                        <img src="cid:line_chart" style="max-width: 100%; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.05);">
                    </div>
                """
            
            # Tradução de Datas
            meses_pt = {
                1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril', 5: 'Maio', 6: 'Junho',
                7: 'Julho', 8: 'Agosto', 9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
            }
            hoje = datetime.datetime.now()
            data_pt = f"{hoje.day} de {meses_pt[hoje.month]} de {hoje.year}"
            
            # Montar HTML (Tema Premium: Espaçado, Legível e Sofisticado)
            # Fundo Off-White (#f4f6f8) para menos cansaço visual
            # Fontes maiores e com line-height generoso
            html_body = f"""
            <html>
            <body style="font-family: 'Segoe UI', 'Roboto', Helvetica, Arial, sans-serif; background-color: #f4f6f8; margin: 0; padding: 20px;">
                
                <div style="max-width: 650px; margin: 0 auto; background-color: #ffffff; border-radius: 16px; box-shadow: 0 8px 30px rgba(0,0,0,0.08); overflow: hidden;">
                    
                    <!-- Header Sofisticado (Verde Profundo) -->
                    <div style="background: linear-gradient(135deg, #06402b 0%, #1e5c3f 100%); color: white; padding: 40px 30px; text-align: center;">
                        <h2 style="margin: 0; font-size: 26px; font-weight: 500; letter-spacing: 0.5px;">Relatório Diário de Obra</h2>
                        <div style="width: 50px; height: 3px; background-color: #ffc107; margin: 15px auto; border-radius: 2px;"></div>
                        <p style="margin: 10px 0 0 0; font-size: 16px; opacity: 0.9;">Monitoramento de Acesso</p>
                        <div style="margin-top: 20px; font-size: 14px; color: #e8f5e9; font-weight: 500;">
                            📅 {data_pt}
                        </div>
                    </div>
                    
                    <div style="padding: 40px 30px;">
                        
                        <!-- Cards Totais (Design Limpo) -->
                        <div style="display: flex; gap: 20px; margin-bottom: 40px;">
                            <!-- Card Total -->
                            <div style="flex: 1; padding: 25px 20px; border-radius: 12px; text-align: center; border: 1px solid #e0e0e0; background: #ffffff; box-shadow: 0 2px 8px rgba(0,0,0,0.03);">
                                <div style="color: #555; font-size: 13px; text-transform: uppercase; letter-spacing: 1px; font-weight: 600; margin-bottom: 8px;">Total de Movimentos</div>
                                <div style="font-size: 36px; font-weight: 700; color: #2c3e50;">{total_movimentos}</div>
                            </div>
                            
                            <!-- Card Presentes (Destaque) -->
                            <div style="flex: 1; padding: 25px 20px; border-radius: 12px; text-align: center; border: 1px solid #c3e6cb; background: #f1f8f3; box-shadow: 0 4px 12px rgba(25, 135, 84, 0.1);">
                                <div style="color: #155724; font-size: 13px; text-transform: uppercase; letter-spacing: 1px; font-weight: 600; margin-bottom: 8px;">Pessoas na Obra</div>
                                <div style="font-size: 36px; font-weight: 700; color: #198754;">{total_presentes}</div>
                            </div>
                        </div>

                        <!-- Tabela (Espaçada e Legível) -->
                        <div style="margin-bottom: 40px;">
                            <h3 style="color: #2c3e50; font-size: 18px; margin-bottom: 20px; padding-left: 15px; border-left: 4px solid #198754; font-weight: 600;">
                                Presença por Departamento
                            </h3>
                            
                            <table style="width: 100%; border-collapse: separate; border-spacing: 0;">
                                <thead>
                                    <tr style="background-color: #fafafa;">
                                        <th style="padding: 15px 20px; text-align: left; color: #7f8c8d; font-size: 12px; font-weight: 600; text-transform: uppercase; border-bottom: 2px solid #eee;">Departamento</th>
                                        <th style="padding: 15px 20px; text-align: right; color: #7f8c8d; font-size: 12px; font-weight: 600; text-transform: uppercase; border-bottom: 2px solid #eee;">Qtd.</th>
                                        <th style="padding: 15px 20px; text-align: right; color: #7f8c8d; font-size: 12px; font-weight: 600; text-transform: uppercase; border-bottom: 2px solid #eee;">%</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {rows_html}
                                </tbody>
                            </table>
                        </div>
                        
                        <!-- Gráfico -->
                        {graphs_html}
                        
                    </div>
                    
                    <!-- Footer -->
                    <div style="background-color: #2c3e50; padding: 25px; text-align: center;">
                        <p style="margin: 0; font-size: 12px; color: #bdc3c7; line-height: 1.5;">
                            Relatório gerado automaticamente pelo sistema <b>BotEliva</b>.<br>
                            Este é um email informativo, por favor não responda.
                        </p>
                    </div>
                </div>
            </body>
            </html>
            """

            
            # --- GERAR IMAGEM RESUMO (Social Media) ---
            summary_img_path = None
            try:
                summary_img_path = self.generate_summary_card(
                    total_movimentos, 
                    total_presentes, 
                    dept_stats, 
                    data_pt
                )
            except Exception as e:
                self.log(f"Erro ao gerar imagem resumo: {e}")

            return {'html': html_body, 'images': graph_imgs, 'summary_image': summary_img_path}

        except Exception as e:
            return f"Erro processamento dados: {e}"

    def generate_summary_card(self, total_mov, total_pres, df_dept, data_str):
        """Gera uma imagem (card) para compartilhamento."""
        try:
            # Configurações de Design
            width = 800
            # Altura dinâmica baseada nos departamentos + cabeçalho/rodapé
            row_height = 50
            header_height = 250
            stats_height = 150
            footer_height = 80
            padding = 40
            
            table_height = (len(df_dept) * row_height) + 60 # +60 para titulo da tabela
            total_height = header_height + stats_height + table_height + footer_height
            
            # Cores
            bg_color = "#ffffff"
            header_bg = "#198754" # Verde Intelbras/Juparana
            text_color = "#333333"
            accent_color = "#ffc107" # Amarelo
            light_gray = "#f8f9fa"
            
            # Criar Imagem
            img = Image.new('RGB', (width, total_height), bg_color)
            draw = ImageDraw.Draw(img)
            
            # --- Fonts (Tentativa de carregar fontes do sistema ou padrão) ---
            try:
                # Tenta Arial ou Segoe UI (Windows)
                font_title = ImageFont.truetype("arialbd.ttf", 40)
                font_subtitle = ImageFont.truetype("arial.ttf", 18)
                font_big_num = ImageFont.truetype("arialbd.ttf", 55)
                font_label = ImageFont.truetype("arialbd.ttf", 16)
                font_row = ImageFont.truetype("arial.ttf", 22)
                font_row_bold = ImageFont.truetype("arialbd.ttf", 22)
            except:
                # Fallback
                font_title = ImageFont.load_default()
                font_subtitle = ImageFont.load_default()
                font_big_num = ImageFont.load_default()
                font_label = ImageFont.load_default()
                font_row = ImageFont.load_default()
                font_row_bold = ImageFont.load_default()

            # --- 1. HEADER ---
            draw.rectangle([(0, 0), (width, header_height)], fill=header_bg)
            
            # Logo
            logo_path = "logojup.png"
            if getattr(sys, 'frozen', False):
                logo_path = os.path.join(sys._MEIPASS, "logojup.png")
            
            if os.path.exists(logo_path):
                try:
                    logo = Image.open(logo_path).convert("RGBA")
                    # Redimensionar mantendo aspect ratio (max height 100px)
                    logo.thumbnail((200, 100), Image.Resampling.LANCZOS)
                    # Colocar no canto superior direito (com padding)
                    logo_x = width - logo.width - padding
                    logo_y = padding
                    # Colar alpha channel corretamente (mas fundo é verde, precisa compor)
                    # Melhor colar mascara
                    img.paste(logo, (logo_x, logo_y), logo)
                except Exception as e:
                    self.log(f"Erro ao carregar logo: {e}")

            # Título
            draw.text((padding, padding), "Relatório Diário", font=font_title, fill="white")
            draw.text((padding, padding + 55), "de Acesso e Presença", font=ImageFont.truetype("arial.ttf", 28) if 'arial' in str(font_title.path) else font_title, fill="#e8f5e9")
            
            # Data
            draw.text((padding, header_height - 60), f"Data: {data_str}", font=ImageFont.truetype("arialbd.ttf", 20) if 'arial' in str(font_title.path) else font_title, fill="white")
            
            # Detalhe amarelo
            draw.rectangle([(padding, padding + 48), (padding + 100, padding + 51)], fill=accent_color)

            # --- 2. CARDS STATS ---
            current_y = header_height + 20
            card_width = (width - (3 * padding)) // 2
            card_height = 120
            
            # Card 1: Total Movimentos
            draw.rounded_rectangle([(padding, current_y), (padding + card_width, current_y + card_height)], radius=15, fill=light_gray, outline="#d1d1d1", width=1)
            draw.text((padding + 20, current_y + 20), "TOTAL MOVIMENTOS", font=font_label, fill="#555")
            draw.text((padding + 20, current_y + 50), str(total_mov), font=font_big_num, fill="#2c3e50")
            
            # Card 2: Presentes
            x_card2 = padding + card_width + padding
            # Fundo verde claro para destaque
            draw.rounded_rectangle([(x_card2, current_y), (x_card2 + card_width, current_y + card_height)], radius=15, fill="#e8f5e9", outline="#c3e6cb", width=1)
            draw.text((x_card2 + 20, current_y + 20), "PESSOAS NA OBRA", font=font_label, fill="#155724")
            draw.text((x_card2 + 20, current_y + 50), str(total_pres), font=font_big_num, fill="#198754")

            # --- 3. TABELA DEPARTAMENTOS ---
            current_y += card_height + 40
            
            draw.text((padding, current_y), "Distribuição por Departamento", font=ImageFont.truetype("arialbd.ttf", 24) if 'arial' in str(font_title.path) else font_label, fill="#198754")
            # Linha decorativa
            draw.rectangle([(padding, current_y + 35), (width - padding, current_y + 36)], fill="#eee")
            
            current_y += 50
            
            # Cabeçalho da tabela
            draw.text((padding + 10, current_y), "DEPARTAMENTO", font=font_label, fill="#999")
            draw.text((width - padding - 180, current_y), "QTD", font=font_label, fill="#999", anchor="ra")
            draw.text((width - padding - 50, current_y), "%", font=font_label, fill="#999", anchor="ra")
            
            current_y += 30
            
            # Linhas
            for _, row in df_dept.iterrows():
                dept = row['Departamento']
                # Truncar nome se muito longo
                if len(dept) > 35: dept = dept[:32] + "..."
                
                qtd = str(row['Qtd'])
                pct = f"{row['Pct']:.1f}%"
                
                # Fundo alternado (opcional)
                # draw.rectangle([(padding, current_y), (width-padding, current_y+row_height)], fill="#fafafa")
                
                draw.text((padding + 10, current_y + 10), dept, font=font_row, fill="#333")
                draw.text((width - padding - 180, current_y + 10), qtd, font=font_row_bold, fill="#333", anchor="ra")
                
                # Pill para porcentagem
                # draw.rounded_rectangle ... (complexo no PIL puro, vamos texto simples verde)
                draw.text((width - padding - 50, current_y + 10), pct, font=font_row_bold, fill="#198754", anchor="ra")
                
                # Linha separadora sutil
                draw.line([(padding, current_y + row_height), (width - padding, current_y + row_height)], fill="#f0f0f0", width=1)
                
                current_y += row_height

            # --- 4. FOOTER ---
            # current_y agora está no fim da tabela. Vamos mover para o fim da imagem fixo ou relativo?
            # Imagem foi criada com altura dinâmica, então current_y deve estar perto do fim menos footer_height
            
            footer_y = total_height - footer_height
            draw.rectangle([(0, footer_y), (width, total_height)], fill=light_gray)
            draw.text((width//2, footer_y + 30), "Gerado automaticamente por BotEliva", font=font_subtitle, fill="#999", anchor="mm")
            
            # Salvar
            # Usa path configurado ou atual
            target_dir = self.config.get("download_path", os.getcwd())
            if not os.path.exists(target_dir):
                target_dir = os.getcwd()
                
            output_path = os.path.join(target_dir, "IMGGerada.png")
            
            # Salva (sobrescreve se existir)
            img.save(output_path)
            self.log(f"Imagem resumo salva/atualizada em: {output_path}")
            return output_path
            
        except Exception as e:
            self.log(f"Erro ao desenhar summary card: {e}")
            return None

    def send_email(self, report_data):
        self.status("Enviando Email...")
        
        # Se houve erro no processamento, report_data será string
        if isinstance(report_data, str):
            return f"Erro Dados: {report_data}"
            
        sender_email = self.config.get("smtp_user")
        recipients_raw = self.config.get("email_recipients", "")
        if not recipients_raw: return "Erro: Nenhum destinatário." 
        recipients_list = [email.strip() for email in recipients_raw.split(",") if email.strip()]
        
        password = self.config.get("smtp_password")
        smtp_server = self.config.get("smtp_server")
        smtp_port = self.config.get("smtp_port")

        if not all([sender_email, recipients_list, password, smtp_server, smtp_port]):
            return "Email não configurado."

        try:
            server = smtplib.SMTP(smtp_server, int(smtp_port))
            server.starttls()
            server.login(sender_email, password)
            
            count_sent = 0
            for receiver_email in recipients_list:
                msg = MIMEMultipart('related') # Important for images
                msg['From'] = sender_email
                msg['To'] = receiver_email
                msg['Subject'] = f"📊 Relatório Diário de Obra - {datetime.datetime.now().strftime('%d/%m')}"
                
                msg_alternative = MIMEMultipart('alternative')
                msg.attach(msg_alternative)
                
                # HTML Body
                msg_text = MIMEText(report_data['html'], 'html')
                msg_alternative.attach(msg_text)
                
                # Attach Images
                for c_id, img_data in report_data['images'].items():
                    img = MIMEBase('image', 'png')
                    img.set_payload(img_data)
                    encoders.encode_base64(img)
                    img.add_header('Content-ID', f'<{c_id}>')
                    img.add_header('Content-Disposition', 'inline', filename=f'{c_id}.png')
                    img.add_header('Content-Disposition', 'inline', filename=f'{c_id}.png')
                    msg.attach(img)
                
                # Anexar Imagem Resumo (Social Media)
                summary_path = report_data.get('summary_image')
                if summary_path and os.path.exists(summary_path):
                    try:
                        with open(summary_path, 'rb') as f:
                            part = MIMEBase('image', 'png')
                            part.set_payload(f.read())
                            encoders.encode_base64(part)
                            part.add_header('Content-Disposition', 'attachment; filename="Resumo_Obra.png"')
                            msg.attach(part)
                            self.log("Imagem resumo anexada com sucesso.")
                    except Exception as e:
                        self.log(f"Erro ao anexar resumo: {e}")
                
                text_msg = msg.as_string()
                server.sendmail(sender_email, receiver_email, text_msg)
                count_sent += 1
            
            server.quit()
            return f"Relatório enviado para {count_sent} emails!"
        except Exception as e:
            import traceback
            traceback.print_exc()
            return f"Erro envio: {e}"

    def stop(self):
        self.stop_flag = True
        self.log("Parando serviço...")

    def run(self):
        self.log("Serviço de Monitoramento Iniciado (Agendamento Customizado).")
        schedule.clear()
        
        sched_type = self.config.get("schedule_type", "Diário").lower()
        times_str = self.config.get("schedule_times", "08:00")
        
        # Formatador de tempo ("3" -> "03:00", "8:30" -> "08:30")
        def pad_time(t_str):
            t_str = t_str.strip()
            if not t_str: return "08:00"
            if ":" in t_str:
                h, m = t_str.split(":", 1)
                try: return f"{int(h):02d}:{int(m):02d}"
                except: return "08:00"
            else:
                try: return f"{int(t_str):02d}:00"
                except: return "08:00"

        raw_times = [t for t in times_str.split(",") if t.strip()]
        if not raw_times: raw_times = ["08:00"]
        times_list = [pad_time(t) for t in raw_times]
        
        days_str = self.config.get("schedule_days", "")
        days_list = [d.strip().lower() for d in days_str.split(",") if d.strip()]
        
        for t in times_list:
            if "diario" in sched_type or "diário" in sched_type:
                schedule.every().day.at(t).do(self.run_task)
                self.log(f"Agendado: Diariamente às {t}")
            
            elif "semanal" in sched_type:
                day_map = {
                    "seg": schedule.every().monday, "ter": schedule.every().tuesday,
                    "qua": schedule.every().wednesday, "qui": schedule.every().thursday,
                    "sex": schedule.every().friday, "sab": schedule.every().saturday,
                    "sáb": schedule.every().saturday, "dom": schedule.every().sunday
                }
                if not days_list:
                    schedule.every().monday.at(t).do(self.run_task)
                    self.log(f"Agendado: Semanalmente (Seg) às {t}")
                else:
                    for d in days_list:
                        for k, v in day_map.items():
                            if k in d:
                                v.at(t).do(self.run_task)
                                self.log(f"Agendado: Semanalmente ({k.capitalize()}) às {t}")
                                break
                                
            elif "mensal" in sched_type:
                schedule.every().day.at(t).do(self.run_task_monthly, days_list)
                self.log(f"Agendado: Mensalmente nos dias '{days_str}' às {t}")

        def _update_next_run_ui():
            nxt = schedule.next_run()
            if nxt and self.update_stats_callback:
                # Enviar apenas Sucesso e Falhas mantidas, e Próximo ciclo
                self.update_stats_callback(self.success_count, self.fail_count, nxt.strftime('%d/%m %H:%M'))
                
        _update_next_run_ui()
        self.status("Aguardando próximo horário...")
        
        while not self.stop_flag:
            schedule.run_pending()
            _update_next_run_ui()
            time.sleep(1)
        
        self.status("Serviço Parado.")

    def run_task_monthly(self, days_list):
        hoje = str(datetime.datetime.now().day)
        if not days_list:
            if hoje == "1": self.run_task()
        else:
            if hoje in days_list: self.run_task()

    def run_task(self):
        if self.stop_flag: return
        
        if self.clear_log_callback:
            self.clear_log_callback()
            
        if self.update_stats_callback:
             self.update_stats_callback(self.success_count, self.fail_count, "Executando...")
             
        self.log(f"--- Iniciando Ciclo: {datetime.datetime.now().strftime('%H:%M')} ---")
        
        csv_path = self.login_and_export_report()
        if not csv_path or self.stop_flag:
            if not self.stop_flag:
                self.fail_count += 1
                self.log("Falha na extração. Tentando novamente no próximo ciclo.")
            self.status("Aguardando...")
            return

        report = self.process_data(csv_path)
        self.log("Relatório gerado.")
        
        email_status = self.send_email(report)
        if "Erro" in str(email_status):
            self.fail_count += 1
        else:
            self.success_count += 1
            
        self.log(email_status)
        self.log("Ciclo finalizado. Aguardando próximo horário.")
        self.status("Aguardando...")
