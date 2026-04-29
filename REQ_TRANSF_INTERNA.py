# -*- coding: utf-8 -*-
import pandas as pd
import win32com.client
import sys
import gspread
from datetime import datetime, timedelta
import subprocess
import time
import re
import os
import configparser
import pywintypes
import ssl
from dotenv import load_dotenv

# Ajuste SSL para requisições
ssl._create_default_https_context = ssl._create_unverified_context

# --- Configuração de Cores para o Terminal ---
class Colors:
    RESET = "\033[0m"
    VERDE = "\033[92m"
    AMARELO = "\033[93m"
    VERMELHO = "\033[91m"
    AZUL = "\033[94m"
    CIANO = "\033[96m"

class SAPBotCLI:
    # Mapeamento de Depósitos por Origem
    DEPOSITO_MAPPING = {
        'BR0G': 'AE01', 'BR0Q': 'AE01', 'BR0D': 'AE01', 'BR0H': 'AE01', 'BR0O': 'AE01',
        'BR0P': 'AE01', 'BR0E': 'AE01', 'BR0R': 'AE01', 'BR0S': 'AE01', 'BR0Y': 'AE01',
        'BR0Z': 'AE01', 'BR1A': 'AE01', 'BR1C': 'AE01', 'BR1D': 'AE01', 'BR1G': 'AE01',
        'BR1I': 'AE01', 'BR1J': 'AE01', 'BR1K': 'AE01', 'BR1L': 'AE01', 'BR1T': 'AE01',
        'BR2A': 'AE01', 'BR2B': 'AE01', 'BR2C': 'AE01', 'BR2D': 'AE01', 'BR2E': 'AE01',
        'BR2Q': 'AE01', 'BR2U': 'AE01', 'BR2V': 'AE01', 'BR3A': 'AE01', 'BR3E': 'AE01',
        'BR3F': 'AE01', 'BR3K': 'AE01', 'BR3N': 'AE01', 'BRDN': 'AE01', 'BR8A': 'AE13',
        'BR2I': 'AE01', 'BR0I': 'AE13', 'BR0U': 'AE01', 'BR0K': 'AE13', 'BR0X': 'AE13',
        'BR0J': 'AE01', 'BR1E': 'AE01', 'BR1F': 'AE01', 'BR0V': 'AE01', 'BR8E': 'AE13',
        'BR1B': 'AE01', 'BR0F': 'AE01', 'BR8I': 'AE01', 'BRIJ': 'AE01', 'BR8G': 'AE01'
    }

    def __init__(self):
        self.running = True
        self.session = None
        self.config = configparser.ConfigParser()
        
        # Define os caminhos base
        if getattr(sys, 'frozen', False):
            self.base_path = os.path.dirname(sys.executable)
        else:
            self.base_path = os.path.dirname(os.path.abspath(__file__))
            
        self.config_path = os.path.join(self.base_path, 'config.ini')
        self.log_file_path = os.path.join(self.base_path, 'app_log.txt')
        
        # Carrega variáveis de ambiente do arquivo .env
        env_path = os.path.join(self.base_path, '.env')
        load_dotenv(env_path)
        
        # Inicializa logs
        with open(self.log_file_path, 'a', encoding='utf-8') as log_file:
            log_file.write(f"\n{'='*50}\nSessão iniciada em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
            log_file.write(f"Caminho base: {self.base_path}\n{'='*50}\n")
            
        self.load_config()

    def load_config(self):
        if not os.path.exists(self.config_path):
            self.create_default_config()
            self.print_aviso(f"Arquivo 'config.ini' criado em: {self.config_path}")
            self.print_aviso("Por favor, verifique as configurações e rode o script novamente.")
            sys.exit(0)
            
        self.config.read(self.config_path, encoding='utf-8')
        
        # Validação simples
        if not self.config.get('GOOGLE', 'credenciais') or not self.config.get('SAP', 'caminho_logon'):
            self.print_erro("Configurações ausentes no config.ini. Preencha-o antes de continuar.")
            sys.exit(1)

    def create_default_config(self):
        self.config['SAP'] = {
            'caminho_logon': r'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe', 
            'sistema': 'ECC PRODUÇÃO'
        }
        self.config['GOOGLE'] = {
            'credenciais': 'credentials.json', 
            'planilha': 'MAPEAMENTO PLANNING', 
            'aba': 'REQ INTERNA'
        }
        with open(self.config_path, 'w', encoding='utf-8') as configfile:
            self.config.write(configfile)

    # --- Funções de Log ---
    def print_header(self, texto):
        log_text = f"\n{'='*60}\n {texto.center(58)}\n {'='*60}"
        print(f"{Colors.AZUL}{log_text}{Colors.RESET}")
        self._write_to_log_file(log_text)

    def print_sucesso(self, texto):
        log_text = f"[SUCESSO] {texto}"
        print(f"{Colors.VERDE}{log_text}{Colors.RESET}")
        self._write_to_log_file(log_text)

    def print_info(self, texto):
        log_text = f"[INFO]    {texto}"
        print(f"{Colors.CIANO}{log_text}{Colors.RESET}")
        self._write_to_log_file(log_text)

    def print_aviso(self, texto):
        log_text = f"[AVISO]   {texto}"
        print(f"{Colors.AMARELO}{log_text}{Colors.RESET}")
        self._write_to_log_file(log_text)

    def print_erro(self, texto):
        log_text = f"[ERRO]    {texto}"
        print(f"{Colors.VERMELHO}{log_text}{Colors.RESET}")
        self._write_to_log_file(log_text)

    def _write_to_log_file(self, text_to_log):
        try:
            timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            cleaned_text = re.sub(r'\033\[[0-9;]*m', '', text_to_log) # Remove códigos de cor ANSI
            with open(self.log_file_path, 'a', encoding='utf-8') as log_file:
                log_file.write(f"[{timestamp}] {cleaned_text.strip()}\n")
        except Exception: pass

    # --- Automação SAP ---
    def run(self):
        try:
            self.print_header("Iniciando Robô de Requisição de Compra no SAP")
            
            # Conexão SAP
            if not self.is_session_valid():
                self.print_aviso("Sessão SAP inválida ou inexistente. Tentando conectar...")
                self.session = self.sap_login_handler()

            if not self.session:
                self.print_erro("Falha na conexão com o SAP. Verifique se o SAP está acessível e as credenciais no .env estão corretas.")
                return
            
            self.print_sucesso("Sessão SAP estabelecida com sucesso!")
            
            # Processamento Planilha
            try:
                self.print_header("CONECTANDO À PLANILHA")
                credenciais_path = os.path.join(self.base_path, self.config.get('GOOGLE', 'credenciais'))
                gc = gspread.service_account(filename=credenciais_path)
                spreadsheet = gc.open(self.config.get('GOOGLE', 'planilha'))
                worksheet = spreadsheet.worksheet(self.config.get('GOOGLE', 'aba'))
                self.print_sucesso("Conexão com a planilha estabelecida.")
                
                headers = worksheet.row_values(1)
                status_col_index = headers.index("Status") + 1
                req_col_index = headers.index("REQUISIÇÃO") + 1
                
                df = pd.DataFrame(worksheet.get_all_records())
                df['linha_planilha'] = df.index + 2
                
                # Considera apenas linhas sem status
                df_para_processar = df[df['Status'] == ''].copy()

                if df_para_processar.empty:
                    self.print_aviso("Nenhuma linha nova para processar.")
                else:
                    self.processar_lotes(df_para_processar, worksheet, status_col_index, req_col_index)

            except Exception as e:
                self.print_erro(f"Erro crítico no ciclo principal: {e}")
                self.session = None 
            
        except KeyboardInterrupt:
            self.print_aviso("\nExecução interrompida pelo usuário (Ctrl+C).")
        except Exception as e:
            self.print_erro(f"Erro fatal na automação: {str(e)}")
        finally:
            self.print_header("FIM DO CICLO")

    def aguardar_sap(self, timeout=30):
        if not self.session: return False
        start_time = time.time()
        while self.running:
            try:
                if not self.session.busy: return True
            except: return False
            if time.time() - start_time > timeout:
                self.print_aviso(f"Timeout ao aguardar SAP após {timeout} segundos")
                return False
            time.sleep(0.2)
        return False

    def is_session_valid(self):
        if self.session is None: return False
        try:
            self.session.findById("wnd")
            return True
        except (pywintypes.com_error, Exception):
            return False

    def sap_login_handler(self):
        try:
            self.print_info("Procurando por uma sessão SAP GUI...")
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            application = sap_gui_auto.GetScriptingEngine
            
            if application.Connections.Count > 0:
                for conn_idx in range(application.Connections.Count):
                    connection = application.Connections(conn_idx)
                    if connection.Sessions.Count > 0:
                        for session_idx in range(connection.Sessions.Count):
                            session = connection.Sessions(session_idx)
                            try:
                                session.findById("wnd")
                                self.print_sucesso(f"Sessão SAP ativa encontrada.")
                                return session
                            except: continue
            
            self.print_aviso("Nenhuma sessão SAP válida encontrada. Iniciando nova conexão...")
            return self.open_and_login_sap()
        except (pywintypes.com_error, Exception):
            self.print_aviso("Iniciando processo de login...")
            return self.open_and_login_sap()

    # --- SUA FUNÇÃO INTEGRADA AQUI ---
    def open_and_login_sap(self):
        try:
            sap_path = self.config.get('SAP', 'caminho_logon')
            sap_system = self.config.get('SAP', 'sistema').strip()
            
            if not os.path.exists(sap_path):
                self.print_erro(f"Caminho do SAP Logon não encontrado: '{sap_path}'")
                return None
                
            self.print_info(f"Abrindo SAP Logon...")
            subprocess.Popen(sap_path)
            time.sleep(5)
            
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            application = sap_gui_auto.GetScriptingEngine
            
            self.print_info(f"Conectando ao sistema '{sap_system}'...")
            connection = application.OpenConnection(sap_system, True)
            time.sleep(3)
            session = connection.Children(0)
            
            start_time = time.time()
            while session.busy:
                time.sleep(0.5)
                if time.time() - start_time > 30: return None
            
            main_window = session.findById("wnd")
            
            try:
                main_window.findById("usr/txtRSYST-BNAME")
                self.print_info("Preenchendo credenciais...")
                
                # --- Lógica do .env injetada no script que funciona ---
                user = str(os.getenv("SAP_USER") or "").strip().replace('"', '').replace("'", "")
                password = str(os.getenv("SAP_PASSWORD") or "").strip().replace('"', '').replace("'", "")
                
                if not user or not password:
                    self.print_erro("Usuário ou senha não encontrados no arquivo .env!")
                    return None
                # --------------------------------------------------------

                main_window.findById("usr/txtRSYST-BNAME").text = user
                main_window.findById("usr/pwdRSYST-BCODE").text = password
                main_window.sendVKey(0)

                start_time = time.time()
                while session.busy:
                    time.sleep(0.5)
                    if time.time() - start_time > 30: return None
                
                try: session.findById("wnd").sendVKey(0) 
                except: pass
                
                if "easy access" in session.findById("wnd").text.lower() or "menú" in session.findById("wnd").text.lower() or "sap" in session.findById("wnd").text.lower():
                    self.print_sucesso("Login no SAP realizado com sucesso!")
                    return session
                else:
                    self.print_erro(f"Falha no login: {session.findById('sbar').text}")
                    return None
            except:
                self.print_sucesso("Sessão existente detectada.")
                return session

        except Exception as e:
            self.print_erro(f"Erro crítico login: {str(e)}")
            return None

    def processar_lotes(self, df_para_processar, worksheet, status_col_index, req_col_index):
        self.print_info(f"Encontradas {len(df_para_processar)} linhas pendentes.")
        
        # Agrupa por Origem e Destino
        grupos = df_para_processar.groupby(['ORIGEM', 'DESTINO'])
        
        lotes_para_processar = []
        
        for (origem, destino), grupo in grupos:
            # Divide cada grupo em pedaços de 10 linhas
            for i in range(0, len(grupo), 10):
                chunk = grupo.iloc[i : i + 10].copy()
                chunk['grid_index'] = range(len(chunk))
                lotes_para_processar.append(chunk)
        
        total_lotes = len(lotes_para_processar)
        self.print_info(f"Total de RCs a serem criadas (Lotes): {total_lotes}")
        
        for idx, lote_df in enumerate(lotes_para_processar):
            if not self.running: break
            if not self.is_session_valid():
                self.print_erro("Sessão SAP perdida. Tentando reconectar...")
                self.session = self.sap_login_handler()
                if not self.session: break
            
            origem_val = lote_df.iloc['ORIGEM']
            destino_val = lote_df.iloc['DESTINO']
            self.print_header(f"Processando Lote {idx + 1}/{total_lotes} | {origem_val} -> {destino_val}")
            
            # --- Validação ---
            resultados = self.validar_lote_na_rc(lote_df)
            
            # Atualização da Planilha (Validação)
            validation_updates = []
            linhas_ok = []
            for res in resultados:
                validation_updates.append({'range': f'{gspread.utils.rowcol_to_a1(res["linha_planilha"], status_col_index)}', 'values': [[str(res['status'])]]})
                validation_updates.append({'range': f'{gspread.utils.rowcol_to_a1(res["linha_planilha"], req_col_index)}', 'values': [[str(res['numero_rc'])]]})
                if res['status'] == 'OK': linhas_ok.append(res['linha_planilha'])

            if validation_updates:
                try: 
                    worksheet.batch_update(validation_updates)
                except Exception as e: 
                    self.print_erro(f"Erro update planilha: {e}")

            # --- Criação (apenas itens OK) ---
            if not linhas_ok:
                self.print_aviso("Nenhum item válido neste lote. Pulando criação.")
                continue
                
            lote_df_ok = lote_df[lote_df['linha_planilha'].isin(linhas_ok)].copy()
            lote_df_ok['grid_index'] = range(len(lote_df_ok))
            
            if not self.is_session_valid():
                self.print_erro("Sessão SAP perdida.")
                break
            
            numero_rc, msg_status = self.criar_rc_para_lote_ok(lote_df_ok)
            
            # Atualização Final
            creation_updates = []
            for linha in lote_df_ok['linha_planilha']:
                creation_updates.append({'range': f'{gspread.utils.rowcol_to_a1(linha, status_col_index)}', 'values': [[str(msg_status)]]})
                if numero_rc:
                    creation_updates.append({'range': f'{gspread.utils.rowcol_to_a1(linha, req_col_index)}', 'values': [[str(numero_rc)]]})
            
            if creation_updates:
                try:
                    worksheet.batch_update(creation_updates)
                    self.print_sucesso("RC Criada e Planilha atualizada.")
                except Exception as e: 
                    self.print_erro(f"Erro update final: {e}")

    def validar_lote_na_rc(self, lote_de_itens):
        if lote_de_itens.empty: return []
        resultados_finais = []
        try:
            self.print_info(f"Validando Lote ({len(lote_de_itens)} itens)")
            self.session.findById("wnd").maximize()
            self.session.findById("wnd/tbar/okcd").text = "/NME51N"
            self.session.findById("wnd").sendVKey(0)
            self.aguardar_sap()
            time.sleep(1)
            self.session.findById("wnd/usr/subSUB0:SAPLMEGUI:0016/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:3327/cmbMEREQ_TOPLINE-BSART").key = "ZRT"
            self.session.findById("wnd").sendVKey(0)
            self.aguardar_sap()
            grid = self.session.findById("wnd/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell")
            
            for _, item in lote_de_itens.iterrows():
                if not self.running: break
                
                grid_index = int(item['grid_index'])
                
                mat_id = item.get('PN')
                origem = item.get('ORIGEM')
                destino = item.get('DESTINO')
                qtd = str(item.get('QTD', '1')).replace(',', '.')
                texto = item.get('TEXTO')
                
                try:
                    lt_dias = int(str(item.get('LT', 0)).strip() or 0)
                except ValueError:
                    lt_dias = 0
                
                data_remessa = (datetime.now() + timedelta(days=lt_dias)).strftime('%d.%m.%Y')
                
                status_item = "OK"
                print(f" -> Avaliando Item {grid_index + 1} (Mat: {mat_id} | Data: {data_remessa})")
                try:
                    grid.modifyCell(grid_index, "MATNR", str(mat_id))
                    grid.modifyCell(grid_index, "MENGE", qtd)
                    grid.modifyCell(grid_index, "RESWK", str(origem))
                    grid.modifyCell(grid_index, "EEIND", data_remessa)
                    grid.modifyCell(grid_index, "EPSTP", "U")
                    grid.modifyCell(grid_index, "NAME1", str(destino))
                    grid.modifyCell(grid_index, "EKGRP", "P04")
                    grid.modifyCell(grid_index, "TXZ01", str(texto))
                    self.session.findById("wnd").sendVKey(0)
                    self.aguardar_sap()
                    time.sleep(1.5)
                    try: self.session.findById("wnd").sendVKey(0)
                    except: pass
                    
                    status_bar = self.session.findById("wnd/sbar")
                    if status_bar.messageType in ('E', 'A') or "não está atualizado no centro" in status_bar.text:
                        status_item = status_bar.text
                        self.print_erro(f"    Erro: {status_item}")
                    else:
                        status_item = "OK"
                        self.print_sucesso("    Item OK")
                except Exception as e:
                    status_item = f"Erro crítico: {str(e)}"
                    self.print_erro(f"    {status_item}")
                resultados_finais.append({'linha_planilha': item['linha_planilha'], 'status': status_item, 'numero_rc': '' if status_item == 'OK' else 'ERRO'})
            return resultados_finais
        finally:
            try:
                if self.is_session_valid():
                    self.session.findById("wnd/tbar/okcd").text = "/N"
                    self.session.findById("wnd").sendVKey(0)
            except: pass

    def criar_rc_para_lote_ok(self, lote_de_itens_ok):
        if lote_de_itens_ok.empty: return None, "Lote vazio."
        try:
            self.print_info(f"Criando RC para {len(lote_de_itens_ok)} itens aprovados...")
            self.session.findById("wnd").maximize()
            self.session.findById("wnd/tbar/okcd").text = "/NME51N"
            self.session.findById("wnd").sendVKey(0)
            self.aguardar_sap()
            self.session.findById("wnd/usr/subSUB0:SAPLMEGUI:0016/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:3327/cmbMEREQ_TOPLINE-BSART").key = "ZRT"
            self.session.findById("wnd").sendVKey(0)
            self.aguardar_sap()
            grid = self.session.findById("wnd/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell")
            
            lote = lote_de_itens_ok.reset_index(drop=True)
            for i, item in lote.iterrows():
                if not self.running: return None, "Cancelado."
                
                mat_id = item.get('PN')
                origem = item.get('ORIGEM')
                destino = item.get('DESTINO')
                qtd = str(item.get('QTD', '1')).replace(',', '.')
                texto = item.get('TEXTO')

                try:
                    lt_dias = int(str(item.get('LT', 0)).strip() or 0)
                except ValueError:
                    lt_dias = 0
                data_remessa = (datetime.now() + timedelta(days=lt_dias)).strftime('%d.%m.%Y')

                grid.modifyCell(i, "MATNR", str(mat_id))
                grid.modifyCell(i, "MENGE", qtd)
                grid.modifyCell(i, "RESWK", str(origem))
                grid.modifyCell(i, "EEIND", data_remessa)
                grid.modifyCell(i, "EPSTP", "U")
                grid.modifyCell(i, "NAME1", str(destino))
                grid.modifyCell(i, "EKGRP", "P04")
                grid.modifyCell(i, "TXZ01", str(texto))
            
            self.session.findById("wnd").sendVKey(0)
            self.aguardar_sap()
            
            self.print_info("Inserindo Depósitos...")
            for i, item in lote.iterrows():
                if not self.running: return None, "Cancelado."
                
                origem_key = str(item.get('ORIGEM')).strip().upper()
                deposito = self.DEPOSITO_MAPPING.get(origem_key, 'AE01')
                
                grid.setCurrentCell(i, "MATNR")
                self.aguardar_sap()
                self.session.findById("wnd/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT16").select()
                self.session.findById("wnd/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT16/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsTABREITER1/tabpTRANS").select()
                depot = self.session.findById("wnd/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT16/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsTABREITER1/tabpTRANS/ssubSUBBILD1:SAPLXM02:0114/ctxtEBAN-ZZDEP_FORNEC")
                depot.text = str(deposito)
                if i < len(lote) - 1:
                    self.session.findById("wnd/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/btn%#AUTOTEXT002").press()
                    self.aguardar_sap()

            self.print_info("Salvando RC...")
            self.session.findById("wnd/tbar/btn").press()
            self.aguardar_sap()

            try:
                self.session.findById("wnd").sendVKey(0) 
                self.aguardar_sap()
            except: pass
            
            msg = self.session.findById("wnd/sbar").text
            match = re.search(r'(\d{10,})', msg)
            if match:
                rc = match.group(0)
                self.print_sucesso(f"RC Criada: {rc}")
                return rc, msg
            else:
                self.print_erro(f"Falha ao salvar RC: {msg}")
                return None, msg
        except Exception as e:
            return None, f"Erro criação: {e}"

if __name__ == "__main__":
    bot = SAPBotCLI()
    bot.run()