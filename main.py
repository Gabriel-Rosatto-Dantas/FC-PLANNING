import sys
import time
import logging
from logging.handlers import RotatingFileHandler
import gspread
import win32com.client
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
import os

# ==========================================
# CONFIGURAÇÕES GERAIS
# ==========================================
class Config:
    GOOGLE_CREDENTIALS_FILE = 'credentials.json' 
    SHEET_NAME = 'MAPEAMENTO PLANNING'
    NOME_ABA_DADOS = 'BD GERAL'    
    
    # --- VARIÁVEIS PADRÃO ---
    CENTRO_PADRAO = 'BR8E'
    DIAS_PARA_REMESSA = 120
    
    # --- ID DO GRID SAP ATUALIZADO (BASEADO NO LOG) ---
    # Atualizado de ...0013 para ...0016 conforme log de debug
    GRID_ID_PADRAO = "wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell"

    OPCOES_GRUPO = {
        '1': {'codigo': 'P01', 'desc': 'Recomendação'},
        '2': {'codigo': 'P02', 'desc': 'Retorno de Itens'},
        '3': {'codigo': 'P03', 'desc': 'Sinergia'},
        '4': {'codigo': 'P04', 'desc': 'MRP'},
        '5': {'codigo': 'P05', 'desc': 'EO'},
        '6': {'codigo': 'P06', 'desc': 'IP / Projetos'},
        '7': {'codigo': 'P07', 'desc': 'Reposição Scrap'},
        '0': {'codigo': 'SAIR', 'desc': 'Finalizar Programa'}
    }

# ==========================================
# CLASSE PRINCIPAL DE AUTOMAÇÃO
# ==========================================
class SAPAutomation:
    def __init__(self):
        self.session = None
        self.sheet_client = None
        self.workbook = None
        self.worksheet = None 
        self.grupo_selecionado = None 
        self.data_remessa_calculada = None
        self.logger = logging.getLogger(__name__)

    # --- UTILITÁRIOS ---
    @staticmethod
    def format_decimal_sap(val):
        if not val: return ""
        try:
            val_str = str(val).strip().replace('R$', '').replace('$', '').strip()
            if '.' in val_str and ',' in val_str:
                val_str = val_str.replace('.', '').replace(',', '.')
            elif ',' in val_str:
                val_str = val_str.replace(',', '.')
            return "{:.2f}".format(float(val_str)).replace('.', ',')
        except: return str(val)

    @staticmethod
    def _parse_price_to_float(val):
        try:
            if isinstance(val, (int, float)): return float(val)
            val_str = str(val).strip().replace('R$', '').replace('$', '').strip()
            if '.' in val_str and ',' in val_str:
                val_str = val_str.replace('.', '').replace(',', '.')
            elif ',' in val_str:
                val_str = val_str.replace(',', '.')
            return float(val_str)
        except:
            return 0.0

    def find_column_index(self, headers, col_name):
        try:
            return headers.index(col_name) + 1
        except ValueError:
            col_name_lower = col_name.lower()
            for i, h in enumerate(headers):
                if h.lower() == col_name_lower: return i + 1
            return len(headers) + 1

    def _atualizar_status_planilha(self, item, col_idx, msg):
        try:
            self.worksheet.update_cell(item['sheet_row_index'], col_idx, msg)
        except Exception as e:
            self.logger.warning("Erro atualizando planilha (tentando retry): %s", e)
            time.sleep(2)
            try:
                self.worksheet.update_cell(item['sheet_row_index'], col_idx, msg)
            except Exception:
                self.logger.exception("Falha no retry ao atualizar planilha.")

    def classificar_faixa_preco(self, preco_float):
        p = preco_float
        if p <= 1500: return '0-1500', 10
        elif p <= 5000: return '1501-5000', 10
        elif p <= 25000: return '5001-25000', 10
        elif p <= 100000: return '25001-100000', 10
        elif p <= 200000: return '100001-200000', 1
        else: return '>200000', 1

    def configurar_parametros_execucao(self):
        data_futura = datetime.now() + timedelta(days=Config.DIAS_PARA_REMESSA)
        self.data_remessa_calculada = data_futura.strftime('%d.%m.%Y')
        
        self.logger.info("%s", "\n" + "="*40)
        self.logger.info(" DATA REMESSA DEFINIDA: %s", self.data_remessa_calculada)
        self.logger.info("%s", "="*40)

        self.logger.info("\n>>> SELECIONE O TIPO DE REQUISIÇÃO (GRUPO):")
        chaves_ordenadas = sorted(Config.OPCOES_GRUPO.keys())
        for key in chaves_ordenadas:
            info = Config.OPCOES_GRUPO[key]
            self.logger.info(" [%s] - %s (%s)", key, info['codigo'], info['desc'])
        
        while True:
            escolha = input("\nDigite o número da opção: ").strip()
            if escolha in Config.OPCOES_GRUPO:
                if escolha == '0':
                    self.logger.info("Encerrando.")
                    sys.exit()
                selecao = Config.OPCOES_GRUPO[escolha]
                self.grupo_selecionado = selecao['codigo']
                self.logger.info(" Grupo selecionado: %s", self.grupo_selecionado)
                break
            else:
                self.logger.warning(" Opção inválida: %s", escolha)
        time.sleep(1)

    # --- CONEXÕES ---
    def connect_google(self):
        try:
            scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
            creds = Credentials.from_service_account_file(Config.GOOGLE_CREDENTIALS_FILE, scopes=scopes)
            self.sheet_client = gspread.authorize(creds)
            self.workbook = self.sheet_client.open(Config.SHEET_NAME)
            self.logger.info("Planilha '%s' conectada.", Config.SHEET_NAME)
            return True
        except Exception as e:
            self.logger.exception("Erro Google Sheets: %s", e)
            return False

    def connect_sap(self):
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            application = SapGuiAuto.GetScriptingEngine
            connection = application.Children(0)
            self.session = connection.Children(0)
            self.logger.info("Conectado ao SAP.")
            return True
        except Exception as e:
            self.logger.exception("Erro SAP: %s", e)
            return False

    def _lidar_com_popups(self, max_tentativas=4):
        time.sleep(0.5) 
        for _ in range(max_tentativas):
            try:
                if self.session.findById("wnd[1]", False): 
                    self.session.findById("wnd[1]").sendVKey(0)
                    time.sleep(0.5)
                else:
                    break
            except:
                break

    # --- DEBUGGER DE IDS SAP (NOVIDADE) ---
    def debug_encontrar_ids(self, obj, profundidade=0):
        """Função recursiva para encontrar onde o Grid se escondeu"""
        try:
            if profundidade > 10: return
            
            # Tenta listar os filhos deste objeto
            children = obj.Children
            for child in children:
                try:
                    child_id = child.Id
                    child_type = child.Type
                    
                    # Se for um Container Shell (onde grids ficam), avisa!
                    if "shellcont" in child_id or "GRID" in child_id:
                        self.logger.info(f" [DEBUG] ENCONTRADO: {child_id} (Tipo: {child_type})")
                    
                    # Recursão para olhar dentro
                    self.debug_encontrar_ids(child, profundidade + 1)
                except:
                    pass
        except:
            pass

    # --- TRANSAÇÃO ME51N ---
    def create_purchase_requisition_batch(self, batch_rows):
        try:
            # Reinicia transação
            self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/NME51N"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1.5) 
            
            # Tenta expandir a "Síntese de Itens" se estiver fechada (botão comum)
            # Isso ajuda a estabilizar o ID
            try:
                btn_sintese = self.session.findById("wnd[0]/tbar[1]/btn[14]", False)
                if btn_sintese: 
                    btn_sintese.press()
                    time.sleep(1)
            except:
                pass

            grid_id = Config.GRID_ID_PADRAO
            
            # Verifica se o grid existe
            if not self.session.findById(grid_id, False):
                self.logger.error("ERRO: Grid ME51N não carregou com o ID padrão.")
                self.logger.info(">>> INICIANDO BUSCA DE IDS (DEBUG)...")
                self.logger.info(">>> Copie um dos IDs abaixo que contenha 'shellcont/shell' e substitua no código.")
                
                # Procura dentro da área de usuário (usr)
                usr_area = self.session.findById("wnd[0]/usr")
                self.debug_encontrar_ids(usr_area)
                
                return "Erro: Grid não encontrado (Verifique o LOG)"
            
            grid = self.session.findById(grid_id)
            
            linhas_preenchidas = 0
            for i, row in enumerate(batch_rows):
                try:
                    material = str(row.get('Material', '')).strip()
                    qtd = self.format_decimal_sap(row.get('Qtd', ''))
                    preco = self.format_decimal_sap(row.get('Preço', ''))

                    try: grid.setCurrentCell(i, "MATNR")
                    except: pass

                    grid.modifyCell(i, "NAME1", Config.CENTRO_PADRAO)
                    grid.modifyCell(i, "MATNR", material)
                    grid.modifyCell(i, "MENGE", qtd)
                    grid.modifyCell(i, "PREIS", preco)
                    grid.modifyCell(i, "EEIND", self.data_remessa_calculada)
                    grid.modifyCell(i, "EKGRP", self.grupo_selecionado)
                    grid.modifyCell(i, "WAERS", "USD")
                    
                    linhas_preenchidas += 1
                except Exception as e:
                    self.logger.warning("Aviso linha %s: %s", i, e)

            if linhas_preenchidas == 0:
                return "Erro: Nenhuma linha preenchida."

            # Processamento
            self.session.findById("wnd[0]").sendVKey(0)
            self._lidar_com_popups() 

            try:
                # Botão Check
                self.session.findById("wnd[0]/tbar[1]/btn[9]").press() 
                self._lidar_com_popups() 
            except: pass

            sbar = self.session.findById("wnd[0]/sbar")
            if sbar.MessageType == "E":
                self.logger.error("Erro SAP (sbar): %s", sbar.Text)
                return f"Erro SAP: {sbar.Text}"

            # Botão Save
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press() 
            self._lidar_com_popups(max_tentativas=5)

            sbar_final = self.session.findById("wnd[0]/sbar")
            texto_final = sbar_final.Text
            
            if sbar_final.MessageType == "S" or any(x in texto_final.lower() for x in ['criad', 'creat', 'gravad']):
                self.logger.info("Operação SAP concluída: %s", texto_final)
                return texto_final
            else:
                self.logger.warning("Status final SAP: %s", texto_final)
                return f"Status Final: {texto_final}"

        except Exception as e:
            self.logger.exception("Erro Crítico Script: %s", e)
            return f"Erro Crítico Script: {str(e)}"

    def run(self):
        if not self.connect_google(): return
        self.configurar_parametros_execucao()
        if not self.connect_sap(): return

        self.logger.info("\n>>> LENDO DADOS DA ABA: %s", Config.NOME_ABA_DADOS)
        
        try:
            self.worksheet = self.workbook.worksheet(Config.NOME_ABA_DADOS)
            data = self.worksheet.get_all_records()
        except gspread.exceptions.WorksheetNotFound:
            self.logger.error("ERRO: Aba '%s' não encontrada!", Config.NOME_ABA_DADOS)
            return

        if not data: return

        headers = self.worksheet.row_values(1)
        col_status_idx = self.find_column_index(headers, 'Status')

        itens_pendentes = []
        for i, row in enumerate(data):
            status = str(row.get('Status', '')).strip()
            if status == '' or 'NAO' in status.upper():
                row['sheet_row_index'] = i + 2
                itens_pendentes.append(row)

        if not itens_pendentes:
            self.logger.info("Nenhum item pendente para processar.")
            return

        self.logger.info("Total de itens pendentes: %s", len(itens_pendentes))

        grupos_processamento = {}
        for item in itens_pendentes:
            preco_float = self._parse_price_to_float(item.get('Preço', 0))
            faixa_nome, tamanho_lote = self.classificar_faixa_preco(preco_float)
            
            if faixa_nome not in grupos_processamento:
                grupos_processamento[faixa_nome] = {'batch_size': tamanho_lote, 'items': []}
            grupos_processamento[faixa_nome]['items'].append(item)

        for faixa_nome in sorted(grupos_processamento.keys()):
            grupo = grupos_processamento[faixa_nome]
            items = grupo['items']
            batch_size = grupo['batch_size']
            
            self.logger.info("\n>>> PROCESSANDO FAIXA: %s", faixa_nome)
            
            for i in range(0, len(items), batch_size):
                chunk = items[i : i + batch_size]
                self.logger.info(" - Processando lote %s...", i//batch_size + 1)
                
                resultado = self.create_purchase_requisition_batch(chunk)
                sucesso = any(x in resultado.lower() for x in ['criad', 'creat', 'gravad'])
                
                if not sucesso and len(chunk) > 1:
                    self.logger.warning("   [!] Erro no lote. Tentando item a item...")
                    for sub_item in chunk:
                        res_indiv = self.create_purchase_requisition_batch([sub_item])
                        self.logger.info("      > Item %s: %s", sub_item.get('Material'), res_indiv)
                        self._atualizar_status_planilha(sub_item, col_status_idx, res_indiv)
                else:
                    self.logger.info("   Resultado: %s", resultado)
                    for item in chunk:
                        self._atualizar_status_planilha(item, col_status_idx, resultado)

        self.logger.info("\nAutomação concluída.")

def setup_logging():
    base = os.path.dirname(os.path.abspath(__file__))
    log_file = os.path.join(base, 'fc_planning.log')
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(levelname)s [%(name)s] %(message)s',
        handlers=[
            RotatingFileHandler(log_file, maxBytes=5*1024*1024, backupCount=5, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )

if __name__ == "__main__":
    setup_logging()
    app = SAPAutomation()
    app.run()