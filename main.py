import sys
import time
import logging
import re
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
    
    # --- ID DO GRID SAP (ITENS) ---
    GRID_ID_PADRAO = "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell"

    # --- ID DO EDITOR DE TEXTO ---
    ID_EDITOR_TEXTO = "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/tabpTABREQHDT1/ssubTABSTRIPCONTROL3SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell"

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
        self.grupo_descricao = None 
        self.data_remessa_calculada = None
        self.logger = logging.getLogger(__name__)

    # --- UTILITÁRIOS ---
    @staticmethod
    def format_decimal_sap(val):
        """ 
        Recebe o valor (agora sempre STRING vindo do get_all_values)
        e garante a formatação '0,27'.
        """
        if val is None or val == "": return "0,00"
        
        try:
            # 1. Garante que é string e limpa espaços/moeda
            val_str = str(val).strip().replace('R$', '').replace('$', '').strip()
            
            # 2. Tratamento para converter texto BR "0,27" para Float Python 0.27
            # Se tiver ponto de milhar (ex 1.000,00), remove o ponto
            if '.' in val_str and ',' in val_str:
                val_str = val_str.replace('.', '')
            
            # Troca a vírgula decimal por ponto para o Python entender
            val_str = val_str.replace(',', '.')
            
            # 3. Converte para float matemático
            val_float = float(val_str)
            
            # 4. Formata de volta para String com VÍRGULA (Padrão SAP BR)
            # {:.2f} gera "0.27", replace troca para "0,27"
            return "{:.2f}".format(val_float).replace('.', ',')
            
        except Exception as e:
            # Se falhar, retorna string original trocando ponto por virgula por segurança
            return str(val).replace('.', ',')

    @staticmethod
    def _parse_price_to_float(val):
        """ Converte para float apenas para lógica interna de Lotes """
        try:
            val_str = str(val).strip().replace('R$', '').replace('$', '').strip()
            if '.' in val_str and ',' in val_str:
                val_str = val_str.replace('.', '')
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

    def _atualizar_status_planilha(self, row_index, col_idx, msg):
        try:
            self.worksheet.update_cell(row_index, col_idx, msg)
        except Exception:
            time.sleep(2)
            try:
                self.worksheet.update_cell(row_index, col_idx, msg)
            except Exception: pass

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
                self.grupo_descricao = selecao['desc']
                self.logger.info(" Grupo selecionado: %s (%s)", self.grupo_selecionado, self.grupo_descricao)
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

    # --- TRANSAÇÃO ME51N ---
    def create_purchase_requisition_batch(self, batch_rows):
        try:
            # 1. Inicia Transação (/NME51N)
            self.session.findById("wnd[0]").maximize()
            self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/NME51N"
            self.session.findById("wnd[0]").sendVKey(0)
            
            time.sleep(2) 

            # 2. ESCREVE O TEXTO DE CABEÇALHO
            data_hoje = datetime.now().strftime('%d.%m.%Y')
            texto_final = f"Compra para Atender demanda {self.grupo_descricao}\r\n{data_hoje}\r\n"
            
            try:
                self.session.findById(Config.ID_EDITOR_TEXTO).text = texto_final
                try: self.session.findById(Config.ID_EDITOR_TEXTO).setSelectionIndexes(92, 92)
                except: pass
                self.logger.info("Texto de cabeçalho preenchido.")
            except Exception as e:
                self.logger.warning(f"Erro ao preencher texto (ID correto?): {e}")

            # 3. PREENCHE O GRID (ITENS)
            grid = self.session.findById(Config.GRID_ID_PADRAO)
            
            linhas_preenchidas = 0
            for i, row in enumerate(batch_rows):
                try:
                    material = str(row.get('Material', '')).strip()
                    
                    # LOG DE DEBUG
                    valor_bruto = row.get('Preço', '')
                    self.logger.info(f" -> Item {i+1} Valor BRUTO (Texto): '{valor_bruto}'")
                    
                    # FORMATAÇÃO
                    qtd = self.format_decimal_sap(row.get('Qtd', ''))
                    preco = self.format_decimal_sap(valor_bruto)
                    
                    self.logger.info(f" -> Enviando para SAP: Mat={material}, Qtd={qtd}, Preço={preco}")
                    
                    try: grid.modifyCell(i, "NAME1", Config.CENTRO_PADRAO)
                    except: pass 
                    
                    grid.modifyCell(i, "MATNR", material)
                    grid.modifyCell(i, "MENGE", qtd)
                    grid.modifyCell(i, "PREIS", preco)
                    grid.modifyCell(i, "EEIND", self.data_remessa_calculada)
                    grid.modifyCell(i, "EKGRP", self.grupo_selecionado)
                    grid.modifyCell(i, "WAERS", "USD")
                    
                    linhas_preenchidas += 1
                except Exception as e:
                    self.logger.warning(f"Erro linha {i}: {e}")

            if linhas_preenchidas == 0:
                return "Erro: Nenhuma linha preenchida."

            # 4. FINALIZA O GRID
            try:
                grid.currentCellColumn = "WAERS"
                grid.pressEnter()
            except:
                self.session.findById("wnd[0]").sendVKey(0)

            # 5. TRATA O POPUP
            time.sleep(1)
            try:
                if self.session.findById("wnd[1]", False):
                    self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except: pass

            # 6. GRAVAR
            self.logger.info("Gravando...")
            try:
                self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            except Exception as e:
                self.logger.error(f"Erro ao pressionar Gravar: {e}")

            # TRATA POPUPS FINAIS
            time.sleep(1)
            try:
                if self.session.findById("wnd[1]", False):
                     self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except: pass

            # 7. CAPTURA MENSAGEM FINAL (SÓ NÚMERO)
            sbar = self.session.findById("wnd[0]/sbar")
            texto_status = sbar.Text
            
            if sbar.MessageType == "S" or any(x in texto_status.lower() for x in ['criad', 'creat', 'gravad']):
                self.logger.info("Sucesso (Log): %s", texto_status)
                
                try: self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                except: pass
                
                numeros = re.findall(r'\d+', texto_status)
                if numeros:
                    return numeros[-1]
                else:
                    return texto_status
            else:
                self.logger.warning("Status: %s", texto_status)
                return f"Status Final: {texto_status}"

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
            
            # --- MUDANÇA CRÍTICA AQUI ---
            # get_all_values() retorna TUDO como String (Lista de Listas)
            # Evita que o Google Sheets converta "0,27" para int 27
            raw_data = self.worksheet.get_all_values()
            
            if not raw_data or len(raw_data) < 2:
                self.logger.info("Planilha vazia ou sem dados.")
                return

            # Reconstrói a estrutura de dicionário manualmente
            headers = raw_data[0]
            data = []
            for row_vals in raw_data[1:]:
                row_dict = {}
                for i, header in enumerate(headers):
                    val = row_vals[i] if i < len(row_vals) else ""
                    row_dict[header] = val
                data.append(row_dict)

        except Exception as e:
            self.logger.error(f"Erro ao ler planilha: {e}")
            return

        col_status_idx = self.find_column_index(headers, 'Status')

        itens_pendentes = []
        for i, row in enumerate(data):
            status = str(row.get('Status', '')).strip()
            # row index para update tem que considerar cabeçalho (+1) e indice 0 (+1) = +2
            row['sheet_row_index'] = i + 2
            
            if status == '' or 'NAO' in status.upper():
                itens_pendentes.append(row)

        if not itens_pendentes:
            self.logger.info("Nenhum item pendente.")
            return

        self.logger.info("Itens pendentes: %s", len(itens_pendentes))

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
            
            self.logger.info("\n>>> FAIXA: %s", faixa_nome)
            
            for i in range(0, len(items), batch_size):
                chunk = items[i : i + batch_size]
                self.logger.info(" - Lote %s...", i//batch_size + 1)
                
                resultado = self.create_purchase_requisition_batch(chunk)
                
                eh_numero = resultado.isdigit()
                sucesso = eh_numero or any(x in resultado.lower() for x in ['criad', 'creat', 'gravad'])
                
                if not sucesso and len(chunk) > 1:
                    for sub_item in chunk:
                        res_indiv = self.create_purchase_requisition_batch([sub_item])
                        self._atualizar_status_planilha(sub_item['sheet_row_index'], col_status_idx, res_indiv)
                else:
                    for item in chunk:
                        self._atualizar_status_planilha(item['sheet_row_index'], col_status_idx, resultado)

        self.logger.info("\nFim.")

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