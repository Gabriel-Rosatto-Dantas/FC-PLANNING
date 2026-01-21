import sys
import time
import gspread
import pandas as pd
import win32com.client
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta # Adicionado timedelta para somar dias

# ==========================================
# CONFIGURAÇÕES GERAIS
# ==========================================
class Config:
    GOOGLE_CREDENTIALS_FILE = 'credentials.json' 
    SHEET_NAME = 'MAPEAMENTO PLANNING'
    
    # --- VARIÁVEIS PADRÃO SOLICITADAS ---
    CENTRO_PADRAO = 'BR8E'
    DIAS_PARA_REMESSA = 120  # Data de hoje + 120 dias
    
    # Lista das abas que o script deve percorrer
    ABAS_PARA_PROCESSAR = [
        '0-1500', '1501-5000', '5001-25000', 
        '25001-100000', '100001-200000', '>200000'
    ]

    # Mapeamento dos Grupos de Compradores
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
        
        # Variáveis que serão definidas na execução
        self.grupo_selecionado = None 
        self.data_remessa_calculada = None

    # --- UTILITÁRIOS DE FORMATAÇÃO ---
    @staticmethod
    def format_decimal(val):
        if not val: return ""
        try:
            val_str = str(val).strip().replace('R$', '').replace('$', '').strip()
            if '.' in val_str and ',' in val_str:
                val_str = val_str.replace('.', '').replace(',', '.')
            elif ',' in val_str:
                val_str = val_str.replace(',', '.')
            return "{:.2f}".format(float(val_str)).replace('.', ',')
        except: return str(val)

    def find_column_index(self, headers, col_name):
        try:
            return headers.index(col_name) + 1
        except ValueError:
            col_name_lower = col_name.lower()
            for i, h in enumerate(headers):
                if h.lower() == col_name_lower: return i + 1
            return len(headers) + 1

    # --- CONFIGURAÇÃO INICIAL (NOVO) ---
    def configurar_parametros_execucao(self):
        """Define a data calculada e pede o Grupo ao usuário."""
        
        # 1. Calcular Data (Hoje + 120 dias)
        data_futura = datetime.now() + timedelta(days=Config.DIAS_PARA_REMESSA)
        self.data_remessa_calculada = data_futura.strftime('%d.%m.%Y')
        
        print("\n" + "="*40)
        print(f" DATA REMESSA DEFINIDA: {self.data_remessa_calculada} (+{Config.DIAS_PARA_REMESSA} dias)")
        print(f" CENTRO PADRÃO DEFINIDO: {Config.CENTRO_PADRAO}")
        print("="*40)

        # 2. Menu de Seleção de Grupo
        print("\n>>> SELECIONE O TIPO DE REQUISIÇÃO (GRUPO):")
        # Ordena as chaves para que '0' apareça primeiro
        chaves_ordenadas = sorted(Config.OPCOES_GRUPO.keys())
        for key in chaves_ordenadas:
            info = Config.OPCOES_GRUPO[key]
            print(f" [{key}] - {info['codigo']} ({info['desc']})")
        
        while True:
            escolha = input("\nDigite o número da opção desejada: ").strip()
            if escolha in Config.OPCOES_GRUPO:
                # Verifica se o usuário escolheu SAIR
                if escolha == '0':
                    print("\n Programa finalizado pelo usuário. Até logo!")
                    sys.exit() # Encerra o script

                selecao = Config.OPCOES_GRUPO[escolha]
                self.grupo_selecionado = selecao['codigo']
                print(f"\n OK! Grupo selecionado: {self.grupo_selecionado} - {selecao['desc']}")
                break
            else:
                print(" Opção inválida. Tente novamente.")
        
        print("="*40 + "\n")
        time.sleep(1)

    # --- CONEXÕES ---
    def connect_google(self):
        try:
            scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
            credentials = Credentials.from_service_account_file(Config.GOOGLE_CREDENTIALS_FILE, scopes=scopes)
            self.sheet_client = gspread.authorize(credentials)
            self.workbook = self.sheet_client.open(Config.SHEET_NAME)
            print(f"Planilha '{Config.SHEET_NAME}' conectada.")
            return True
        except Exception as e:
            print(f"Erro Google Sheets: {e}")
            return False

    def connect_sap(self):
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            application = SapGuiAuto.GetScriptingEngine
            connection = application.Children(0)
            self.session = connection.Children(0)
            print("Conectado ao SAP.")
            return True
        except Exception as e:
            print(f"Erro SAP: {e}")
            return False

    def _tratar_popup(self):
        try:
            if self.session.findById("wnd[1]", False):
                self.session.findById("wnd[1]").sendVKey(0)
                return True
        except: pass
        return False

    # --- TRANSAÇÃO ME51N ---
    def create_purchase_requisition_batch(self, batch_rows, is_rotable=False):
        """Executa a ME51N usando as variáveis globais fixas/calculadas."""
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/NME51N"
            self.session.findById("wnd[0]").sendVKey(0)
            
            grid = self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell")
            
            itens_com_erro = []
            linhas_preenchidas = 0

            for row in batch_rows:
                # 1. Preenchimento dos dados do item
                material = str(row.get('Material', '')).strip()
                try:
                    # O índice no grid sempre será 'linhas_preenchidas' para o novo item
                    if is_rotable: grid.modifyCell(linhas_preenchidas, "KNTTP", "P")
                    
                    # --- ALTERAÇÃO AQUI: USANDO VARIÁVEIS FIXAS/CALCULADAS ---
                    # Usa o Centro definido no Config
                    grid.modifyCell(linhas_preenchidas, "NAME1", Config.CENTRO_PADRAO) 
                    
                    # Material continua vindo da planilha
                    grid.modifyCell(linhas_preenchidas, "MATNR", material)
                    
                    # Qtd e Preço continuam vindo da planilha
                    grid.modifyCell(linhas_preenchidas, "MENGE", self.format_decimal(row.get('Qtd', '')))
                    grid.modifyCell(linhas_preenchidas, "PREIS", self.format_decimal(row.get('Preço', '')))
                    
                    # Usa a Data calculada (Hoje + 120 dias)
                    grid.modifyCell(linhas_preenchidas, "EEIND", self.data_remessa_calculada)
                    
                    # Usa o Grupo selecionado pelo usuário no início
                    grid.modifyCell(linhas_preenchidas, "EKGRP", self.grupo_selecionado)
                    
                    grid.modifyCell(linhas_preenchidas, "WAERS", "USD")

                    # 2. Validação Individual (Enter + Botão Verificar)
                    grid.currentCellRow = linhas_preenchidas
                    self.session.findById("wnd[0]").sendVKey(0) # Enter para processar linha
                    self._tratar_popup()
                    
                    # Clica no botão Verificar (Check)
                    self.session.findById("wnd[0]/tbar[1]/btn[9]").press()
                    self._tratar_popup()

                    # 3. Analisa erros
                    sbar = self.session.findById("wnd[0]/sbar")
                    if sbar.MessageType == "E":
                        erro_msg = sbar.Text
                        print(f" !!! Item {material} com erro: {erro_msg}. Removendo...")
                        itens_com_erro.append(f"{material}: {erro_msg}")
                        
                        # Seleciona a linha e deleta
                        grid.selectedRows = str(linhas_preenchidas)
                        self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/tbar[0]/btn[14]").press()
                    else:
                        linhas_preenchidas += 1 

                except Exception as e:
                    print(f"Erro ao processar item {material}: {e}")
                    continue

            # 4. Finalização do Lote
            if linhas_preenchidas == 0:
                return f"Erro: Nenhum item do lote passou na validação. ({', '.join(itens_com_erro)})"

            # Salvar Requisição
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            self._tratar_popup()
            
            status_sap = self.session.findById("wnd[0]/sbar").Text
            
            if itens_com_erro:
                return f"{status_sap} | Itens Ignorados: {len(itens_com_erro)}"
            return status_sap

        except Exception as e:
            return f"Erro Crítico: {str(e)}"

    # --- LOOP POR TODAS AS ABAS ---
    def run(self):
        # 1. Conexão Google
        if not self.connect_google(): return
        
        # 2. Configura Parâmetros (Pergunta o Grupo e Calcula Data)
        self.configurar_parametros_execucao()
        
        # 3. Conexão SAP
        if not self.connect_sap(): return

        for nome_aba in Config.ABAS_PARA_PROCESSAR:
            print(f"\n>>> ACESSANDO ABA: {nome_aba}")
            try:
                worksheet = self.workbook.worksheet(nome_aba)
                data = worksheet.get_all_records()
                headers = worksheet.row_values(1)
                col_status_idx = self.find_column_index(headers, 'Status')

                items_to_process = []
                for i, row in enumerate(data):
                    status = str(row.get('Status', '')).strip()
                    if status == '' or 'NAO' in status.upper():
                        items_to_process.append({'sheet_row': i + 2, 'data': row})

                if not items_to_process:
                    print(f"Aba {nome_aba} sem itens pendentes.")
                    continue

                # --- CONFIGURAÇÃO DE LOTE POR ABA ---
                # Se for aba de alto valor, processa 1 por 1
                if nome_aba in ['100001-200000', '>200000']:
                    BATCH_SIZE = 1
                    print(f" [!] Aba de alto valor detectada. Modo: Item a item (1 por vez).")
                else:
                    BATCH_SIZE = 10
                    print(f" [i] Aba padrão. Modo: Lote de {BATCH_SIZE} itens.")

                for i in range(0, len(items_to_process), BATCH_SIZE):
                    chunk = items_to_process[i : i + BATCH_SIZE]
                    batch_data = [item['data'] for item in chunk]
                    
                    print(f" - Processando lote {i//BATCH_SIZE + 1} de {nome_aba}...")
                    
                    # O Grupo e Data agora são lidos de self.grupo_selecionado e self.data_remessa_calculada dentro da função
                    resultado = self.create_purchase_requisition_batch(batch_data, is_rotable=False)
                    
                    # Atualiza Status na Planilha
                    for item in chunk:
                        worksheet.update_cell(item['sheet_row'], col_status_idx, resultado)
                        
            except Exception as e:
                print(f"Erro ao processar aba {nome_aba}: {e}")

        print("\nAutomação concluída em todas as abas.")

if __name__ == "__main__":
    app = SAPAutomation()
    app.run()