import sys
import time
import gspread
import win32com.client
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta

# ==========================================
# CONFIGURAÇÕES GERAIS
# ==========================================
class Config:
    GOOGLE_CREDENTIALS_FILE = 'credentials.json' 
    SHEET_NAME = 'MAPEAMENTO PLANNING'
    
    # NOME DA ABA ONDE ESTÃO TODOS OS DADOS AGORA
    NOME_ABA_DADOS = 'DADOS_GERAIS' 
    
    # --- VARIÁVEIS PADRÃO ---
    CENTRO_PADRAO = 'BR8E'
    DIAS_PARA_REMESSA = 120  # Data de hoje + 120 dias
    
    # Mapeamento dos Grupos de Requisição
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
        self.worksheet = None # Referência da aba ativa
        self.grupo_selecionado = None 
        self.data_remessa_calculada = None

    # --- UTILITÁRIOS ---
    @staticmethod
    def format_decimal_sap(val):
        """Formata para STRING no padrão SAP (vírgula decimal)"""
        if not val: return ""
        try:
            val_str = str(val).strip().replace('R$', '').replace('$', '').strip()
            if '.' in val_str and ',' in val_str:
                val_str = val_str.replace('.', '').replace(',', '.')
            elif ',' in val_str:
                val_str = val_str.replace(',', '.')
            # Formata com 2 casas e troca ponto por vírgula para o SAP
            return "{:.2f}".format(float(val_str)).replace('.', ',')
        except: return str(val)

    @staticmethod
    def _parse_price_to_float(val):
        """Converte string de preço para FLOAT Python para fazer comparações"""
        try:
            if isinstance(val, (int, float)): return float(val)
            val_str = str(val).strip().replace('R$', '').replace('$', '').strip()
            # Remove separador de milhar (.) e troca decimal (,) por ponto
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
        """Helper para atualizar a planilha com retry simples em caso de erro de API"""
        try:
            self.worksheet.update_cell(item['sheet_row_index'], col_idx, msg)
        except Exception as e:
            time.sleep(2) # Espera cota da API
            try: self.worksheet.update_cell(item['sheet_row_index'], col_idx, msg)
            except: pass

    # --- LÓGICA DE CLASSIFICAÇÃO DE PREÇO ---
    def classificar_faixa_preco(self, preco_float):
        """Retorna a chave da categoria e o tamanho do lote recomendado"""
        p = preco_float
        
        # Definição das faixas conforme sua regra original
        if p <= 1500:
            return '0-1500', 10
        elif p <= 5000:
            return '1501-5000', 10
        elif p <= 25000:
            return '5001-25000', 10
        elif p <= 100000:
            return '25001-100000', 10
        elif p <= 200000:
            return '100001-200000', 1  # Lote de 1 (Alto Valor)
        else:
            return '>200000', 1       # Lote de 1 (Alto Valor)

    # --- SETUP INICIAL ---
    def configurar_parametros_execucao(self):
        # 1. Calcular Data
        data_futura = datetime.now() + timedelta(days=Config.DIAS_PARA_REMESSA)
        self.data_remessa_calculada = data_futura.strftime('%d.%m.%Y')
        
        print("\n" + "="*40)
        print(f" DATA REMESSA DEFINIDA: {self.data_remessa_calculada}")
        print("="*40)

        # 2. Menu de Seleção
        print("\n>>> SELECIONE O TIPO DE REQUISIÇÃO (GRUPO):")
        chaves_ordenadas = sorted(Config.OPCOES_GRUPO.keys())
        for key in chaves_ordenadas:
            info = Config.OPCOES_GRUPO[key]
            print(f" [{key}] - {info['codigo']} ({info['desc']})")
        
        while True:
            escolha = input("\nDigite o número da opção: ").strip()
            if escolha in Config.OPCOES_GRUPO:
                if escolha == '0':
                    print("Encerrando.")
                    sys.exit()
                selecao = Config.OPCOES_GRUPO[escolha]
                self.grupo_selecionado = selecao['codigo']
                print(f" Grupo selecionado: {self.grupo_selecionado}")
                break
            else:
                print(" Opção inválida.")
        time.sleep(1)

    # --- CONEXÕES ---
    def connect_google(self):
        try:
            scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
            creds = Credentials.from_service_account_file(Config.GOOGLE_CREDENTIALS_FILE, scopes=scopes)
            self.sheet_client = gspread.authorize(creds)
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

    # --- TRATAMENTO DE POPUPS ---
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

    # --- TRANSAÇÃO ME51N ---
    def create_purchase_requisition_batch(self, batch_rows):
        try:
            # Reinicia transação
            self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/NME51N"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1) 
            
            grid_id = "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell"
            
            if not self.session.findById(grid_id, False):
                return "Erro: Grid ME51N não carregou."
            
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
                    print(f"Aviso linha {i}: {e}")

            if linhas_preenchidas == 0:
                return "Erro: Nenhuma linha preenchida."

            # Processamento
            self.session.findById("wnd[0]").sendVKey(0)
            self._lidar_com_popups() 

            try:
                self.session.findById("wnd[0]/tbar[1]/btn[9]").press() # Check
                self._lidar_com_popups() 
            except: pass

            sbar = self.session.findById("wnd[0]/sbar")
            if sbar.MessageType == "E":
                return f"Erro SAP: {sbar.Text}"

            self.session.findById("wnd[0]/tbar[0]/btn[11]").press() # Save
            self._lidar_com_popups(max_tentativas=5)

            sbar_final = self.session.findById("wnd[0]/sbar")
            texto_final = sbar_final.Text
            
            # Verifica sucesso (padrão SAP: Mensagem tipo S, ou texto contendo 'criada'/'gravada')
            if sbar_final.MessageType == "S" or "criad" in texto_final.lower() or "creat" in texto_final.lower() or "gravad" in texto_final.lower():
                return texto_final
            else:
                return f"Status Final: {texto_final}"

        except Exception as e:
            return f"Erro Crítico Script: {str(e)}"

    # --- LOOP PRINCIPAL REFORMULADO ---
    def run(self):
        if not self.connect_google(): return
        self.configurar_parametros_execucao()
        if not self.connect_sap(): return

        print(f"\n>>> LENDO DADOS DA ABA: {Config.NOME_ABA_DADOS}")
        
        try:
            self.worksheet = self.workbook.worksheet(Config.NOME_ABA_DADOS)
            data = self.worksheet.get_all_records()
        except gspread.exceptions.WorksheetNotFound:
            print(f"ERRO: Aba '{Config.NOME_ABA_DADOS}' não encontrada!")
            return

        if not data:
            print("Planilha vazia.")
            return

        headers = self.worksheet.row_values(1)
        col_status_idx = self.find_column_index(headers, 'Status')

        # 1. Separar itens pendentes
        itens_pendentes = []
        for i, row in enumerate(data):
            status = str(row.get('Status', '')).strip()
            # Se status vazio ou contendo NAO (para reprocessamento), adiciona
            if status == '' or 'NAO' in status.upper():
                # Guardamos o indice real da planilha (i+2 pois i começa em 0 e tem header)
                # para atualizar o status depois
                row['sheet_row_index'] = i + 2
                itens_pendentes.append(row)

        if not itens_pendentes:
            print("Nenhum item pendente para processar.")
            return

        print(f"Total de itens pendentes: {len(itens_pendentes)}")

        # 2. Agrupar por faixa de preço (Buckets)
        grupos_processamento = {}
        
        for item in itens_pendentes:
            preco_float = self._parse_price_to_float(item.get('Preço', 0))
            faixa_nome, tamanho_lote = self.classificar_faixa_preco(preco_float)
            
            if faixa_nome not in grupos_processamento:
                grupos_processamento[faixa_nome] = {
                    'batch_size': tamanho_lote,
                    'items': []
                }
            grupos_processamento[faixa_nome]['items'].append(item)

        # 3. Processar cada grupo
        # Ordenamos as chaves para processar do menor valor para o maior (opcional)
        for faixa_nome in sorted(grupos_processamento.keys()):
            grupo = grupos_processamento[faixa_nome]
            items = grupo['items']
            batch_size = grupo['batch_size']
            
            print(f"\n>>> PROCESSANDO FAIXA: {faixa_nome}")
            print(f"    Quantidade de itens: {len(items)}")
            print(f"    Tamanho do lote: {batch_size}")

            # Divide os itens do grupo em lotes
            for i in range(0, len(items), batch_size):
                chunk = items[i : i + batch_size]
                
                print(f" - Processando lote {i//batch_size + 1} de {faixa_nome}...")
                
                # --- TENTATIVA 1: LOTE COMPLETO ---
                resultado = self.create_purchase_requisition_batch(chunk)
                
                # Checagem de sucesso baseada no texto de retorno
                sucesso = "criad" in resultado.lower() or "creat" in resultado.lower() or "gravad" in resultado.lower()
                
                # --- LÓGICA DE FALLBACK (ISOLAMENTO DE ERRO) ---
                if not sucesso and len(chunk) > 1:
                    print(f"   [!] Erro no lote ({resultado}). Tentando processar item a item para isolar o erro...")
                    
                    for sub_item in chunk:
                        # Tenta processar o item individualmente
                        res_indiv = self.create_purchase_requisition_batch([sub_item])
                        print(f"      > Item {sub_item.get('Material')}: {res_indiv}")
                        
                        # Atualiza planilha individualmente
                        self._atualizar_status_planilha(sub_item, col_status_idx, res_indiv)
                else:
                    # Se deu sucesso, OU se falhou mas já era lote de 1 (não dá pra dividir)
                    print(f"   Resultado SAP: {resultado}")
                    for item in chunk:
                        self._atualizar_status_planilha(item, col_status_idx, resultado)

        print("\nAutomação concluída.")

if __name__ == "__main__":
    app = SAPAutomation()
    app.run()