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
    
    # --- VARIÁVEIS PADRÃO ---
    CENTRO_PADRAO = 'BR8E'
    DIAS_PARA_REMESSA = 120  # Data de hoje + 120 dias
    
    # Lista das abas
    ABAS_PARA_PROCESSAR = [
        '0-1500', '1501-5000', '5001-25000', 
        '25001-100000', '100001-200000', '>200000'
    ]

    # Mapeamento dos Grupos
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
        self.grupo_selecionado = None 
        self.data_remessa_calculada = None

    # --- UTILITÁRIOS ---
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

    # --- TRATAMENTO DE POPUPS (CORRIGIDO) ---
    def _lidar_com_popups(self, max_tentativas=4):
        """
        CORREÇÃO: Ao invés de buscar um botão específico (que causava erro se não existisse),
        enviamos o comando de tecla ENTER (VKey 0) direto para a janela de aviso (wnd[1]).
        Isso simula o usuário apertando Enter para limpar avisos amarelos.
        """
        time.sleep(0.5) # Breve pausa para o popup renderizar
        for _ in range(max_tentativas):
            try:
                # Verifica se existe uma janela secundária (modal)
                if self.session.findById("wnd[1]", False): 
                    # Envia ENTER na janela do popup
                    self.session.findById("wnd[1]").sendVKey(0)
                    time.sleep(0.5)
                else:
                    # Se não tem popup, sai do loop
                    break
            except:
                # Se der erro ao acessar wnd[1], assumimos que ele fechou/não existe
                break

    # --- TRANSAÇÃO ME51N ---
    def create_purchase_requisition_batch(self, batch_rows):
        try:
            # Reinicia transação
            self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/NME51N"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1) # Espera carregar
            
            # Identifica o Grid
            grid_id = "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell"
            
            # Verificação de segurança: Grid existe?
            if not self.session.findById(grid_id, False):
                return "Erro: Grid ME51N não carregou ou ID mudou."
            
            grid = self.session.findById(grid_id)
            
            # 1. Preenchimento
            linhas_preenchidas = 0
            for i, row in enumerate(batch_rows):
                try:
                    material = str(row.get('Material', '')).strip()
                    qtd = self.format_decimal(row.get('Qtd', ''))
                    preco = self.format_decimal(row.get('Preço', ''))

                    # Tenta focar na célula (opcional, ajuda na estabilidade)
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

            # 2. Processamento Inicial (Enter)
            self.session.findById("wnd[0]").sendVKey(0)
            self._lidar_com_popups() # Limpa avisos iniciais

            # 3. Verificar (Check)
            try:
                self.session.findById("wnd[0]/tbar[1]/btn[9]").press()
                self._lidar_com_popups() # Limpa avisos do check
            except: pass

            # 4. Checa se há erros impeditivos (Vermelhos)
            sbar = self.session.findById("wnd[0]/sbar")
            if sbar.MessageType == "E":
                # Se deu erro vermelho, aborta o lote
                return f"Erro SAP: {sbar.Text}"

            # 5. Salvar (Save)
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            
            # CRÍTICO: Loop agressivo de confirmação após salvar
            # É aqui que o "Data no passado" aparece e precisa ser confirmado com Enter
            self._lidar_com_popups(max_tentativas=5)

            # 6. Captura resultado final
            sbar_final = self.session.findById("wnd[0]/sbar")
            texto_final = sbar_final.Text
            
            if sbar_final.MessageType == "S" or "criad" in texto_final.lower() or "creat" in texto_final.lower():
                return texto_final
            else:
                return f"Status Final: {texto_final}"

        except Exception as e:
            return f"Erro Crítico Script: {str(e)}"

    # --- LOOP PRINCIPAL ---
    def run(self):
        if not self.connect_google(): return
        self.configurar_parametros_execucao()
        if not self.connect_sap(): return

        for nome_aba in Config.ABAS_PARA_PROCESSAR:
            print(f"\n>>> ACESSANDO ABA: {nome_aba}")
            try:
                worksheet = self.workbook.worksheet(nome_aba)
                data = worksheet.get_all_records()
                
                # Se a aba estiver vazia ou só cabeçalho
                if not data:
                    print(f"Aba {nome_aba} vazia.")
                    continue

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

                # Configuração de Lote
                if nome_aba in ['100001-200000', '>200000']:
                    BATCH_SIZE = 1
                    print(f" [!] Alto valor: Processando 1 por vez.")
                else:
                    BATCH_SIZE = 10
                    print(f" [i] Aba padrão: Lote de {BATCH_SIZE}.")

                for i in range(0, len(items_to_process), BATCH_SIZE):
                    chunk = items_to_process[i : i + BATCH_SIZE]
                    batch_data = [item['data'] for item in chunk]
                    
                    print(f" - Processando lote {i//BATCH_SIZE + 1}...")
                    
                    resultado = self.create_purchase_requisition_batch(batch_data)
                    print(f"   Resultado SAP: {resultado}")
                    
                    # Atualiza Planilha
                    for item in chunk:
                        try:
                            worksheet.update_cell(item['sheet_row'], col_status_idx, resultado)
                        except Exception as e:
                            print(f"   Erro update planilha: {e}")
                            time.sleep(2) # Espera API e tenta de novo
                            try: worksheet.update_cell(item['sheet_row'], col_status_idx, resultado)
                            except: pass
                        
            except Exception as e:
                print(f"Erro ao processar aba {nome_aba}: {e}")

        print("\nAutomação concluída em todas as abas.")

if __name__ == "__main__":
    app = SAPAutomation()
    app.run()