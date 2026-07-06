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
    NOME_ABA_DADOS = 'DANTAS'    
    
    # --- VARIÁVEIS PADRÃO ---
    CENTRO_PADRAO = 'BR8E'
    DIAS_PARA_REMESSA_FALLBACK = 0 # Usado caso a coluna LT esteja vazia
    
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
        '8': {'codigo': 'P08', 'desc': '787'},
        '9': {'codigo': 'C01', 'desc': 'Médio Prazo'},
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

    def calcular_data_remessa(self, lt_raw):
        """Calcula a Data de Remessa baseada no valor da coluna LT da planilha"""
        try:
            # Se vier vazio ou em branco, assume o valor de Fallback (ex: 0 dias)
            lt_val = int(float(str(lt_raw).strip())) if str(lt_raw).strip() else Config.DIAS_PARA_REMESSA_FALLBACK
        except ValueError:
            lt_val = Config.DIAS_PARA_REMESSA_FALLBACK
            
        return (datetime.now() + timedelta(days=lt_val)).strftime('%d.%m.%Y')

    def configurar_parametros_execucao(self):
        self.logger.info("%s", "\n" + "="*40)
        self.logger.info(" DATA REMESSA: Será calculada item a item (Coluna LT)")
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

    # --- ETAPA DE PRÉ-VERIFICAÇÃO (BASEADA NO SCRIPT DE BASE) ---
    def validar_chunk_sap(self, chunk):
        self.logger.info("Iniciando pré-verificação do lote no SAP...")
        resultados = []
        
        try:
            self.session.findById("wnd[0]").maximize()
            self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/NME51N"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(2)
            
            grid = self.session.findById(Config.GRID_ID_PADRAO)
            
            for i, row in enumerate(chunk):
                material = str(row.get('Material', '')).strip()
                pep_valor = str(row.get('PEP', '')).strip()
                qtd = self.format_decimal_sap(row.get('Qtd', ''))
                preco = self.format_decimal_sap(row.get('Preço', ''))
                data_remessa = self.calcular_data_remessa(row.get('LT', ''))
                
                status_item = "OK"
                try:
                    try: grid.modifyCell(i, "NAME1", Config.CENTRO_PADRAO)
                    except: pass 
                    
                    grid.modifyCell(i, "MATNR", material)
                    grid.modifyCell(i, "MENGE", qtd)
                    grid.modifyCell(i, "PREIS", preco)
                    grid.modifyCell(i, "EEIND", data_remessa)
                    grid.modifyCell(i, "EKGRP", self.grupo_selecionado)
                    grid.modifyCell(i, "WAERS", "USD")
                    
                    if pep_valor:
                        grid.modifyCell(i, "KNTTP", "P")
                    
                    # Pressiona Enter para validar a linha corrente
                    try:
                        grid.currentCellColumn = "WAERS"
                        grid.pressEnter()
                    except:
                        self.session.findById("wnd[0]").sendVKey(0)
                        
                    time.sleep(1)
                    
                    # Trata popups se surgirem
                    try:
                        if self.session.findById("wnd[1]", False):
                            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                            time.sleep(0.5)
                    except: pass
                    
                    # Captura erro na barra de status do SAP
                    sbar = self.session.findById("wnd[0]/sbar")
                    if sbar.MessageType in ('E', 'A') or "não está atualizado" in sbar.Text.lower():
                        status_item = sbar.Text
                        self.logger.warning(f" -> Item {i+1} ({material}) com erro no SAP: {status_item}")
                except Exception as e:
                    status_item = f"Erro na validação do item: {str(e)}"
                    
                resultados.append((row, status_item))
                
        except Exception as e:
            self.logger.error(f"Erro geral crítico na pré-verificação: {e}")
            for row in chunk:
                resultados.append((row, f"Erro crítico na validação: {str(e)}"))
        finally:
            # Sai da transação sem salvar (/N) para resetar a tela para o lote final
            try:
                self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/N"
                self.session.findById("wnd[0]").sendVKey(0)
                time.sleep(1)
            except: pass
            
        return resultados

    # --- TRANSAÇÃO ME51N CRIAÇÃO ---
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

            def _tentar_escrever_cabecalho():
                try:
                    self.session.findById(Config.ID_EDITOR_TEXTO).text = texto_final
                    try:
                        self.session.findById(Config.ID_EDITOR_TEXTO).setSelectionIndexes(92, 92)
                    except:
                        pass
                    return True
                except:
                    return False

            if not _tentar_escrever_cabecalho():
                self.logger.info("Cabeçalho recolhido. Expandindo com Ctrl+F2 (VKey 26)...")
                try:
                    self.session.findById("wnd[0]").sendVKey(26)
                    time.sleep(0.5)
                except Exception as ex:
                    self.logger.warning(f"Erro ao expandir cabeçalho: {ex}")

                if _tentar_escrever_cabecalho():
                    self.logger.info("Texto de cabeçalho preenchido (após expansão).")
                else:
                    self.logger.warning("Não foi possível preencher o texto de cabeçalho mesmo após expansão.")
            else:
                self.logger.info("Texto de cabeçalho preenchido.")

            # 3. PREENCHE O GRID (ITENS)
            grid = self.session.findById(Config.GRID_ID_PADRAO)
            itens_com_pep = []
            linhas_preenchidas = 0
            
            for i, row in enumerate(batch_rows):
                try:
                    material = str(row.get('Material', '')).strip()
                    pep_valor = str(row.get('PEP', '')).strip()
                    
                    valor_bruto = row.get('Preço', '')
                    self.logger.info(f" -> Item {i+1} Valor BRUTO (Texto): '{valor_bruto}'")
                    
                    qtd = self.format_decimal_sap(row.get('Qtd', ''))
                    preco = self.format_decimal_sap(valor_bruto)
                    data_remessa = self.calcular_data_remessa(row.get('LT', ''))
                    
                    self.logger.info(f" -> Enviando: Mat={material}, Qtd={qtd}, Preço={preco}, Remessa={data_remessa}, PEP={pep_valor}")
                    
                    try: grid.modifyCell(i, "NAME1", Config.CENTRO_PADRAO)
                    except: pass 
                    
                    grid.modifyCell(i, "MATNR", material)
                    grid.modifyCell(i, "MENGE", qtd)
                    grid.modifyCell(i, "PREIS", preco)
                    grid.modifyCell(i, "EEIND", data_remessa)
                    grid.modifyCell(i, "EKGRP", self.grupo_selecionado)
                    grid.modifyCell(i, "WAERS", "USD")
                    
                    if pep_valor:
                        grid.modifyCell(i, "KNTTP", "P")
                        itens_com_pep.append({'grid_index': i, 'pep': pep_valor, 'material': material})
                        self.logger.info(f"    -> PEP detectado: Classificação contábil = 'P' (Projeto)")
                    
                    linhas_preenchidas += 1
                except Exception as e:
                    self.logger.warning(f"Erro linha {i}: {e}")

            if linhas_preenchidas == 0:
                return "Erro: Nenhuma linha preenchida."

            # 4. VALIDA A PRIMEIRA INSERÇÃO E FECHA POPUPS
            try:
                grid.currentCellColumn = "WAERS"
                grid.pressEnter()
            except:
                self.session.findById("wnd[0]").sendVKey(0)
            
            time.sleep(1)
            try:
                if self.session.findById("wnd[1]", False):
                    self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except: pass

            # 5. TRAVA DE SEGURANÇA DAS DATAS
            self.logger.info("Forçando novamente a Data de Remessa (LT) contra padrão do SAP...")
            for i, row in enumerate(batch_rows):
                try:
                    data_remessa = self.calcular_data_remessa(row.get('LT', ''))
                    grid.modifyCell(i, "EEIND", data_remessa)
                except: pass
                
            try:
                grid.currentCellColumn = "EEIND"
                grid.pressEnter()
            except:
                self.session.findById("wnd[0]").sendVKey(0)

            time.sleep(1)
            try:
                if self.session.findById("wnd[1]", False):
                    self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except: pass

            # 5.1 PREENCHIMENTO DO ELEMENTO PEP
            if itens_com_pep:
                self.logger.info(f"Preenchendo Elemento PEP para {len(itens_com_pep)} item(ns)...")
                self._preencher_pep_itens(grid, itens_com_pep)

            # 6. GRAVAR
            self.logger.info("Gravando...")
            try:
                self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            except Exception as e:
                self.logger.error(f"Erro ao pressionar Gravar: {e}")

            time.sleep(1)
            try:
                popup = self.session.findById("wnd[1]", False)
                if popup:
                    try:
                        self.session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
                        self.logger.info("Popup 'Gravar doc.' confirmado com 'Gravar'.")
                    except:
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except: pass

            # 7. CAPTURA MENSAGEM FINAL
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

    def _preencher_pep_itens(self, grid, itens_com_pep):
        ID_PEP = (
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200"
            "/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT7"
            "/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100"
            "/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-PS_POSID"
        )

        for item_pep in itens_com_pep:
            idx = item_pep['grid_index']
            pep = item_pep['pep']
            material = item_pep['material']

            try:
                self.logger.info(f"  -> Preenchendo PEP '{pep}' para item {idx+1} (Mat: {material})")
                grid.setCurrentCell(idx, "MATNR")
                grid.selectedRows = str(idx)
                time.sleep(0.5)
                self.session.findById("wnd[0]").sendVKey(0)
                time.sleep(1)

                try:
                    if self.session.findById("wnd[1]", False):
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        time.sleep(0.5)
                except: pass

                ID_ABA_CLASSCONT = "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT7"
                try:
                    self.session.findById(ID_ABA_CLASSCONT).select()
                    time.sleep(0.5)
                except Exception as e:
                    self.logger.warning(f"    -> Não foi possível selecionar a aba ClassCont.: {e}")

                pep_preenchido = False
                try:
                    campo = self.session.findById(ID_PEP)
                    campo.text = pep
                    campo.caretPosition = len(pep)
                    pep_preenchido = True
                except Exception as e:
                    self.logger.warning(f"    -> Falha no campo primário: {e}")

                if not pep_preenchido:
                    ID_PEP_ALT = ID_PEP.replace("/subSUB3:SAPLMEVIEWS:1100", "/subSUB2:SAPLMEVIEWS:1100")
                    try:
                        campo = self.session.findById(ID_PEP_ALT)
                        campo.text = pep
                        pep_preenchido = True
                    except Exception as e:
                        self.logger.warning(f"    -> Fallback subSUB2 também falhou: {e}")

                if not pep_preenchido:
                    self.logger.warning(f"    -> FALHA: Campo PEP não encontrado para item {idx+1}.")

                self.session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.5)
                try:
                    if self.session.findById("wnd[1]", False):
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                except: pass

            except Exception as e:
                self.logger.warning(f"  -> Erro ao preencher PEP para item {idx+1}: {e}")
                continue

        self.logger.info("Preenchimento de PEP concluído.")

    def run(self):
        if not self.connect_google(): return
        self.configurar_parametros_execucao()
        if not self.connect_sap(): return

        self.logger.info("\n>>> LENDO DADOS DA ABA: %s", Config.NOME_ABA_DADOS)
        try:
            self.worksheet = self.workbook.worksheet(Config.NOME_ABA_DADOS)
            raw_data = self.worksheet.get_all_values()
            
            if not raw_data or len(raw_data) < 2:
                self.logger.info("Planilha vazia ou sem dados.")
                return

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
                self.logger.info("\n - Analisando lote de %s item(ns)...", len(chunk))

                # -------------------------------------------------------------
                # NOVA ETAPA: EXECUTA A PRÉ-VERIFICAÇÃO DE CADA ITEM DO LOTE
                # -------------------------------------------------------------
                resultados_validacao = self.validar_chunk_sap(chunk)
                
                # Separa apenas os itens que retornaram "OK"
                chunk_ok = []
                for item, status_validacao in resultados_validacao:
                    if status_validacao != "OK":
                        # Sinaliza o erro retornado pelo SAP direto na planilha
                        self._atualizar_status_planilha(item['sheet_row_index'], col_status_idx, status_validacao)
                    else:
                        chunk_ok.append(item)

                # Se o lote ficou vazio após desconsiderar os erros, pula para o próximo
                if not chunk_ok:
                    self.logger.info(" -> Lote desconsiderado por completo (Nenhum item válido).")
                    continue
                
                self.logger.info(" -> Prosseguindo com %s item(ns) válidos no lote.", len(chunk_ok))
                # -------------------------------------------------------------

                # Itens com PEP são sempre processados 1 a 1
                tem_pep = any(str(it.get('PEP', '')).strip() for it in chunk_ok)
                if tem_pep and len(chunk_ok) > 1:
                    self.logger.info(" - Lote contém PEP → processando item(ns) individualmente...")
                    for sub_item in chunk_ok:
                        res_indiv = self.create_purchase_requisition_batch([sub_item])
                        self._atualizar_status_planilha(sub_item['sheet_row_index'], col_status_idx, res_indiv)
                    continue

                # Cria a Requisição apenas com os itens válidos (chunk_ok)
                resultado = self.create_purchase_requisition_batch(chunk_ok)

                eh_numero = resultado.isdigit()
                sucesso = eh_numero or any(x in resultado.lower() for x in ['criad', 'creat', 'gravad'])

                if not sucesso and len(chunk_ok) > 1:
                    for sub_item in chunk_ok:
                        res_indiv = self.create_purchase_requisition_batch([sub_item])
                        self._atualizar_status_planilha(sub_item['sheet_row_index'], col_status_idx, res_indiv)
                else:
                    for item in chunk_ok:
                        self._atualizar_status_planilha(item['sheet_row_index'], col_status_idx, resultado)

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