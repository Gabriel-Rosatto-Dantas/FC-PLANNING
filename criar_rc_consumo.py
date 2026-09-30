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
import unicodedata

# Versao 3: PEP dinamico e recuperacao do grid apos mudanca de dynpro da ME51N.
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
        self._grid_id_itens_cache = None

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

    def _pontuar_grid_itens(self, grid):
        """Retorna (pontuacao, colunas) para identificar o grid principal da ME51N."""
        tipo = str(self._propriedade_sap(grid, "Type", ""))
        subtipo = str(self._propriedade_sap(grid, "SubType", ""))
        grid_id = str(self._propriedade_sap(grid, "Id", ""))
        tipo_upper = f"{tipo} {subtipo}".upper()
        id_upper = grid_id.upper()

        # Rejeita controles comuns antes de acessar RowCount. Isso evita milhares
        # de excecoes COM durante a busca recursiva pela arvore do SAP GUI.
        if "GRID" not in tipo_upper and "GRIDCONTROL" not in id_upper:
            return -1, set()

        try:
            # Acessar RowCount tambem testa se a referencia COM ainda e valida.
            int(grid.RowCount)
        except Exception:
            return -1, set()

        colunas = set()
        try:
            ordem = self._colecao_sap_para_lista(grid.ColumnOrder)
            colunas.update(
                str(coluna).strip().upper()
                for coluna in ordem
                if str(coluna).strip()
            )
        except Exception:
            pass

        # Alguns wrappers COM nao entregam ColumnOrder. Testa as colunas tecnicas
        # conhecidas sem modificar nenhum valor do SAP.
        colunas_esperadas = (
            "MATNR", "MENGE", "PREIS", "EEIND", "EKGRP", "WAERS",
            "KNTTP", "NAME1", "TXZ01", "WERKS",
        )
        for coluna in colunas_esperadas:
            if coluna in colunas:
                continue
            try:
                grid.GetColumnTitles(coluna)
                colunas.add(coluna)
                continue
            except Exception:
                pass
            try:
                if int(grid.RowCount) > 0:
                    grid.GetCellValue(0, coluna)
                    colunas.add(coluna)
            except Exception:
                pass

        pontuacao = 0
        if "GUIGRIDVIEW" in tipo_upper or "GRIDVIEW" in tipo_upper:
            pontuacao += 80
        if "GRIDCONTROL" in id_upper:
            pontuacao += 120
        if "SAPLMEGUI:32" in id_upper:
            pontuacao += 80

        pesos = {
            "MATNR": 180,
            "MENGE": 100,
            "PREIS": 70,
            "EEIND": 70,
            "EKGRP": 50,
            "WAERS": 50,
            "KNTTP": 50,
            "NAME1": 20,
        }
        for coluna, peso in pesos.items():
            if coluna in colunas:
                pontuacao += peso

        # O grid de classificacao contabil pode ser um GuiGridView, mas normalmente
        # nao possui MATNR. Exigir MATNR evita selecionar o grid errado.
        if "MATNR" not in colunas and "GRIDCONTROL" not in id_upper:
            return -1, colunas

        return pontuacao, colunas

    def _obter_grid_itens(self, grid_anterior=None, aguardar_segundos=2.0):
        """
        Localiza o grid principal de itens sem depender do dynpro fixo 0013.

        Ao abrir o detalhe de classificacao contabil, a ME51N pode trocar o
        caminho de SAPLMEGUI:0013 para SAPLMEGUI:0019. A referencia COM antiga
        costuma continuar valida; por isso ela e tentada primeiro.
        """
        inicio = time.time()
        ultimo_erro = None

        while True:
            candidatos_objeto = []
            if grid_anterior is not None:
                candidatos_objeto.append(("referencia anterior", grid_anterior))

            ids = []
            if self._grid_id_itens_cache:
                ids.append(self._grid_id_itens_cache)
            ids.append(Config.GRID_ID_PADRAO)

            base = Config.GRID_ID_PADRAO
            for dynpro in ("0013", "0019", "0018", "0014", "0015", "0020"):
                candidato = base.replace(
                    "SAPLMEGUI:0013",
                    f"SAPLMEGUI:{dynpro}",
                    1,
                )
                ids.append(candidato)
                ids.append(
                    candidato.replace(
                        "/subSUB2:SAPLMEVIEWS:1100/",
                        "/subSUB3:SAPLMEVIEWS:1100/",
                        1,
                    )
                )
                for tela_grid in ("3212", "3211", "3210", "3213", "3214"):
                    ids.append(
                        candidato.replace(
                            "SAPLMEGUI:3212",
                            f"SAPLMEGUI:{tela_grid}",
                            1,
                        )
                    )

            vistos_ids = set()
            for grid_id in ids:
                if not grid_id or grid_id in vistos_ids:
                    continue
                vistos_ids.add(grid_id)
                try:
                    controle = self.session.findById(grid_id)
                    candidatos_objeto.append((f"id {grid_id}", controle))
                except Exception as exc:
                    ultimo_erro = exc

            melhor = None
            melhor_pontuacao = -1
            melhor_colunas = set()
            melhor_origem = ""

            for origem, controle in candidatos_objeto:
                pontuacao, colunas = self._pontuar_grid_itens(controle)
                if pontuacao > melhor_pontuacao:
                    melhor = controle
                    melhor_pontuacao = pontuacao
                    melhor_colunas = colunas
                    melhor_origem = origem

            # Se a referencia anterior ou um ID conhecido ja identificou o grid,
            # nao percorre toda a arvore do SAP. A busca recursiva fica como fallback.
            if melhor_pontuacao < 300:
                try:
                    raiz = self.session.findById("wnd[0]/usr")
                    for controle in self._iterar_controles_sap(raiz, limite=6000):
                        pontuacao, colunas = self._pontuar_grid_itens(controle)
                        if pontuacao > melhor_pontuacao:
                            melhor = controle
                            melhor_pontuacao = pontuacao
                            melhor_colunas = colunas
                            melhor_origem = "busca dinamica"
                except Exception as exc:
                    ultimo_erro = exc

            if melhor is not None and melhor_pontuacao >= 180:
                grid_id = str(self._propriedade_sap(melhor, "Id", ""))
                if grid_id:
                    id_anterior = self._grid_id_itens_cache
                    self._grid_id_itens_cache = grid_id
                    if grid_id != Config.GRID_ID_PADRAO and grid_id != id_anterior:
                        self.logger.info(
                            "Grid de itens recuperado apos mudanca de layout: %s",
                            grid_id,
                        )
                self.logger.debug(
                    "Grid de itens selecionado (%s; score=%s; colunas=%s).",
                    melhor_origem,
                    melhor_pontuacao,
                    ",".join(sorted(melhor_colunas)),
                )
                return melhor

            if time.time() - inicio >= aguardar_segundos:
                detalhe = f" Ultimo erro: {ultimo_erro}" if ultimo_erro else ""
                raise RuntimeError(
                    "Grid principal de itens da ME51N nao encontrado no layout atual."
                    + detalhe
                )
            time.sleep(0.25)

    # --- ETAPA DE PRÉ-VERIFICAÇÃO (CORRIGIDA COM LEITURA AVANÇADA DE ERROS) ---
    def validar_chunk_sap(self, chunk):
        self.logger.info("Iniciando pré-verificação individual dos itens no SAP...")
        resultados = []

        for i, row in enumerate(chunk):
            material = str(row.get('Material', '')).strip()
            pep_valor = str(row.get('PEP', '')).strip()
            qtd = self.format_decimal_sap(row.get('Qtd', ''))
            preco = self.format_decimal_sap(row.get('Preço', ''))
            data_remessa = self.calcular_data_remessa(row.get('LT', ''))

            status_item = "OK"
            try:
                # 1. Abre uma transação limpa PARA CADA ITEM. 
                self.session.findById("wnd[0]").maximize()
                self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/NME51N"
                self.session.findById("wnd[0]").sendVKey(0)
                time.sleep(1.5)

                grid = self._obter_grid_itens()

                try: grid.modifyCell(0, "NAME1", Config.CENTRO_PADRAO)
                except: pass

                grid.modifyCell(0, "MATNR", material)
                grid.modifyCell(0, "MENGE", qtd)
                grid.modifyCell(0, "PREIS", preco)
                grid.modifyCell(0, "EEIND", data_remessa)
                grid.modifyCell(0, "EKGRP", self.grupo_selecionado)
                grid.modifyCell(0, "WAERS", "USD")

                if pep_valor:
                    grid.modifyCell(0, "KNTTP", "P")

                # Primeiro Enter para o SAP processar o item
                try:
                    grid.setCurrentCell(0, "WAERS")
                    grid.pressEnter()
                except Exception:
                    self.session.findById("wnd[0]").sendVKey(0)

                time.sleep(1)
                
                if not pep_valor:
                    self._fechar_popup_sap()

                # Se tem PEP, tenta preencher
                if pep_valor:
                    retorno_pep = self._preencher_pep_itens(
                        grid,
                        [{'grid_index': 0, 'pep': pep_valor, 'material': material}],
                        contexto="pré-verificação"
                    )
                    pep_ok, mensagem_pep = retorno_pep.get(0, (False, "Falha ao preencher o PEP."))
                    if not pep_ok:
                        status_item = mensagem_pep

                # 2. Leitura inteligente da Barra de Status (Ignora Warnings e foca nos Erros)
                if status_item == "OK":
                    try:
                        # Tenta dar até 3 Enters para pular mensagens de Aviso (W) ou Informação (I)
                        for tentativa in range(3):
                            sbar = self.session.findById("wnd[0]/sbar")
                            tipo_msg = str(sbar.MessageType).strip()
                            texto_status = str(sbar.Text).strip()
                            texto_lower = texto_status.lower()
                            
                            # Palavras-chave expandidas com base no seu print ("Bloq.", "Inativo", "Não permite")
                            palavras_erro = ['bloq', 'inativ', 'não permite', 'não está atualizado', 'erro', 'obrigatório']
                            
                            # Se for Erro duro (E/A) ou contiver palavras de bloqueio, reprova na hora
                            if tipo_msg in ('E', 'A') or any(p in texto_lower for p in palavras_erro):
                                status_item = texto_status or f"Item bloqueado pelo SAP (Tipo: {tipo_msg})."
                                break
                            
                            # Se for apenas um Aviso (W) ou Informação (I), dá Enter para ver se o SAP libera o item
                            if tipo_msg in ('W', 'I'):
                                self.session.findById("wnd[0]").sendVKey(0)
                                time.sleep(0.8)
                            else:
                                # Se não tem mensagem nenhuma, sai do loop (item está limpo)
                                break
                    except Exception as e:
                        self.logger.warning(f"Erro ao avaliar barra de status: {e}")

                if status_item != "OK":
                    self.logger.warning(f" -> Item {i+1} ({material}) REPROVADO: {status_item}")
                else:
                    self.logger.info(f" -> Item {i+1} ({material}) APROVADO.")

            except Exception as e:
                status_item = f"Erro na validação do item: {str(e)}"
                self.logger.warning(f" -> Item {i+1} ({material}) com falha técnica: {e}")

            resultados.append((row, status_item))

        # Zera a tela ao final
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/N"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
        except Exception: pass

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
            grid = self._obter_grid_itens()
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
            except Exception:
                self.session.findById("wnd[0]").sendVKey(0)

            time.sleep(1)
            # Não feche uma eventual janela de classificação contábil quando
            # existir item com PEP. Ela pode conter justamente o campo WBS/PEP.
            if not itens_com_pep:
                self._fechar_popup_sap()

            # 5. PREENCHIMENTO DO ELEMENTO PEP
            # O preenchimento agora ocorre logo depois que KNTTP='P' foi validado,
            # antes de uma nova validação das datas e antes da gravação.
            if itens_com_pep:
                self.logger.info(
                    "Preenchendo Elemento PEP para %s item(ns)...",
                    len(itens_com_pep),
                )
                retorno_pep = self._preencher_pep_itens(
                    grid,
                    itens_com_pep,
                    contexto="criação",
                )

                falhas_pep = []
                for item_pep in itens_com_pep:
                    idx_pep = item_pep['grid_index']
                    ok_pep, msg_pep = retorno_pep.get(
                        idx_pep,
                        (False, "Falha desconhecida ao preencher o Elemento PEP."),
                    )
                    if not ok_pep:
                        falhas_pep.append(
                            f"item {idx_pep + 1} ({item_pep['material']}): {msg_pep}"
                        )

                if falhas_pep:
                    mensagem = "Erro no preenchimento do PEP: " + "; ".join(falhas_pep)
                    self.logger.error(mensagem)
                    return mensagem

                # O helper altera o dynpro da ME51N (por exemplo, 0013 -> 0019).
                # Primeiro reutiliza a referencia COM anterior; se necessario, localiza
                # o grid dinamicamente. A falta do grid nao deve cancelar uma RC cujo
                # PEP ja foi confirmado: nesse caso apenas pula a trava opcional da data.
                try:
                    grid = self._obter_grid_itens(grid_anterior=grid)
                except Exception as e:
                    self.logger.warning(
                        "PEP confirmado, mas o grid de itens nao ficou acessivel: %s. "
                        "A trava adicional da data sera ignorada e a gravacao continuara.",
                        e,
                    )
                    grid = None

            # 5.1 TRAVA DE SEGURANCA DAS DATAS
            if grid is not None:
                self.logger.info(
                    "Forcando novamente a Data de Remessa (LT) contra padrao do SAP..."
                )
                for i, row in enumerate(batch_rows):
                    try:
                        data_remessa = self.calcular_data_remessa(row.get('LT', ''))
                        grid.modifyCell(i, "EEIND", data_remessa)
                    except Exception as e:
                        self.logger.warning(
                            "Nao foi possivel reforcar a data do item %s: %s",
                            i + 1,
                            e,
                        )

                try:
                    grid.currentCellColumn = "EEIND"
                    grid.pressEnter()
                except Exception:
                    self.session.findById("wnd[0]").sendVKey(0)

                time.sleep(1)
                self._fechar_popup_sap()
            else:
                self.logger.info(
                    "Trava adicional da Data de Remessa ignorada; seguindo para gravacao."
                )

            # Não tenta gravar enquanto o SAP ainda acusa erro de item/PEP.
            try:
                sbar_pre_gravacao = self.session.findById("wnd[0]/sbar")
                texto_pre_gravacao = str(sbar_pre_gravacao.Text).strip()
                if sbar_pre_gravacao.MessageType in ('E', 'A'):
                    mensagem = f"Erro antes de gravar: {texto_pre_gravacao or 'erro retornado pelo SAP'}"
                    self.logger.warning(mensagem)
                    return mensagem
            except Exception as e:
                self.logger.warning("Não foi possível ler a barra de status antes de gravar: %s", e)

            # 6. GRAVAR
            self.logger.info("Gravando a requisição e lidando com pop-ups de alerta...")
            
            # Limpa qualquer pop-up que tenha ficado para trás antes de começar
            self._fechar_popup_sap()
            
            sucesso_gravacao = False
            for tentativa in range(6): # Tenta o fluxo de gravação até 6 vezes
                try:
                    # 1. Aperta o botão Gravar na tela principal
                    btn_gravar = self.session.findById("wnd[0]/tbar[0]/btn[11]", False)
                    if btn_gravar:
                        btn_gravar.press()
                        time.sleep(1.0)
                    
                    # 2. Loop agressivo para tratar os pop-ups "Gravar doc." que empilham
                    for _ in range(5):
                        popup = self.session.findById("wnd[1]", False)
                        if not popup: 
                            break # Sem pop-up na tela, segue o baile!
                        
                        try:
                            # Tenta clicar explicitamente no botão 1 (Gravar / Sim)
                            self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                            time.sleep(0.5)
                        except:
                            try:
                                # Variação do botão de confirmar
                                self.session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
                                time.sleep(0.5)
                            except:
                                try:
                                    # FALLBACK MESTRE: Como o botão 'Gravar' já vem com o foco 
                                    # (pontilhado), dar Enter na janela do pop-up fará o clique nele.
                                    popup.sendVKey(0)
                                    time.sleep(0.5)
                                except:
                                    pass
                    
                    # 3. Trata os alertas amarelos (W) na barra de status que travam a gravação
                    sbar = self.session.findById("wnd[0]/sbar", False)
                    if sbar:
                        tipo_msg = str(getattr(sbar, 'MessageType', '')).strip()
                        texto_msg = str(getattr(sbar, 'Text', '')).strip().lower()
                        
                        # Se já for sucesso, encerra o loop de tentativas imediatamente
                        if tipo_msg == 'S' or any(x in texto_msg for x in ['criad', 'creat', 'gravad']):
                            sucesso_gravacao = True
                            break
                            
                        # Se for Aviso/Warning (W) ou Info (I), dá Enter para pular a mensagem
                        if tipo_msg in ('W', 'I'):
                            self.session.findById("wnd[0]").sendVKey(0) 
                            time.sleep(0.8)
                            continue # Volta ao topo do loop e clica no botão Gravar de novo
                            
                    sucesso_gravacao = True
                    break
                    
                except Exception as e:
                    self.logger.warning(f"Tentativa {tentativa+1} de gravar interceptada: {e}")
                    time.sleep(1)


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

    def _fechar_popup_sap(self, max_tentativas=10):
        """Fecha pop-ups de aviso/confirmação em loop, lidando com múltiplos pop-ups seguidos."""
        fechou_algum = False
        for _ in range(max_tentativas):
            try:
                # Verifica se a janela wnd[1] (popup) existe
                popup = self.session.findById("wnd[1]", False)
                if not popup:
                    break # Se não tem mais popup aberto, sai do loop imediatamente

                # Lista de botões comuns de confirmação ("Continuar", "Sim", "Gravar")
                botoes = (
                    "wnd[1]/tbar[0]/btn[0]",           # Check verde (Continuar / Confirmar)
                    "wnd[1]/usr/btnSPOP-OPTION1",      # Sim
                    "wnd[1]/usr/btnSPOP-VAROPTION1",   # Gravar doc. / Confirmar variação
                )
                
                clicou_botao = False
                for botao_id in botoes:
                    try:
                        btn = self.session.findById(botao_id, False)
                        if btn:
                            btn.press()
                            clicou_botao = True
                            fechou_algum = True
                            time.sleep(0.5)
                            break # Achou um botão, sai da busca e vai para o próximo ciclo do popup
                    except:
                        pass
                
                # Se não conseguiu clicar em nenhum botão padrão, tenta forçar um "Enter"
                if not clicou_botao:
                    try:
                        popup.sendVKey(0)
                        fechou_algum = True
                        time.sleep(0.5)
                    except:
                        pass

            except Exception:
                break
                
        return fechou_algum

    @staticmethod
    def _filhos_controle_sap(controle):
        """Retorna os filhos de um controle SAP GUI de forma tolerante ao COM."""
        try:
            filhos = controle.Children
            quantidade = int(filhos.Count)
        except Exception:
            return []

        resultado = []
        for posicao in range(quantidade):
            filho = None
            try:
                filho = filhos(posicao)
            except Exception:
                try:
                    filho = filhos.Item(posicao)
                except Exception:
                    pass
            if filho is not None:
                resultado.append(filho)
        return resultado

    def _localizar_controle_sap(self, nomes=(), sufixos_id=(), limite=2500):
        """Procura um controle visível pelo Name técnico ou pelo final do Id."""
        try:
            raiz = self.session.findById("wnd[0]/usr")
        except Exception:
            raiz = self.session.findById("wnd[0]")

        pilha = [raiz]
        examinados = 0
        nomes = set(nomes)
        sufixos_id = tuple(sufixos_id)

        while pilha and examinados < limite:
            controle = pilha.pop()
            examinados += 1

            try:
                nome = str(controle.Name)
            except Exception:
                nome = ""

            try:
                controle_id = str(controle.Id)
            except Exception:
                controle_id = ""

            if nome in nomes or any(controle_id.endswith(sufixo) for sufixo in sufixos_id):
                return controle

            pilha.extend(self._filhos_controle_sap(controle))

        return None

    def _selecionar_aba_classificacao_contabil(self):
        """
        Seleciona a aba de Classificação/Atribuição contábil.

        O número TABREQDT muda conforme versão, personalização e estado da
        ME51N. Por isso a busca usa, nesta ordem: texto da aba, nomes técnicos
        conhecidos (TABREQDT16/TABREQDT7), IDs completos e busca recursiva.
        """
        try:
            area_usuario = self.session.findById("wnd[0]/usr")
        except Exception:
            area_usuario = None

        abas_encontradas = {}

        # 1) Método mais estável: encontra a aba pelo texto apresentado ao usuário.
        if area_usuario is not None:
            for numero in range(1, 31):
                nome_aba = f"TABREQDT{numero}"
                try:
                    aba = area_usuario.findByName(nome_aba, "GuiTab")
                except Exception:
                    continue

                abas_encontradas[nome_aba] = aba
                textos = []
                for propriedade in ("Text", "Tooltip", "AccText"):
                    try:
                        textos.append(str(getattr(aba, propriedade)))
                    except Exception:
                        pass

                descricao = unicodedata.normalize("NFKD", " ".join(textos))
                descricao = "".join(
                    caractere
                    for caractere in descricao
                    if not unicodedata.combining(caractere)
                ).lower()

                eh_aba_contabil = (
                    ("contabil" in descricao and ("class" in descricao or "atrib" in descricao))
                    or "account assignment" in descricao
                    or "imputacion" in descricao
                )
                if eh_aba_contabil:
                    try:
                        aba.select()
                        self.logger.info(
                            "    -> Aba contábil selecionada pelo texto: %s",
                            nome_aba,
                        )
                        time.sleep(0.5)
                        return True
                    except Exception:
                        pass

        # 2) Fallbacks observados em layouts diferentes da ME51N.
        for nome_aba in ("TABREQDT16", "TABREQDT7"):
            aba = abas_encontradas.get(nome_aba)
            if aba is None and area_usuario is not None:
                try:
                    aba = area_usuario.findByName(nome_aba, "GuiTab")
                except Exception:
                    aba = None

            if aba is not None:
                try:
                    aba.select()
                    self.logger.info(
                        "    -> Aba contábil selecionada pelo nome técnico: %s",
                        nome_aba,
                    )
                    time.sleep(0.5)
                    return True
                except Exception:
                    pass

        # 3) IDs completos para os números de subscreen mais comuns.
        for tela in ("0019", "0013", "0018", "0014"):
            for subview in ("subSUB3", "subSUB2"):
                base = (
                    f"wnd[0]/usr/subSUB0:SAPLMEGUI:{tela}/{subview}:SAPLMEVIEWS:1100/"
                    "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/"
                    "subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL"
                )
                for nome_aba in ("TABREQDT16", "TABREQDT7"):
                    aba_id = f"{base}/tabp{nome_aba}"
                    try:
                        aba = self.session.findById(aba_id)
                        aba.select()
                        self.logger.debug(
                            "Aba ClassCont. localizada por ID: %s",
                            aba_id,
                        )
                        time.sleep(0.5)
                        return True
                    except Exception:
                        continue

        # 4) Último fallback: percorre a árvore de controles visíveis.
        aba = self._localizar_controle_sap(
            nomes=("TABREQDT16", "TABREQDT7"),
            sufixos_id=("/tabpTABREQDT16", "/tabpTABREQDT7"),
        )
        if aba is not None:
            try:
                aba.select()
                self.logger.debug(
                    "Aba ClassCont. localizada dinamicamente: %s",
                    aba.Id,
                )
                time.sleep(0.5)
                return True
            except Exception as e:
                self.logger.warning(
                    "Aba ClassCont. encontrada, mas não pôde ser selecionada: %s",
                    e,
                )

        return False

    @staticmethod
    def _normalizar_texto_sap(valor):
        """Normaliza textos de rótulos/abas para comparação independente de idioma."""
        texto = unicodedata.normalize("NFKD", str(valor or ""))
        texto = "".join(
            caractere for caractere in texto
            if not unicodedata.combining(caractere)
        )
        return " ".join(texto.lower().strip().split())

    @classmethod
    def _texto_indica_pep(cls, valor):
        """Retorna True quando o texto descreve Elemento PEP/WBS/EAP."""
        texto = cls._normalizar_texto_sap(valor)
        if not texto:
            return False

        expressoes = (
            "elemento pep",
            "elem. pep",
            "elemento eap",
            "elemento psp",
            "wbs element",
            "work breakdown structure",
            "estrutura analitica",
            "elemento de proyecto",
            "elemento del proyecto",
        )
        if any(expressao in texto for expressao in expressoes):
            return True

        return bool(
            re.search(r"\bpep\b", texto)
            or re.search(r"\bwbs\b", texto)
        )

    @staticmethod
    def _propriedade_sap(controle, nome, padrao=""):
        try:
            valor = getattr(controle, nome)
            return padrao if valor is None else valor
        except Exception:
            return padrao

    @staticmethod
    def _colecao_sap_para_lista(colecao):
        """Converte coleções/arrays COM em lista sem depender de um único wrapper."""
        if colecao is None:
            return []
        if isinstance(colecao, (str, bytes)):
            return [str(colecao)]
        if isinstance(colecao, (list, tuple)):
            return list(colecao)

        itens = []
        try:
            quantidade = int(colecao.Count)
        except Exception:
            quantidade = None

        if quantidade is not None:
            for indice in range(quantidade):
                item = None
                for modo in ("call", "item", "element"):
                    try:
                        if modo == "call":
                            item = colecao(indice)
                        elif modo == "item":
                            item = colecao.Item(indice)
                        else:
                            item = colecao.ElementAt(indice)
                        break
                    except Exception:
                        continue
                if item is not None:
                    itens.append(item)
            return itens

        try:
            return list(colecao)
        except Exception:
            return itens

    def _iterar_controles_sap(self, raiz, limite=8000):
        """Percorre a árvore visível do SAP GUI a partir de uma raiz."""
        pilha = [raiz]
        vistos = set()
        examinados = 0

        while pilha and examinados < limite:
            controle = pilha.pop()
            examinados += 1

            controle_id = str(self._propriedade_sap(controle, "Id", ""))
            controle_tipo = str(self._propriedade_sap(controle, "Type", ""))
            chave = controle_id or f"{controle_tipo}:{id(controle)}"
            if chave in vistos:
                continue
            vistos.add(chave)

            yield controle

            filhos = self._filhos_controle_sap(controle)
            if filhos:
                pilha.extend(reversed(filhos))

    def _janelas_sap_abertas(self):
        """Retorna wnd[0], wnd[1]... atualmente disponíveis na sessão."""
        janelas = []
        for indice in range(5):
            try:
                janela = self.session.findById(f"wnd[{indice}]", False)
            except Exception:
                try:
                    janela = self.session.findById(f"wnd[{indice}]")
                except Exception:
                    janela = None
            if janela is not None:
                janelas.append(janela)
        return janelas

    def _pontuar_campo_pep(self, campo):
        """Pontua um campo editável conforme nome técnico, ID e rótulo associado."""
        tipo = str(self._propriedade_sap(campo, "Type", ""))
        if tipo not in ("GuiCTextField", "GuiTextField"):
            return 0

        try:
            if int(self._propriedade_sap(campo, "Changeable", 1)) == 0:
                return 0
        except Exception:
            pass

        nome = str(self._propriedade_sap(campo, "Name", ""))
        campo_id = str(self._propriedade_sap(campo, "Id", ""))
        nome_id = f"{nome} {campo_id}".upper()

        nomes_exatos = {
            "COBL-PS_POSID",
            "MEACCT1100-PS_PSP_PNR",
            "MEACCT1100-PS_POSID",
            "EBKN-PS_PSP_PNR",
            "EKKN-PS_PSP_PNR",
            "RM06B-PS_PSP_PNR",
            "PS_PSP_PNR",
            "PS_POSID",
        }

        pontuacao = 0
        if nome.upper() in nomes_exatos:
            pontuacao = max(pontuacao, 1000)
        if "PS_POSID" in nome_id:
            pontuacao = max(pontuacao, 980)
        if "PS_PSP_PNR" in nome_id:
            pontuacao = max(pontuacao, 970)
        if "PSPNR" in nome_id:
            pontuacao = max(pontuacao, 900)
        if "POSID" in nome_id and (
            "SAPLMEACCTVI" in nome_id
            or "SAPLKACB" in nome_id
            or "ACCOUNT" in nome_id
        ):
            pontuacao = max(pontuacao, 880)

        textos_associados = []
        for propriedade in (
            "Tooltip",
            "DefaultTooltip",
            "AccText",
            "AccTooltip",
        ):
            valor = self._propriedade_sap(campo, propriedade, "")
            if valor:
                textos_associados.append(str(valor))

        for propriedade_rotulo in ("LeftLabel", "RightLabel"):
            rotulo = self._propriedade_sap(campo, propriedade_rotulo, None)
            if rotulo is not None:
                for propriedade in ("Text", "Tooltip", "AccText"):
                    valor = self._propriedade_sap(rotulo, propriedade, "")
                    if valor:
                        textos_associados.append(str(valor))

        try:
            rotulos_acessiveis = self._colecao_sap_para_lista(campo.AccLabelCollection)
        except Exception:
            rotulos_acessiveis = []
        for rotulo in rotulos_acessiveis:
            for propriedade in ("Text", "Tooltip", "AccText"):
                valor = self._propriedade_sap(rotulo, propriedade, "")
                if valor:
                    textos_associados.append(str(valor))

        if self._texto_indica_pep(" ".join(textos_associados)):
            pontuacao = max(pontuacao, 960)

        return pontuacao

    def _localizar_campo_pep(self):
        """
        Localiza o campo PEP em qualquer janela aberta.

        Além de COBL-PS_POSID, contempla o campo usado em outros layouts da
        ME51N, como MEACCT1100-PS_PSP_PNR, e também identifica o campo pelo
        rótulo visível (Elemento PEP/WBS/EAP).
        """
        nomes = (
            "COBL-PS_POSID",
            "MEACCT1100-PS_PSP_PNR",
            "MEACCT1100-PS_POSID",
            "EBKN-PS_PSP_PNR",
            "EKKN-PS_PSP_PNR",
            "RM06B-PS_PSP_PNR",
            "PS_PSP_PNR",
            "PS_POSID",
        )
        tipos = ("GuiCTextField", "GuiTextField")

        # Janelas modais primeiro; em alguns sistemas a conta do item abre em wnd[1].
        janelas = list(reversed(self._janelas_sap_abertas()))
        candidatos = []
        ids_vistos = set()

        for janela in janelas:
            # Busca exata e recursiva fornecida pelo próprio SAP GUI.
            for nome in nomes:
                for tipo in tipos:
                    try:
                        campo = janela.findByName(nome, tipo)
                    except Exception:
                        continue
                    campo_id = str(self._propriedade_sap(campo, "Id", ""))
                    if campo_id not in ids_vistos:
                        ids_vistos.add(campo_id)
                        candidatos.append((2000, campo))

            # Busca tolerante por nome/ID/rótulo associado.
            controles_janela = list(self._iterar_controles_sap(janela))
            for controle in controles_janela:
                pontuacao = self._pontuar_campo_pep(controle)
                if pontuacao <= 0:
                    continue
                controle_id = str(self._propriedade_sap(controle, "Id", ""))
                if controle_id in ids_vistos:
                    continue
                ids_vistos.add(controle_id)
                candidatos.append((pontuacao, controle))

            # Em GuiTableControl, as células são objetos, mas alguns layouts só
            # expõem claramente o significado do campo no título da coluna.
            for tabela in controles_janela:
                if str(self._propriedade_sap(tabela, "Type", "")) != "GuiTableControl":
                    continue
                try:
                    colunas_tabela = self._colecao_sap_para_lista(tabela.Columns)
                except Exception:
                    colunas_tabela = []
                for indice_coluna, coluna in enumerate(colunas_tabela):
                    descricao_coluna = " ".join(
                        str(self._propriedade_sap(coluna, propriedade, ""))
                        for propriedade in ("Title", "Tooltip", "DefaultTooltip")
                    )
                    if not self._texto_indica_pep(descricao_coluna):
                        continue
                    try:
                        celula = tabela.GetCell(0, indice_coluna)
                    except Exception:
                        continue
                    celula_id = str(self._propriedade_sap(celula, "Id", ""))
                    if celula_id in ids_vistos:
                        continue
                    if str(self._propriedade_sap(celula, "Type", "")) not in (
                        "GuiCTextField",
                        "GuiTextField",
                    ):
                        continue
                    ids_vistos.add(celula_id)
                    candidatos.append((1900, celula))

        if not candidatos:
            return None

        candidatos.sort(key=lambda item: item[0], reverse=True)
        return candidatos[0][1]

    def _listar_abas_detalhe_item(self):
        """Lista as abas TABREQDT do detalhe do item, com as mais prováveis primeiro."""
        try:
            raiz = self.session.findById("wnd[0]/usr")
        except Exception:
            return []

        abas = {}

        for controle in self._iterar_controles_sap(raiz):
            tipo = str(self._propriedade_sap(controle, "Type", ""))
            nome = str(self._propriedade_sap(controle, "Name", ""))
            controle_id = str(self._propriedade_sap(controle, "Id", ""))
            if tipo != "GuiTab":
                continue
            if not (nome.startswith("TABREQDT") or "tabsREQ_ITEM_DETAIL" in controle_id):
                continue

            descricao = " ".join(
                str(self._propriedade_sap(controle, propriedade, ""))
                for propriedade in ("Text", "Tooltip", "AccText")
            ).strip()
            abas[controle_id or nome] = {
                "id": controle_id,
                "nome": nome,
                "descricao": descricao,
            }

        # Suplementa a árvore com FindByName, pois alguns temas não expõem todos
        # os GuiTab como filhos enumeráveis.
        for numero in range(1, 41):
            nome = f"TABREQDT{numero}"
            try:
                aba = raiz.findByName(nome, "GuiTab")
            except Exception:
                continue
            aba_id = str(self._propriedade_sap(aba, "Id", ""))
            descricao = " ".join(
                str(self._propriedade_sap(aba, propriedade, ""))
                for propriedade in ("Text", "Tooltip", "AccText")
            ).strip()
            abas[aba_id or nome] = {
                "id": aba_id,
                "nome": nome,
                "descricao": descricao,
            }

        def prioridade(info):
            descricao = self._normalizar_texto_sap(info["descricao"])
            pontos = 0
            if self._texto_indica_pep(descricao):
                pontos += 300
            if "contabil" in descricao and ("class" in descricao or "atrib" in descricao):
                pontos += 250
            if "account assignment" in descricao or "imputacion" in descricao:
                pontos += 250
            # Apenas uma preferência leve; nunca assume que 16 ou 7 é a aba correta.
            if info["nome"] in ("TABREQDT16", "TABREQDT7"):
                pontos += 5
            return (-pontos, info["nome"])

        return sorted(abas.values(), key=prioridade)

    def _localizar_campo_pep_varrendo_abas(self):
        """Seleciona cada aba do detalhe do item até encontrar o campo PEP real."""
        campo = self._localizar_campo_pep()
        if campo is not None:
            return campo

        for info in self._listar_abas_detalhe_item():
            aba = None
            if info["id"]:
                try:
                    aba = self.session.findById(info["id"])
                except Exception:
                    aba = None
            if aba is None and info["nome"]:
                try:
                    aba = self.session.findById("wnd[0]/usr").findByName(
                        info["nome"],
                        "GuiTab",
                    )
                except Exception:
                    aba = None
            if aba is None:
                continue

            try:
                aba.select()
                time.sleep(0.3)
            except Exception:
                continue

            campo = self._localizar_campo_pep()
            if campo is not None:
                descricao = info["descricao"] or "sem descrição exposta pelo SAP"
                self.logger.info(
                    "    -> Aba que contém o PEP localizada: %s (%s)",
                    info["nome"],
                    descricao,
                )
                return campo

        return None

    def _ler_status_sap(self):
        """Lê a barra de status da janela modal ou principal."""
        for status_id in ("wnd[1]/sbar", "wnd[0]/sbar"):
            try:
                barra = self.session.findById(status_id, False)
            except Exception:
                try:
                    barra = self.session.findById(status_id)
                except Exception:
                    barra = None
            if barra is None:
                continue
            tipo = str(self._propriedade_sap(barra, "MessageType", ""))
            texto = str(self._propriedade_sap(barra, "Text", "")).strip()
            if tipo or texto:
                return tipo, texto
        return "", ""

    def _escrever_validar_campo_pep(self, campo, pep):
        """Escreve o PEP em um GuiTextField/GuiCTextField e confirma no SAP."""
        try:
            campo.setFocus()
        except Exception:
            pass

        try:
            campo.Text = pep
        except Exception:
            campo.text = pep

        try:
            campo.caretPosition = len(pep)
        except Exception:
            pass

        try:
            valor_escrito = str(campo.Text).strip()
        except Exception:
            valor_escrito = str(self._propriedade_sap(campo, "text", "")).strip()

        if not valor_escrito:
            return False, "O SAP não manteve o valor digitado no campo do Elemento PEP."

        campo_id = str(self._propriedade_sap(campo, "Id", ""))
        janela_id = campo_id.split("/", 1)[0] if campo_id.startswith("wnd[") else "wnd[0]"
        self.logger.info(
            "    -> PEP escrito no campo SAP: %s",
            campo_id or str(self._propriedade_sap(campo, "Name", "campo sem ID")),
        )

        try:
            self.session.findById(janela_id).sendVKey(0)
        except Exception:
            self.session.findById("wnd[0]").sendVKey(0)
        time.sleep(0.8)

        # Em janela modal, o primeiro Enter pode apenas validar o campo. Confirme
        # o botão padrão antes de ler a barra principal, evitando interpretar a
        # mensagem antiga "Inserir Elemento PEP" como se fosse um erro novo.
        if janela_id != "wnd[0]":
            try:
                if self.session.findById(janela_id, False):
                    self._fechar_popup_sap()
                    time.sleep(0.4)
            except Exception:
                pass

        tipo_status, texto_status = self._ler_status_sap()
        if tipo_status in ("E", "A"):
            return False, texto_status or "O SAP rejeitou o Elemento PEP informado."

        # Se o campo continua na mesma tela, confirme que o SAP não apagou o valor.
        if campo_id:
            try:
                campo_confirmacao = self.session.findById(campo_id, False)
            except Exception:
                campo_confirmacao = None
            if campo_confirmacao is not None:
                try:
                    valor_confirmado = str(campo_confirmacao.Text).strip()
                except Exception:
                    valor_confirmado = str(
                        self._propriedade_sap(campo_confirmacao, "text", "")
                    ).strip()
                if not valor_confirmado:
                    return False, "O campo PEP ficou vazio após a validação do SAP."

        return True, texto_status or "PEP preenchido e validado."

    def _tentar_preencher_pep_no_grid(self, grid, indice, pep):
        """
        Fallback: tenta preencher o PEP diretamente numa coluna do grid de itens.

        Retorno:
            (True, mensagem)  -> coluna encontrada e valor aceito
            (False, mensagem) -> coluna encontrada, mas SAP rejeitou o valor
            (None, mensagem)  -> nenhuma coluna compatível existe no layout
        """
        colunas = [
            "COBL-PS_POSID",
            "MEACCT1100-PS_PSP_PNR",
            "PS_PSP_PNR",
            "PS_POSID",
            "POSID",
            "PSPEX",
        ]

        try:
            ordem = self._colecao_sap_para_lista(grid.ColumnOrder)
        except Exception:
            ordem = []

        for coluna in ordem:
            chave = str(coluna)
            textos = [chave]
            try:
                titulos = self._colecao_sap_para_lista(grid.GetColumnTitles(chave))
                textos.extend(str(titulo) for titulo in titulos)
            except Exception:
                pass
            combinado = " ".join(textos)
            combinado_upper = combinado.upper()
            if (
                self._texto_indica_pep(combinado)
                or "PS_PSP_PNR" in combinado_upper
                or "PS_POSID" in combinado_upper
            ):
                if chave not in colunas:
                    colunas.insert(0, chave)

        tentadas = []
        for coluna in colunas:
            if coluna in tentadas:
                continue
            tentadas.append(coluna)
            try:
                grid.modifyCell(indice, coluna, pep)
            except Exception:
                continue

            try:
                grid.triggerModified()
            except Exception:
                pass
            try:
                grid.setCurrentCell(indice, coluna)
                grid.pressEnter()
            except Exception:
                try:
                    self.session.findById("wnd[0]").sendVKey(0)
                except Exception:
                    pass
            time.sleep(0.8)

            tipo_status, texto_status = self._ler_status_sap()
            if tipo_status in ("E", "A"):
                return False, texto_status or f"O SAP rejeitou o PEP na coluna {coluna}."

            try:
                valor = str(grid.GetCellValue(indice, coluna)).strip()
            except Exception:
                valor = pep
            if not valor:
                return False, f"A coluna {coluna} permaneceu vazia após o preenchimento."

            self.logger.info(
                "    -> PEP preenchido diretamente na coluna do grid: %s",
                coluna,
            )
            return True, texto_status or f"PEP preenchido na coluna {coluna}."

        return None, "Nenhuma coluna PEP/WBS foi encontrada no grid de itens."

    def _diagnosticar_layout_pep(self):
        """Registra um resumo do layout para facilitar ajuste caso a busca ainda falhe."""
        abas = self._listar_abas_detalhe_item()
        resumo_abas = [
            f"{info['nome']}={info['descricao'] or '<sem texto>'}"
            for info in abas[:25]
        ]

        candidatos = []
        rotulos_pep = []
        for janela in reversed(self._janelas_sap_abertas()):
            for controle in self._iterar_controles_sap(janela, limite=5000):
                tipo = str(self._propriedade_sap(controle, "Type", ""))
                nome = str(self._propriedade_sap(controle, "Name", ""))
                controle_id = str(self._propriedade_sap(controle, "Id", ""))
                texto = str(self._propriedade_sap(controle, "Text", ""))
                tooltip = str(self._propriedade_sap(controle, "Tooltip", ""))
                combinado = f"{nome} {controle_id} {texto} {tooltip}"

                if tipo in ("GuiCTextField", "GuiTextField") and any(
                    marcador in combinado.upper()
                    for marcador in ("PS_", "POSID", "PSPNR", "WBS", "PEP")
                ):
                    candidatos.append(f"{tipo}|{nome}|{controle_id}")

                if tipo == "GuiLabel" and self._texto_indica_pep(f"{texto} {tooltip}"):
                    rotulos_pep.append(f"{texto or tooltip}|{controle_id}")

                if len(candidatos) >= 20 and len(rotulos_pep) >= 10:
                    break

        self.logger.warning(
            "    -> Diagnóstico PEP | Abas: %s",
            "; ".join(resumo_abas) or "nenhuma aba TABREQDT enumerada",
        )
        self.logger.warning(
            "    -> Diagnóstico PEP | Campos candidatos: %s",
            "; ".join(candidatos[:20]) or "nenhum",
        )
        self.logger.warning(
            "    -> Diagnóstico PEP | Rótulos encontrados: %s",
            "; ".join(rotulos_pep[:10]) or "nenhum",
        )

        return (
            f"Abas verificadas: {', '.join(info['nome'] for info in abas) or 'nenhuma'}. "
            "O diagnóstico técnico foi gravado no log."
        )

    def _preencher_pep_itens(self, grid, itens_com_pep, contexto="criação"):
        """
        Preenche o Elemento PEP de cada item e retorna:
            {indice_grid: (sucesso: bool, mensagem: str)}
        """
        resultados = {}

        for item_pep in itens_com_pep:
            idx = int(item_pep['grid_index'])
            pep = str(item_pep.get('pep', '')).strip()
            material = str(item_pep.get('material', '')).strip()
            resultados[idx] = (False, "Elemento PEP não preenchido.")

            if not pep:
                resultados[idx] = (False, "Valor do PEP está vazio na planilha.")
                continue

            try:
                self.logger.info(
                    "  -> [%s] Preenchendo PEP '%s' para item %s (Mat: %s)",
                    contexto,
                    pep,
                    idx + 1,
                    material,
                )

                # Seleciona explicitamente a linha para sincronizar o detalhe do item.
                try:
                    grid.setCurrentCell(idx, "KNTTP")
                except Exception:
                    grid.setCurrentCell(idx, "MATNR")
                try:
                    grid.selectedRows = str(idx)
                except Exception:
                    pass
                try:
                    grid.clickCurrentCell()
                except Exception:
                    pass
                try:
                    grid.currentCellMoved()
                except Exception:
                    pass

                # Se já existe wnd[1], ela pode ser a própria tela de classificação.
                # Não a feche antes de procurar o campo PEP.
                try:
                    popup_aberto = self.session.findById("wnd[1]", False)
                except Exception:
                    popup_aberto = None

                if popup_aberto is None:
                    try:
                        self.session.findById("wnd[0]").sendVKey(0)
                    except Exception:
                        pass
                    time.sleep(0.6)

                # 1) Procura na tela atual e em eventual janela modal.
                campo = self._localizar_campo_pep()

                # Se há popup, mas ele não contém campo PEP, confirma-o para liberar
                # a navegação pelas abas do item.
                if campo is None:
                    try:
                        popup = self.session.findById("wnd[1]", False)
                    except Exception:
                        popup = None
                    if popup is not None:
                        titulo_popup = str(self._propriedade_sap(popup, "Text", "")).strip()
                        self.logger.info(
                            "    -> Popup sem campo PEP%s; confirmando para continuar.",
                            f" ({titulo_popup})" if titulo_popup else "",
                        )
                        self._fechar_popup_sap()
                        time.sleep(0.4)

                # 2) Varre todas as abas TABREQDT. Não presume mais que TABREQDT16
                # ou TABREQDT7 seja necessariamente a aba contábil.
                if campo is None:
                    campo = self._localizar_campo_pep_varrendo_abas()

                if campo is not None:
                    ok, mensagem = self._escrever_validar_campo_pep(campo, pep)
                    resultados[idx] = (ok, mensagem)
                    if ok:
                        self.logger.info("    -> PEP confirmado para o item %s.", idx + 1)
                    else:
                        self.logger.warning("    -> SAP rejeitou o PEP: %s", mensagem)
                    continue

                # 3) Fallback para layouts em que o PEP está como coluna do grid.
                status_grid, mensagem_grid = self._tentar_preencher_pep_no_grid(
                    grid,
                    idx,
                    pep,
                )
                if status_grid is not None:
                    resultados[idx] = (bool(status_grid), mensagem_grid)
                    if status_grid:
                        self.logger.info("    -> PEP confirmado para o item %s.", idx + 1)
                    else:
                        self.logger.warning("    -> SAP rejeitou o PEP: %s", mensagem_grid)
                    continue

                diagnostico = self._diagnosticar_layout_pep()
                mensagem = (
                    "Campo do Elemento PEP não localizado. Foram testados os nomes "
                    "COBL-PS_POSID e MEACCT1100-PS_PSP_PNR, todas as abas do item "
                    f"e as colunas do grid. {diagnostico}"
                )
                self.logger.warning("    -> %s", mensagem)
                resultados[idx] = (False, mensagem)

            except Exception as e:
                mensagem = f"Erro técnico ao preencher o PEP: {str(e)}"
                self.logger.warning(
                    "  -> Erro ao preencher PEP para item %s: %s",
                    idx + 1,
                    e,
                )
                resultados[idx] = (False, mensagem)

        self.logger.info("Preenchimento de PEP concluído (%s).", contexto)
        return resultados

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
