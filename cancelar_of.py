import win32com.client
import gspread
from google.oauth2.service_account import Credentials
import time

def concluir_ofs():
    print("Iniciando o processo...")

    # ---------------------------------------------------------
    # 1. CONFIGURAÇÃO DO GOOGLE SHEETS
    # ---------------------------------------------------------
    # Define as permissões que o script terá (Drive e Sheets)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    try:
        # Carrega o arquivo JSON que você baixou do Google Cloud
        creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
        client = gspread.authorize(creds)
        
        # Abre a planilha pelo nome e seleciona a aba específica
        planilha = client.open("MAPEAMENTO PLANNING")
        aba = planilha.worksheet("CANCELAR OF")
    except Exception as e:
        print(f"Erro ao conectar no Google Sheets. Verifique o credentials.json e os compartilhamentos: {e}")
        return

    # ---------------------------------------------------------
    # 2. CONFIGURAÇÃO DO SAP GUI
    # ---------------------------------------------------------
    try:
        # Pega a sessão ativa do SAP (o SAP precisa estar aberto!)
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
    except Exception as e:
        print("Erro ao conectar ao SAP. Certifique-se de que o SAP está aberto e logado.")
        return

    # ---------------------------------------------------------
    # 3. LÓGICA DE REPETIÇÃO (O "While" do seu VBA)
    # ---------------------------------------------------------
    # Pega todos os valores da Coluna A (Índice 1 no gspread)
    valores_coluna_a = aba.col_values(1)
    
    # Assumindo que a linha 1 seja cabeçalho, vamos começar da linha 2.
    # Se não tiver cabeçalho, mude a linha_atual para 1.
    linha_atual = 2 
    
    for i in range(linha_atual - 1, len(valores_coluna_a)):
        selected_of = str(valores_coluna_a[i]).strip()
        
        # Se a célula estiver vazia, encerra o loop (como o <> "" do VBA)
        if not selected_of:
            break
            
        try:
            # Maximiza e chama a transação
            session.findById("wnd").maximize()
            session.findById("wnd/tbar/okcd").Text = "/NCO02"
            session.findById("wnd").sendVKey(0) # Enter
            
            # Preenche o número da OF
            session.findById("wnd/usr/ctxtCAUFVD-AUFNR").Text = selected_of
            session.findById("wnd").sendVKey(0) # Enter
            
            # Navega no menu
            session.findById("wnd/mbar/menu/menu/menu").Select()
            session.findById("wnd").sendVKey(11) # Salvar (Ctrl+S)
            
            # Posiciona o cursor e dá enter
            session.findById("wnd/usr/ctxtCAUFVD-AUFNR").caretPosition = 7
            session.findById("wnd").sendVKey(0)
            
            # Escreve "FEITO" na Coluna B (Índice 2)
            aba.update_cell(linha_atual, 2, "FEITO")
            print(f"Linha {linha_atual}: OF {selected_of} -> FEITO")
            
        except Exception as e:
            # Em caso de erro (On Error GoTo Handler)
            aba.update_cell(linha_atual, 2, "ERRO")
            print(f"Linha {linha_atual}: OF {selected_of} -> ERRO ({e})")
            
            # Volta para a tela inicial para não travar o loop na próxima OF
            session.findById("wnd/tbar/okcd").Text = "/N"
            session.findById("wnd").sendVKey(0)
            
        linha_atual += 1
        
        # Pausa de 1 segundo para não estourar o limite de requisições da API do Google
        time.sleep(1)

    print("\nProcesso concluído com sucesso!")

# Executa a função
if __name__ == "__main__":
    concluir_ofs()