import sys
import os
import keyboard
import time
import pandas as pd
import utils 

# ####################################################################
# --- CONSTANTES GLOBAIS ---
# ####################################################################

NOME_ARQUIVO_ALVO = 'atribuicao.xlsx' 
DEFAULT_DELAY_SECONDS = 0.2
VERSAO_SISTEMA = 'Beta 1.2 (Stop Button)'

# Variáveis de Estado Global
PARAR_AUTOMACAO = False 
CANCELAR_AUTOMACAO = False # <--- NOVA FLAG: Permite abortar a execução
INDICE_ATUAL_DO_CICLO = 0
DELAY_ATUAL = DEFAULT_DELAY_SECONDS

# ####################################################################
# --- FUNÇÕES DE SISTEMA ---
# ####################################################################

def resource_path(relative_path):
    """
    Obtém o caminho absoluto para recursos (imagens/ícones),
    funcionando tanto em desenvolvimento quanto no executável.
    """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def abrir_planilha_alvo():
    """
    Tenta abrir o arquivo Excel alvo.
    """
    if not os.path.exists(NOME_ARQUIVO_ALVO):
        return False, f"❌ Arquivo '{NOME_ARQUIVO_ALVO}' não encontrado."
    try:
        os.startfile(NOME_ARQUIVO_ALVO)
        return True, f"✅ Planilha '{NOME_ARQUIVO_ALVO}' aberta."
    except Exception as e:
        return False, f"❌ Erro ao abrir arquivo: {e}"

def validar_e_obter_delay(delay_input):
    """
    Valida e converte o delay da interface para float.
    """
    if not delay_input: 
        return DEFAULT_DELAY_SECONDS
    try:
        val = float(str(delay_input).replace(',', '.'))
        return val if val >= 0 else DEFAULT_DELAY_SECONDS
    except: 
        return DEFAULT_DELAY_SECONDS

# ####################################################################
# --- MONITORAMENTO DE TECLADO ---
# ####################################################################

def monitorar_tecla_escape(log_textbox):
    """
    Thread que monitora a tecla ESC para solicitar pausa.
    """
    while True: 
        try:
            keyboard.wait('esc')
            if not PARAR_AUTOMACAO: 
                # Usa globals() para modificar a variável deste próprio módulo
                globals()['PARAR_AUTOMACAO'] = True 
                log_textbox.insert("end", "\n[ESC] PAUSA solicitada. Aguardando fim do ciclo atual...\n")
                log_textbox.see("end")
            time.sleep(0.5) 
        except Exception:
            time.sleep(1)

# ####################################################################
# --- LEITURA DE DADOS ---
# ####################################################################

def ler_e_filtrar_dados(arquivo, cidade_filtro, backlog_filtro, log_textbox):
    """
    Lê estritamente a aba chamada 'sheet1' (case-insensitive).
    """
    if not os.path.exists(arquivo):
        raise FileNotFoundError(f"Arquivo '{arquivo}' não encontrado.")
    
    try:
        # 1. Carrega o Excel
        xls = pd.ExcelFile(arquivo)
        abas = xls.sheet_names
        df = None
        
        log_textbox.insert("end", f"📂 Buscando aba 'Sheet1' no arquivo...\n")

        # 2. Busca Estrita por 'sheet1'
        aba_encontrada = None
        for aba in abas:
            if aba.strip().lower() == "sheet1":
                aba_encontrada = aba
                break
        
        if aba_encontrada:
            df = pd.read_excel(xls, sheet_name=aba_encontrada)
            log_textbox.insert("end", f"✅ Aba '{aba_encontrada}' carregada com sucesso.\n")
        else:
            raise ValueError("Aba 'Sheet1' não encontrada. Verifique o nome da aba no Excel.")

        # 3. Validação de Colunas Mínimas
        colunas = [c.strip() for c in df.columns]
        tem_id = any(c in colunas for c in ['Waybill No', 'Motorista ID'])
        
        if not tem_id:
            log_textbox.insert("end", "⚠️ Aviso: Coluna de ID ('Waybill No') não encontrada na Sheet1.\n")

        # 4. Filtro de Backlog
        registros_iniciais = len(df)
        if backlog_filtro and str(backlog_filtro).strip():
            if 'Backlog' in df.columns:
                try:
                    val = int(backlog_filtro)
                    df = df[df['Backlog'] == val]
                except: pass 
            else:
                 log_textbox.insert("end", "⚠️ Coluna 'Backlog' não existe. Filtro ignorado.\n")

        # 5. Filtro de Cidade
        if cidade_filtro and cidade_filtro not in ["NENHUM FILTRO", "SELECIONE", ""]:
            col_cidade = 'Destination City' if 'Destination City' in df.columns else 'Cidade'
            
            if col_cidade in df.columns:
                lista_desejada = [c.strip().upper() for c in cidade_filtro.split(',') if c.strip()]
                df = df[df[col_cidade].astype(str).str.strip().str.upper().isin(lista_desejada)]
            else:
                log_textbox.insert("end", f"⚠️ Coluna '{col_cidade}' não encontrada. Filtro ignorado.\n")

        count = len(df)
        msg = f"📊 Processamento: {registros_iniciais} registros -> {count} filtrados."
        return df, count, msg

    except Exception as e:
        raise ValueError(f"Erro ao ler Excel: {e}")