import sys
import os
import keyboard
import time
import pandas as pd
import subprocess

# --- CONSTANTES ---
NOME_ARQUIVO_ALVO = 'atribuicao.xlsx' 
DEFAULT_DELAY_SECONDS = 0.2
VERSAO_SISTEMA = 'v3.9 (Silent & Robust)' 

# Estado Global
PARAR_AUTOMACAO = False 
CANCELAR_AUTOMACAO = False 
INDICE_ATUAL_DO_CICLO = 0
DELAY_ATUAL = DEFAULT_DELAY_SECONDS

# --- CAMINHOS ---
def resource_path(relative_path):
    """Caminho absoluto para recursos internos (Assets no EXE)."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_external_path(relative_path):
    """Caminho para arquivos editáveis pelo usuário (Excel/Config)."""
    if getattr(sys, 'frozen', False):
        app_path = os.path.dirname(sys.executable)
    else:
        app_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(app_path, relative_path)

# --- SISTEMA ---
def abrir_planilha_alvo():
    caminho = get_external_path(NOME_ARQUIVO_ALVO)
    if not os.path.exists(caminho):
        return False, f"Arquivo não encontrado: {caminho}"
    try:
        os.startfile(caminho)
        return True, "Planilha aberta."
    except Exception as e:
        return False, str(e)

def monitorar_tecla_escape(log_textbox):
    while True: 
        try:
            keyboard.wait('esc')
            if not globals()['PARAR_AUTOMACAO']: 
                globals()['PARAR_AUTOMACAO'] = True 
                log_textbox.insert("end", "\n[ESC] Pausa solicitada.\n")
                log_textbox.see("end")
            time.sleep(0.5)
        except: time.sleep(1)

def validar_e_obter_delay(delay_input):
    if not delay_input: return DEFAULT_DELAY_SECONDS
    try:
        val = float(str(delay_input).replace(',', '.'))
        return val if val >= 0 else DEFAULT_DELAY_SECONDS
    except: return DEFAULT_DELAY_SECONDS

# --- LEITURA EXCEL ---
def ler_e_filtrar_dados(arquivo, cidade_filtro, backlog_filtro, log_textbox):
    try:
        path = get_external_path(arquivo)
        if not os.path.exists(path):
             return pd.DataFrame(), 0, f"Planilha não encontrada: {path}"

        df = pd.read_excel(path, dtype=str)
        df.columns = df.columns.str.strip()
        
        # Filtro 1: Scan Type (Recebido no DS)
        col_scan = next((c for c in df.columns if "scan type" in c.lower()), None)
        if col_scan:
            df = df[df[col_scan].str.strip().isin(["(recebido no DS)", "recebido no DS"])]
        
        # Filtro 2: Cidade
        if cidade_filtro and cidade_filtro not in ["SELECIONE", "NENHUM FILTRO", ""]:
            col = 'Destination City' if 'Destination City' in df.columns else 'Cidade'
            if col in df.columns:
                lista = [c.strip().upper() for c in cidade_filtro.split(',') if c.strip()]
                df = df[df[col].astype(str).str.strip().str.upper().isin(lista)]

        # Filtro 3: Backlog
        if backlog_filtro and str(backlog_filtro).strip():
            col_bk = next((c for c in ['Backlog', 'Backlog time(Station)'] if c in df.columns), None)
            if col_bk:
                df = df[pd.to_numeric(df[col_bk], errors='coerce') == int(backlog_filtro)]

        if df.empty: return pd.DataFrame(), 0, "Nenhum dado após filtros."
        return df, len(df), f"{len(df)} registros carregados."

    except Exception as e:
        return pd.DataFrame(), 0, f"Erro leitura: {e}"