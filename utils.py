import sys
import os
import keyboard
import time
import pandas as pd
import subprocess

# ####################################################################
# --- CONSTANTES GLOBAIS ---
# ####################################################################

NOME_ARQUIVO_ALVO = 'atribuicao.xlsx' 
DEFAULT_DELAY_SECONDS = 0.2
VERSAO_SISTEMA = 'v3.8 (Assets Fix)' 

# Variáveis de Estado Global
PARAR_AUTOMACAO = False 
CANCELAR_AUTOMACAO = False 
INDICE_ATUAL_DO_CICLO = 0
DELAY_ATUAL = DEFAULT_DELAY_SECONDS

# ####################################################################
# --- GERENCIADOR DE CAMINHOS ---
# ####################################################################

def resource_path(relative_path):
    """
    Retorna o caminho absoluto para recursos internos (dentro do EXE).
    Use para: Imagens de erro, ícones, assets fixos.
    
    Exemplo: resource_path('assets/erro_baixada.png')
    """
    try:
        # PyInstaller cria uma pasta temporária em _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # Modo de desenvolvimento (pasta atual)
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

def get_external_path(relative_path):
    """
    Retorna o caminho para arquivos externos (ao lado do EXE/Script).
    Use para: Planilhas Excel, Banco de Dados, Configs JSON, Logs.
    
    Exemplo: get_external_path('atribuicao.xlsx')
    """
    if getattr(sys, 'frozen', False):
        # Se for executável, pega a pasta onde o .exe está
        app_path = os.path.dirname(sys.executable)
    else:
        # Se for script, pega a pasta do script .py
        app_path = os.path.dirname(os.path.abspath(__file__))
    
    return os.path.join(app_path, relative_path)

# ####################################################################
# --- FUNÇÕES DE SISTEMA ---
# ####################################################################

def abrir_planilha_alvo():
    """Tenta abrir a planilha Excel usando o caminho externo."""
    caminho = get_external_path(NOME_ARQUIVO_ALVO)
    if not os.path.exists(caminho):
        return False, f"Arquivo não encontrado: {caminho}"
    try:
        os.startfile(caminho)
        return True, "Planilha aberta com sucesso."
    except Exception as e:
        return False, str(e)

def monitorar_tecla_escape(log_textbox):
    """Monitora a tecla ESC em background para pausar a automação."""
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
    """Converte a entrada de delay da GUI para float seguro."""
    if not delay_input: return DEFAULT_DELAY_SECONDS
    try:
        val = float(str(delay_input).replace(',', '.'))
        return val if val >= 0 else DEFAULT_DELAY_SECONDS
    except: return DEFAULT_DELAY_SECONDS

# ####################################################################
# --- LEITURA EXCEL ---
# ####################################################################

def ler_e_filtrar_dados(arquivo, cidade_filtro, backlog_filtro, log_textbox):
    """Lê o Excel, aplica filtros e retorna o DataFrame filtrado."""
    try:
        path = get_external_path(arquivo)
        if not os.path.exists(path):
             return pd.DataFrame(), 0, f"Planilha não encontrada em: {path}"

        # Lê como string para preservar zeros à esquerda (ex: IDs)
        df = pd.read_excel(path, dtype=str)
        df.columns = df.columns.str.strip()
        
        # 1. Filtro Scan (Prioritário)
        col_scan = next((c for c in df.columns if "scan type" in c.lower()), None)
        if col_scan:
            # Filtra apenas itens "recebido no DS"
            df = df[df[col_scan].str.strip().isin(["(recebido no DS)", "recebido no DS"])]
        
        # 2. Filtro Cidade
        if cidade_filtro and cidade_filtro not in ["SELECIONE", "NENHUM FILTRO", ""]:
            col = 'Destination City' if 'Destination City' in df.columns else 'Cidade'
            if col in df.columns:
                lista = [c.strip().upper() for c in cidade_filtro.split(',') if c.strip()]
                df = df[df[col].astype(str).str.strip().str.upper().isin(lista)]

        # 3. Filtro Backlog
        if backlog_filtro and str(backlog_filtro).strip():
            col_bk = next((c for c in ['Backlog', 'Backlog time(Station)'] if c in df.columns), None)
            if col_bk:
                # Converte para numérico apenas para comparar
                df = df[pd.to_numeric(df[col_bk], errors='coerce') == int(backlog_filtro)]

        if df.empty: return pd.DataFrame(), 0, "Nenhum dado encontrado após aplicar os filtros."
        return df, len(df), f"{len(df)} registros carregados e prontos."

    except Exception as e:
        return pd.DataFrame(), 0, f"Erro crítico na leitura: {e}"