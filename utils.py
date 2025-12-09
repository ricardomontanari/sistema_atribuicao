import sys
import os
import keyboard
import time
import subprocess
import pandas as pd
import utils 
from datetime import datetime # NOVO: Import para obter a hora atual

# --- CONSTANTES E CONFIGURAÇÕES GLOBAIS ---
# NOME DO ARQUIVO ALVO DO EXCEL
NOME_ARQUIVO_ALVO = 'atribuicao.xlsx' 
DEFAULT_DELAY_SECONDS = 0.2
VERSAO_SISTEMA = '1.0.1' # Versão atualizada

# Variáveis de Estado Global
PARAR_AUTOMACAO = False 
INDICE_ATUAL_DO_CICLO = 0
DELAY_ATUAL = DEFAULT_DELAY_SECONDS
monitor_thread_started = False 

# --- FUNÇÕES DE UTILIDADE ---

def log_time(message):
    """Retorna uma mensagem com o carimbo de tempo no formato [HH:MM:SS]."""
    # NOVO: Formata a hora atual e concatena com a mensagem
    return f"[{datetime.now().strftime('%H:%M:%S')}] {message}"


def resource_path(relative_path):
    """
    Obtém o caminho absoluto para recursos INTERNOS (embutidos no EXE, como imagens).
    ESTA FUNÇÃO DEVE SER USADA APENAS PARA RECURSOS EMBUTIDOS (IMAGENS DE ERRO).
    """
    try:
        # Caso 1: Código rodando dentro do executável (pasta temporária)
        base_path = sys._MEIPASS
    except Exception:
        # Caso 2: Código rodando no ambiente de desenvolvimento
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def get_external_path(relative_path):
    """
    Obtém o caminho absoluto para ARQUIVOS EXTERNOS (Excel e DB), 
    usando o diretório de onde o executável (ou script) está rodando.
    """
    base_path = os.path.dirname(sys.argv[0])
    
    return os.path.join(base_path, relative_path)


def abrir_planilha_alvo():
    """Localiza o arquivo alvo na pasta ATUAL DO EXECUTÁVEL e o abre."""
    caminho_arquivo = get_external_path(NOME_ARQUIVO_ALVO)
    
    if not os.path.exists(caminho_arquivo):
        # NOVO: Usa log_time no retorno de erro
        return False, log_time(f"❌ ERRO: Arquivo '{NOME_ARQUIVO_ALVO}' não encontrado em:\n{caminho_arquivo}")
        
    try:
        if sys.platform == "win32":
            subprocess.Popen(['start', '', caminho_arquivo], shell=True)
        else:
            opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
            subprocess.Popen([opener, caminho_arquivo])
            
        return True, log_time(f"✅ Planilha '{NOME_ARQUIVO_ALVO}' aberta com sucesso.")

    except Exception as e:
        return False, log_time(f"❌ ERRO ao tentar abrir o arquivo: {e}")

def monitorar_tecla_escape(log_textbox):
    """
    Monitora a tecla ESC e define a flag PARAR_AUTOMACAO como True (PAUSA).
    """
    while True: 
        try:
            keyboard.wait('esc')
            
            # APENAS PAUSA se ainda não estiver pausado
            if not utils.PARAR_AUTOMACAO: 
                utils.PARAR_AUTOMACAO = True 
                # NOVO: Usa log_time
                log_textbox.insert("end", utils.log_time("\n[ESC] PAUSA solicitada. Aguardando fim do ciclo.\n"))
                log_textbox.see("end")

            time.sleep(0.5) 
            
        except Exception:
            time.sleep(1)

def validar_e_obter_delay(delay_input):
    """Lê, valida e retorna o valor do delay em segundos (float)."""
    if not delay_input:
        return DEFAULT_DELAY_SECONDS
    
    try:
        delay_valor = float(str(delay_input).replace(',', '.'))
        if delay_valor < 0:
             raise ValueError("O Delay deve ser zero ou um valor positivo.")
        return delay_valor
    except ValueError:
        raise ValueError("Delay deve ser um número válido (ex: 0.5, 1.0).")

def ler_e_filtrar_dados(nome_arquivo, cidade_filtro, backlog_filtro, log_textbox):
    """
    Lê a planilha e aplica os filtros: Backlog time(Station) e Cidades.
    """
    caminho_arquivo = get_external_path(nome_arquivo)
    
    if not os.path.exists(caminho_arquivo):
        msg = f"❌ ERRO: Arquivo '{nome_arquivo}' não encontrado."
        # NOVO: Usa log_time no retorno de erro
        return None, 0, utils.log_time(msg)
        
    try:
        df = pd.read_excel(caminho_arquivo)
        
        # --- FILTRO 1: Validação de Backlog time(Station) ---
        col_backlog = 'Backlog time(Station)'
        if col_backlog in df.columns:
            if not backlog_filtro:
                backlog_filtro = "1"
            
            # Suporta múltiplos backlogs (ex: "1, 2")
            lista_backlogs = [b.strip() for b in str(backlog_filtro).split(',') if b.strip()]
            df = df[df[col_backlog].astype(str).str.strip().isin(lista_backlogs)]
            
        else:
             # NOVO: Usa log_time
             log_textbox.insert("end", utils.log_time(f"⚠️ Aviso: Coluna '{col_backlog}' não encontrada. Filtro de backlog ignorado.\n"))

        # Validação de colunas de Cidade
        coluna_filtro = None
        if 'Destination City' in df.columns:
             coluna_filtro = 'Destination City'
        elif 'Cidade' in df.columns:
             coluna_filtro = 'Cidade'
        
        if not coluna_filtro:
             # NOVO: Usa log_time
             log_textbox.insert("end", utils.log_time("⚠️ Aviso: Coluna de cidade não encontrada. Verifique o Excel.\n"))


        # --- FILTRO 2: Aplica o filtro de Cidades ---
        if cidade_filtro and cidade_filtro != "NENHUM FILTRO" and cidade_filtro != "SELECIONE" and coluna_filtro:
            
            # Suporta múltiplas cidades (ex: "CURITIBA, SÃO PAULO")
            lista_cidades = [c.strip().upper() for c in cidade_filtro.split(',') if c.strip()]
            
            df_filtrado = df[df[coluna_filtro].astype(str).str.upper().isin(lista_cidades)]
            
            msg = f"✅ Planilha lida (Backlog={backlog_filtro}). Filtro Cidades: {lista_cidades} -> {len(df_filtrado)} registros."
        else:
            df_filtrado = df
            msg = f"✅ Planilha lida (Backlog={backlog_filtro}). Todas as cidades ({len(df)} registros)."
            
        repetir = len(df_filtrado)
        # NOVO: Usa log_time no retorno de sucesso
        return df_filtrado, repetir, utils.log_time(msg)
        
    except Exception as e:
        msg = f"❌ ERRO ao processar Excel: {e}"
        return None, 0, utils.log_time(msg)