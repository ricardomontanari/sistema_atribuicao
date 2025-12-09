import sys
import os
import keyboard
import time
import subprocess
import pandas as pd
import utils # Importa a si mesmo para garantir referências explícitas ao escopo global

# --- CONSTANTES E CONFIGURAÇÕES GLOBAIS ---
NOME_ARQUIVO_ALVO = 'atribuicao.xlsx' 
DEFAULT_DELAY_SECONDS = 0.2
VERSAO_SISTEMA = 'Beta 3.2 (Delay Dinâmico)'

# Variáveis de Estado Global (Acessadas por gui.py e automation_logic.py)
PARAR_AUTOMACAO = False 
INDICE_ATUAL_DO_CICLO = 0
DELAY_ATUAL = DEFAULT_DELAY_SECONDS # <--- NOVA VARIÁVEL GLOBAL PARA CONTROLE DINÂMICO

# --- FUNÇÕES DE UTILIDADE ---

def resource_path(relative_path):
    """
    Obtém o caminho absoluto para o recurso, seja no ambiente de desenvolvimento
    ou no executável PyInstaller (--onefile).
    """
    try:
        # Caso 1: Código rodando dentro do executável (pasta temporária)
        base_path = sys._MEIPASS
    except Exception:
        # Caso 2: Código rodando no ambiente de desenvolvimento
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def abrir_planilha_alvo():
    """
    Localiza o arquivo alvo na pasta atual/executável e o abre com o programa 
    padrão do sistema (geralmente Excel).
    """
    caminho_arquivo = resource_path(NOME_ARQUIVO_ALVO)
    
    if not os.path.exists(caminho_arquivo):
        return False, f"❌ ERRO: Arquivo '{NOME_ARQUIVO_ALVO}' não encontrado em:\n{caminho_arquivo}"
        
    try:
        if sys.platform == "win32":
            os.startfile(caminho_arquivo)
        else:
            # Suporte básico para macOS/Linux
            opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
            subprocess.Popen([opener, caminho_arquivo])
            
        return True, f"✅ Planilha '{NOME_ARQUIVO_ALVO}' aberta com sucesso."

    except Exception as e:
        return False, f"❌ ERRO ao tentar abrir o arquivo: {e}"

def monitorar_tecla_escape(log_textbox):
    """
    Monitora a tecla ESC e define a flag PARAR_AUTOMACAO como True (PAUSA).
    Usa referência explícita 'utils.PARAR_AUTOMACAO' para garantir sincronia.
    """
    while True: 
        try:
            keyboard.wait('esc')
            
            # Apenas PAUSA se ainda não estiver pausado
            if not utils.PARAR_AUTOMACAO: 
                utils.PARAR_AUTOMACAO = True 
                log_textbox.insert("end", "\n[ESC] PAUSA solicitada. Aguardando fim do ciclo.\n")
                log_textbox.see("end")

            time.sleep(0.5) 
            
        except Exception:
            time.sleep(1)

def validar_e_obter_delay(delay_input):
    """Lê, valida e retorna o valor do delay em segundos (float)."""
    if not delay_input:
        return DEFAULT_DELAY_SECONDS
    
    try:
        # Substitui vírgula por ponto para suportar formatos decimais BR
        delay_valor = float(str(delay_input).replace(',', '.'))
        if delay_valor < 0:
             raise ValueError("O Delay deve ser zero ou um valor positivo.")
        return delay_valor
    except ValueError:
        raise ValueError("Delay deve ser um número válido (ex: 0.5, 1.0).")

def ler_e_filtrar_dados(nome_arquivo, cidade_filtro, log_textbox):
    """
    Lê a planilha, filtra os dados pela cidade e retorna o DataFrame filtrado, 
    o número de linhas e uma mensagem de log.
    """
    caminho_arquivo = resource_path(nome_arquivo)
    
    if not os.path.exists(caminho_arquivo):
        msg = f"❌ ERRO: Arquivo '{nome_arquivo}' não encontrado."
        return None, 0, msg
        
    try:
        df = pd.read_excel(caminho_arquivo)
        
        # Validação básica de colunas
        if 'Destination City' not in df.columns and 'Cidade' not in df.columns:
             log_textbox.insert("end", "⚠️ Aviso: Coluna 'Destination City' não encontrada. Verifique o Excel.\n")

        # Aplica o filtro
        if cidade_filtro and cidade_filtro != "NENHUM FILTRO" and cidade_filtro != "SELECIONE":
            coluna_filtro = 'Destination City' if 'Destination City' in df.columns else 'Cidade'
            
            if coluna_filtro in df.columns:
                df_filtrado = df[df[coluna_filtro].astype(str).str.upper() == cidade_filtro.upper()]
                msg = f"✅ Planilha lida. Filtro '{cidade_filtro}' aplicado ({len(df_filtrado)} registros)."
            else:
                df_filtrado = df
                msg = f"⚠️ Coluna de cidade não encontrada. Filtro '{cidade_filtro}' ignorado."
        else:
            df_filtrado = df
            msg = f"✅ Planilha lida. Nenhum filtro aplicado ({len(df)} registros)."
            
        repetir = len(df_filtrado)
        return df_filtrado, repetir, msg
        
    except Exception as e:
        msg = f"❌ ERRO ao processar Excel: {e}"
        return None, 0, msg