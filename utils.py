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
VERSAO_SISTEMA = 'Beta 1.6 (Fix Scan)'

# Variáveis de Estado Global
PARAR_AUTOMACAO = False 
CANCELAR_AUTOMACAO = False
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
        return True, f"✅ Planilha '{NOME_ARQUIVO_ALVO}' aberta com sucesso."
    except Exception as e:
        return False, f"❌ ERRO ao tentar abrir o arquivo: {e}"

def monitorar_tecla_escape(log_textbox):
    """
    Monitora a tecla ESC e define a flag PARAR_AUTOMACAO como True (PAUSA).
    """
    while True: 
        try:
            keyboard.wait('esc')
            if not globals()['PARAR_AUTOMACAO']: 
                globals()['PARAR_AUTOMACAO'] = True 
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
        delay_valor = float(str(delay_input).replace(',', '.'))
        if delay_valor < 0: raise ValueError
        return delay_valor
    except ValueError:
        return DEFAULT_DELAY_SECONDS

# ####################################################################
# --- LEITURA E FILTRAGEM (ROBUSTA) ---
# ####################################################################

def ler_e_filtrar_dados(arquivo, cidade_filtro, backlog_filtro, log_textbox):
    """
    Lê o Excel e aplica os filtros:
    1. Last scan type == "(recebido no DS)" OU "recebido no DS"
    2. Cidades
    3. Backlog (Numérico)
    """
    try:
        # 1. Carregar Excel
        log_textbox.insert("end", f"Lendo arquivo '{arquivo}'...\n")
        
        # Lê o arquivo forçando string inicialmente
        df = pd.read_excel(arquivo, dtype=str)
        
        # Limpa espaços dos nomes das colunas
        df.columns = df.columns.str.strip()
        
        total_inicial = len(df)

        # --- NOVO FILTRO: Last scan type ---
        col_scan = None
        # Busca case-insensitive pela coluna
        for col in df.columns:
            if col.lower() == "last scan type":
                col_scan = col
                break
        
        if col_scan:
            # CORREÇÃO AQUI: Aceita com e sem parênteses
            termos_validos = ["(recebido no DS)", "recebido no DS"]
            
            # Filtra onde a coluna contem um dos termos válidos
            df_scan = df[df[col_scan].str.strip().isin(termos_validos)]
            
            if len(df_scan) == 0:
                log_textbox.insert("end", f"⚠️ Nenhum registro com status 'recebido no DS' encontrado.\n")
                unicos = df[col_scan].unique()[:3]
                log_textbox.insert("end", f"   Valores encontrados: {unicos}\n")
                # Retorna vazio imediatamente para evitar logs confusos depois
                return pd.DataFrame(), 0, "Filtro de Status resultou em 0."
            else:
                log_textbox.insert("end", f"ℹ️ Filtro Scan aplicado: {len(df_scan)} de {total_inicial} registros.\n")
            
            df = df_scan
        else:
            log_textbox.insert("end", "⚠️ Coluna 'Last scan type' não encontrada. Filtro ignorado.\n")


        # --- FILTRO CIDADE ---
        if cidade_filtro and cidade_filtro not in ["NENHUM FILTRO", "SELECIONE", ""]:
            col_cidade = 'Destination City' if 'Destination City' in df.columns else 'Cidade'
            
            if col_cidade in df.columns:
                lista_desejada = [c.strip().upper() for c in cidade_filtro.split(',') if c.strip()]
                # Filtra
                df = df[df[col_cidade].astype(str).str.strip().str.upper().isin(lista_desejada)]
                
                if len(df) == 0:
                    log_textbox.insert("end", f"⚠️ Nenhuma cidade encontrada para: {cidade_filtro}.\n")
                    return pd.DataFrame(), 0, "Filtro de Cidade resultou em 0."
                    
                log_textbox.insert("end", f"ℹ️ Filtro Cidade aplicado: {len(df)} registros restantes.\n")
            else:
                log_textbox.insert("end", f"⚠️ Coluna '{col_cidade}' não encontrada. Filtro de cidade ignorado.\n")


        # --- FILTRO BACKLOG ---
        if backlog_filtro and str(backlog_filtro).strip():
            possiveis_nomes_backlog = ['Backlog', 'Backlog time(Station)']
            col_backlog_encontrada = None
            
            for nome in possiveis_nomes_backlog:
                if nome in df.columns:
                    col_backlog_encontrada = nome
                    break
            
            if col_backlog_encontrada:
                try:
                    val_alvo = int(backlog_filtro)
                    coluna_numerica = pd.to_numeric(df[col_backlog_encontrada], errors='coerce')
                    df_filtrado = df[coluna_numerica == val_alvo]
                    
                    if len(df_filtrado) == 0:
                        log_textbox.insert("end", f"⚠️ Nenhum registro com {col_backlog_encontrada} = {val_alvo}.\n")
                        return pd.DataFrame(), 0, "Filtro de Backlog resultou em 0."
                    else:
                        log_textbox.insert("end", f"ℹ️ Filtro Backlog ({val_alvo}) aplicado. {len(df_filtrado)} registros restantes.\n")
                    
                    df = df_filtrado
                except Exception as e:
                    log_textbox.insert("end", f"❌ Erro ao filtrar Backlog: {e}.\n")
            else:
                log_textbox.insert("end", f"⚠️ Coluna de Backlog não encontrada.\n")

        # --- RESULTADO FINAL ---
        if df.empty:
            log_textbox.insert("end", "⚠️ Atenção: 0 registros encontrados após filtros.\n")
            return pd.DataFrame(), 0, "Nenhum dado."

        msg_final = f"✅ Processamento iniciado com {len(df)} registros válidos."
        log_textbox.insert("end", f"{msg_final}\n")
        return df, len(df), msg_final

    except Exception as e:
        erro_msg = f"❌ Erro crítico ao ler arquivo: {e}"
        log_textbox.insert("end", f"{erro_msg}\n")
        return pd.DataFrame(), 0, erro_msg