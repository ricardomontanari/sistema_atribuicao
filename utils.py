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
VERSAO_SISTEMA = 'Beta 1.4 (Backlog Column Fix)'

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
            # Verifica se PARAR_AUTOMACAO é False antes de setar True para evitar logs duplicados
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
    Lê o Excel e aplica os filtros de Cidade e Backlog de forma segura.
    Agora verifica múltiplas variações de nome para a coluna Backlog.
    """
    try:
        # 1. Carregar Excel
        log_textbox.insert("end", f"Lendo arquivo '{arquivo}'...\n")
        
        # Lê o arquivo forçando string inicialmente para evitar erros de conversão automática
        df = pd.read_excel(arquivo, dtype=str)
        
        # CORREÇÃO CRÍTICA 1: Remove espaços dos nomes das colunas
        df.columns = df.columns.str.strip()
        
        # 2. Filtro de Cidade
        if cidade_filtro and cidade_filtro not in ["NENHUM FILTRO", "SELECIONE", ""]:
            # Tenta encontrar a coluna correta (Aceita 'Destination City' ou 'Cidade')
            col_cidade = 'Destination City' if 'Destination City' in df.columns else 'Cidade'
            
            if col_cidade in df.columns:
                # Normaliza a lista de filtros (Upper e sem espaços)
                lista_desejada = [c.strip().upper() for c in cidade_filtro.split(',') if c.strip()]
                
                # Normaliza a coluna do Excel e filtra
                df = df[df[col_cidade].astype(str).str.strip().str.upper().isin(lista_desejada)]
                
                log_textbox.insert("end", f"ℹ️ Filtro Cidade aplicado: {len(df)} registros restantes.\n")
            else:
                log_textbox.insert("end", f"⚠️ Coluna '{col_cidade}' não encontrada. Filtro de cidade ignorado.\n")

        # 3. Filtro de Backlog (ATUALIZADO PARA MÚLTIPLOS NOMES)
        if backlog_filtro and str(backlog_filtro).strip():
            # Lista de possíveis nomes para a coluna Backlog
            possiveis_nomes_backlog = ['Backlog', 'Backlog time(Station)']
            col_backlog_encontrada = None
            
            # Procura qual nome existe no DataFrame
            for nome in possiveis_nomes_backlog:
                if nome in df.columns:
                    col_backlog_encontrada = nome
                    break
            
            if col_backlog_encontrada:
                try:
                    val_alvo = int(backlog_filtro)
                    
                    # Converte a coluna encontrada para numérico de forma segura
                    coluna_numerica = pd.to_numeric(df[col_backlog_encontrada], errors='coerce')
                    
                    # Filtra comparando número com número
                    df_filtrado = df[coluna_numerica == val_alvo]
                    
                    if len(df_filtrado) == 0:
                        log_textbox.insert("end", f"⚠️ Nenhum registro encontrado com {col_backlog_encontrada} = {val_alvo}.\n")
                        # Debug: Mostra alguns valores únicos encontrados
                        valores_encontrados = df[col_backlog_encontrada].unique()[:5]
                        log_textbox.insert("end", f"   (Valores na coluna '{col_backlog_encontrada}': {valores_encontrados})\n")
                    else:
                        log_textbox.insert("end", f"ℹ️ Filtro Backlog ({val_alvo}) aplicado na coluna '{col_backlog_encontrada}'. {len(df_filtrado)} registros encontrados.\n")
                    
                    df = df_filtrado
                    
                except Exception as e:
                    log_textbox.insert("end", f"❌ Erro técnico ao filtrar Backlog: {e}. Filtro ignorado.\n")
            else:
                # Log detalhado caso nenhuma das colunas seja encontrada
                cols = ", ".join(list(df.columns))
                log_textbox.insert("end", f"⚠️ Coluna de Backlog não encontrada (busquei por: {possiveis_nomes_backlog}).\n   Colunas detectadas: {cols}\n")

        # 4. Resultado Final
        if df.empty:
            log_textbox.insert("end", "⚠️ Atenção: Os filtros resultaram em 0 registros para processar.\n")
            return pd.DataFrame(), 0, "Nenhum dado."

        msg_final = f"✅ Processamento iniciado com {len(df)} registros."
        log_textbox.insert("end", f"{msg_final}\n")
        return df, len(df), msg_final

    except Exception as e:
        erro_msg = f"❌ Erro crítico ao ler arquivo: {e}"
        log_textbox.insert("end", f"{erro_msg}\n")
        return pd.DataFrame(), 0, erro_msg