import pyautogui
import time
import pandas as pd     
import pyperclip
import winsound
import os
import utils
from PIL import Image

# ####################################################################
# --- FUNÇÕES DE FOCO E LEITURA DE DADOS ---
# ####################################################################

def focar_janela_por_titulo(titulo_parcial, log_textbox):
    """
    Tenta encontrar e ativar uma janela com um título que contém o 
    'titulo_parcial' por até 3 vezes (Foco Robusto). Retorna True em caso de sucesso.
    """
    MAX_TRIES = 3
    
    for attempt in range(MAX_TRIES):
        if attempt > 0:
            # NOVO: Usa log_time
            log_textbox.insert("end", utils.log_time(f"  -> Tentativa de foco {attempt+1}/{MAX_TRIES}...\n"))
            log_textbox.see("end")
            time.sleep(0.5) 

        janelas_encontradas = pyautogui.getWindowsWithTitle(titulo_parcial)
        
        if janelas_encontradas:
            janela = janelas_encontradas[0]
            try:
                if janela.isMinimized:
                    janela.restore()
                janela.activate()
                # NOVO: Usa log_time
                log_textbox.insert("end", utils.log_time(f"  -> ✅ Foco obtido: '{janela.title}'\n"))
                log_textbox.see("end")
                return True
            except Exception as e:
                # NOVO: Usa log_time
                log_textbox.insert("end", 
                                   utils.log_time(f"  -> ❌ ERRO ao ativar janela '{titulo_parcial}': {e}\n"))
                log_textbox.see("end")
                return False
        else:
            if attempt == MAX_TRIES - 1:
                # NOVO: Usa log_time
                log_textbox.insert("end", utils.log_time(f"  -> ⚠️ Janela '{titulo_parcial}' não encontrada após {MAX_TRIES} tentativas.\n"))
                log_textbox.see("end")
                return False
    
    return False

def tratar_pausa_e_retomar_foco(log_textbox, safe_configure_buttons_cb):
    """
    Trata o estado de pausa (utils.PARAR_AUTOMACAO = True).
    Recupera o foco do navegador com a nova lógica robusta.
    """
    try: winsound.Beep(500, 500)
    except: pass
    
    # NOVO: Usa log_time
    log_textbox.insert("end", utils.log_time("Aguardando ação do usuário (Pressione 'Continuar')...\n"))
    
    safe_configure_buttons_cb(iniciar_state="disabled", continuar_state="normal")
    
    while utils.PARAR_AUTOMACAO:
        time.sleep(0.5) 
    
    try: winsound.Beep(1000, 500)
    except: pass
    
    # NOVO: Usa log_time
    log_textbox.insert("end", utils.log_time("\nRetomando a automação...\n"))
    log_textbox.see("end")

    safe_configure_buttons_cb(iniciar_state="disabled", continuar_state="disabled")

    # Tenta obter o foco da janela novamente ao retomar (AGORA COM 3 TENTATIVAS)
    foco_recuperado = focar_janela_por_titulo("Opera", log_textbox) or \
                      focar_janela_por_titulo("Google Chrome", log_textbox) or \
                      focar_janela_por_titulo("Microsoft Edge", log_textbox)

    if not foco_recuperado:
        # NOVO: Usa log_time
        log_textbox.insert("end", utils.log_time("❌ ERRO: Foco perdido ao retomar a automação. Abortando.\n"))
        log_textbox.see("end")
        return False 
        
    return True 

# ####################################################################
# --- FUNÇÃO DE VERIFICAÇÃO DE ERRO ---
# ####################################################################

def verificar_erro_e_pausar(log_textbox, timeout=0.1):
    """
    Verifica se alguma imagem de erro aparece na tela durante 'timeout' segundos.
    """
    imagens_erros = ["excessao_devolucao.png", "excessao_baixada.png"]
    
    start_time = time.time()
    
    # Feedback visual no log
    # NOVO: Usa log_time
    log_textbox.insert("end", utils.log_time("    -> 🔎 Verificando erros (0.1s)...\n"))
    log_textbox.see("end")
    
    # Pré-carrega as imagens para memória
    imagens_carregadas = []
    for nome_img in imagens_erros:
        path = utils.resource_path(nome_img)
        if os.path.exists(path):
            try:
                img_obj = Image.open(path)
                imagens_carregadas.append((nome_img, img_obj))
            except Exception as e:
                # NOVO: Usa log_time
                log_textbox.insert("end", utils.log_time(f"⚠️ Falha ao carregar img {nome_img}: {e}\n"))
    
    if not imagens_carregadas:
        return False

    # Loop que roda por X segundos (timeout) procurando a imagem
    while time.time() - start_time < timeout:
        
        for nome_arquivo, imagem_obj in imagens_carregadas:
            # TENTATIVA COM TOLERÂNCIA
            try:
                if pyautogui.locateOnScreen(imagem_obj, confidence=0.7, grayscale=True):
                    try: winsound.Beep(1000, 1000) 
                    except: pass
                    
                    # NOVO: Usa log_time
                    log_textbox.insert("end", utils.log_time(f"\n⚠️ ERRO IDENTIFICADO: {nome_arquivo}\n"))
                    # NOVO: Usa log_time
                    log_textbox.insert("end", utils.log_time("🛑 PAUSANDO AUTOMATICAMENTE...\n"))
                    log_textbox.see("end")
                    utils.PARAR_AUTOMACAO = True
                    return True
            except Exception:
                # TENTATIVA EXATA (FALLBACK)
                try:
                    if pyautogui.locateOnScreen(imagem_obj):
                         try: winsound.Beep(1000, 1000) 
                         except: pass
                         # NOVO: Usa log_time
                         log_textbox.insert("end", utils.log_time(f"\n⚠️ ERRO IDENTIFICADO (Exato): {nome_arquivo}\n"))
                         utils.PARAR_AUTOMACAO = True
                         return True
                except:
                     pass

        time.sleep(0.01)
    
    return False

# ####################################################################
# --- FUNÇÕES DE EXECUÇÃO DE AUTOMAÇÃO ---
# ####################################################################

def executar_passos_ciclo(valor_a_colar, log_textbox):
    """
    Executa os passos de automação e verifica erros após o Enter.
    Usa o utils.DELAY_ATUAL (Tempo Dinâmico).
    """
    try:
        # Passo 1: Copia o valor para o clipboard
        pyperclip.copy(valor_a_colar)
        time.sleep(0.1) 
        # NOVO: Usa log_time
        log_textbox.insert("end", utils.log_time(f"    -> Copiado: '{valor_a_colar}'\n"))

        # Passo 2: Focar no Navegador (AGORA MAIS ROBUSTO)
        foco_navegador = focar_janela_por_titulo("Opera", log_textbox) or \
                         focar_janela_por_titulo("Google Chrome", log_textbox) or \
                         focar_janela_por_titulo("Microsoft Edge", log_textbox)
        
        if not foco_navegador:
            # NOVO: Usa log_time
            log_textbox.insert("end", utils.log_time("❌ ERRO: Navegador não encontrado para colar.\n"))
            return False
            
        # OBTÉM O DELAY ATUALIZADO
        delay = utils.DELAY_ATUAL 

        # Passo 3: Colar (Ctrl+V)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(delay) 
        
        # Passo 4: Enter
        pyautogui.press('enter')
        
        # --- VERIFICAÇÃO INTELIGENTE DE ERRO ---
        erro_encontrado = verificar_erro_e_pausar(log_textbox, timeout=0.1)
        
        if erro_encontrado:
            return True 

        # Ajuste do delay residual
        if delay > 1.0:
            time.sleep(delay - 0.1)
        
        return True
        
    except Exception as e:
        # NOVO: Usa log_time
        log_textbox.insert("end", utils.log_time(f"❌ ERRO durante a execução do ciclo: {e}\n"))
        log_textbox.see("end")
        return False

# ####################################################################
# --- FUNÇÃO PRINCIPAL (CORE) ---
# ####################################################################

def automacao_core(log_textbox, cidade_filtro, backlog_filtro, delay_inicial, safe_configure_buttons_cb, safe_update_gui_cb):
    """
    Lógica principal da automação.
    """
    sucesso = True
    TOTAL_DE_CICLOS = 0
    try:
        # 1. Leitura de Dados (Repassa o backlog_filtro para o utils)
        # O utils.ler_e_filtrar_dados já retorna a mensagem com timestamp
        dados_filtrados, repetir, msg_log = utils.ler_e_filtrar_dados(
            utils.NOME_ARQUIVO_ALVO, cidade_filtro, backlog_filtro, log_textbox)
        
        log_textbox.insert("end", msg_log + "\n")
        log_textbox.see("end")

        if repetir == 0:
            # NOVO: Usa log_time
            log_textbox.insert("end", utils.log_time("⚠️ Nenhuma linha para processar após o filtro.\n"))
            safe_update_gui_cb(status="Finalizado", total_ciclos=0, ciclo_atual=0)
            return 

        # 2. Foco Inicial das Janelas
        # NOVO: Usa log_time
        log_textbox.insert("end", utils.log_time("Iniciando foco nas janelas...\n"))
        
        # Foca Excel (AGORA MAIS ROBUSTO)
        if not focar_janela_por_titulo("Excel", log_textbox):
             # NOVO: Usa log_time
             log_textbox.insert("end", utils.log_time("⚠️ Aviso: Excel não encontrado.\n"))

        # Foca Navegador (AGORA MAIS ROBUSTO)
        foco_navegador = focar_janela_por_titulo("Opera", log_textbox) or \
                         focar_janela_por_titulo("Google Chrome", log_textbox) or \
                         focar_janela_por_titulo("Microsoft Edge", log_textbox)

        if not foco_navegador:
            raise Exception("Nenhum navegador compatível encontrado (Opera, Chrome, Edge).")
        
        # --- ATUALIZAÇÃO DA GUI (Início do loop) ---
        TOTAL_DE_CICLOS = repetir
        safe_update_gui_cb(status="Rodando", 
                           total_ciclos=TOTAL_DE_CICLOS, 
                           ciclo_atual=utils.INDICE_ATUAL_DO_CICLO)
        
        # NOVO: Usa log_time
        log_textbox.insert("end", utils.log_time(f"Total de ciclos: {repetir}\n"))
        log_textbox.see("end")
        
        # FASE 2: LOOP DE REPETIÇÃO
        for i in range(utils.INDICE_ATUAL_DO_CICLO, repetir):
            
            # Atualiza índice global
            utils.INDICE_ATUAL_DO_CICLO = i
            
            # Obtém dados da linha
            linha = dados_filtrados.iloc[i]
            try:
                if 'Waybill No' in dados_filtrados.columns:
                     valor_a_colar = str(linha['Waybill No'])
                elif 'Motorista ID' in dados_filtrados.columns:
                     valor_a_colar = str(linha['Motorista ID'])
                else:
                     valor_a_colar = str(linha.iloc[0]) 
            except:
                valor_a_colar = "DADO_DESCONHECIDO"
            
            # NOVO: Usa log_time
            log_textbox.insert("end", 
                               utils.log_time(f"\n⚙️ Ciclo {i + 1}/{repetir}: {valor_a_colar}\n"))
            log_textbox.see("end")
            
            # Atualiza GUI
            safe_update_gui_cb(ciclo_atual=i + 1) 

            # Executa passos - REMOVE O ARGUMENTO 'delay_atualizado'
            if not executar_passos_ciclo(valor_a_colar, log_textbox):
                # NOVO: Usa log_time
                log_textbox.insert("end", 
                                   utils.log_time("❌ ERRO: Falha no ciclo de automação. Abortando.\n"))
                sucesso = False
                return

            # NOVO: Usa log_time
            log_textbox.insert("end", utils.log_time("🎉 Ciclo concluído com sucesso!\n"))
            log_textbox.see("end")

            # --- VERIFICAÇÃO DE PAUSA (Seja por ESC ou por Erro Visual) ---
            if utils.PARAR_AUTOMACAO:
                # NOVO: Usa log_time
                log_textbox.insert("end", utils.log_time("--- PAUSA ATIVA. TRANCANDO A THREAD. ---\n"))
                log_textbox.see("end")
                safe_update_gui_cb(status="Pausado") 
                
                # Bloqueia a execução aqui até o usuário clicar em Continuar
                if not tratar_pausa_e_retomar_foco(log_textbox, safe_configure_buttons_cb):
                    return 

                safe_update_gui_cb(status="Rodando") 
            # ----------------------------
            
        # NOVO: Usa log_time
        log_textbox.insert("end", utils.log_time("\n✅ Automação de Cópia/Colagem finalizada.\n"))
        log_textbox.see("end")

    except Exception as e:
        # NOVO: Usa log_time
        log_textbox.insert("end", utils.log_time(f"\n❌ ERRO FATAL na automação: {e}\n"))
        log_textbox.see("end")
        safe_update_gui_cb(status="Erro") 
        sucesso = False
        
    finally:
        # Finalização e limpeza
        if sucesso and utils.INDICE_ATUAL_DO_CICLO == repetir - 1:
            utils.INDICE_ATUAL_DO_CICLO = 0 
            if repetir > 0:
                safe_update_gui_cb(status="Finalizado") 
        
        if not utils.PARAR_AUTOMACAO:
            safe_configure_buttons_cb(iniciar_state="normal", 
                                     continuar_state="disabled")