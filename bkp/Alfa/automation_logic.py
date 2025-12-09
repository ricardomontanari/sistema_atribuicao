import pyautogui
import time
import pandas as pd     
import pyperclip
import winsound
import os
import utils
from PIL import Image # Import necessário para carregar imagens de forma segura

# ####################################################################
# --- FUNÇÕES DE FOCO E LEITURA DE DADOS ---
# ####################################################################

def focar_janela_por_titulo(titulo_parcial, log_textbox):
    """
    Tenta encontrar e ativar uma janela com um título que contém o 
    'titulo_parcial'. Retorna True em caso de sucesso, False caso contrário.
    """
    log_textbox.insert("end", f"  -> Buscando janela: '{titulo_parcial}'\n")
    log_textbox.see("end")
    time.sleep(0.1) 

    janelas_encontradas = pyautogui.getWindowsWithTitle(titulo_parcial)
    
    if janelas_encontradas:
        janela = janelas_encontradas[0]
        try:
            if janela.isMinimized:
                janela.restore()
            janela.activate()
            log_textbox.insert("end", f"  -> ✅ Foco obtido: '{janela.title}'\n")
            log_textbox.see("end")
            return True
        except Exception as e:
            log_textbox.insert("end", 
                               f"  -> ❌ ERRO ao ativar janela '{titulo_parcial}': {e}\n")
            log_textbox.see("end")
            return False
    else:
        log_textbox.insert("end", f"  -> ⚠️ Janela '{titulo_parcial}' não encontrada.\n")
        log_textbox.see("end")
        return False

def tratar_pausa_e_retomar_foco(log_textbox, safe_configure_buttons_cb):
    """
    Trata o estado de pausa (utils.PARAR_AUTOMACAO = True).
    Fica em loop até que a flag seja desligada pelo botão 'Continuar'.
    """
    # SOM DE PAUSA (Bipe Médio)
    try:
        winsound.Beep(500, 500)
    except: pass
    
    log_textbox.insert("end", "Aguardando ação do usuário (Pressione 'Continuar')...\n")
    
    # Habilita os botões para que o usuário possa interagir
    safe_configure_buttons_cb(iniciar_state="disabled", continuar_state="normal")
    
    # Loop de espera enquanto estiver pausado
    while utils.PARAR_AUTOMACAO:
        time.sleep(0.5) 
    
    # SOM DE RETOMADA (Bipe Agudo)
    try:
        winsound.Beep(1000, 500)
    except: pass
    
    log_textbox.insert("end", "\nRetomando a automação...\n")
    log_textbox.see("end")

    # Bloqueia botões novamente
    safe_configure_buttons_cb(iniciar_state="disabled", continuar_state="disabled")

    # Tenta obter o foco da janela novamente ao retomar (Prioridade: Opera)
    foco_recuperado = focar_janela_por_titulo("Opera", log_textbox) or \
                      focar_janela_por_titulo("Google Chrome", log_textbox) or \
                      focar_janela_por_titulo("Microsoft Edge", log_textbox)

    if not foco_recuperado:
        log_textbox.insert("end", "❌ ERRO: Foco perdido ao retomar a automação. Abortando.\n")
        log_textbox.see("end")
        return False 
        
    return True 

# ####################################################################
# --- FUNÇÃO DE VERIFICAÇÃO DE ERRO ---
# ####################################################################

def verificar_erro_e_pausar(log_textbox, timeout=0.1):
    """
    Verifica se alguma imagem de erro aparece na tela durante 'timeout' segundos.
    Agora configurado por padrão para 0.1 segundo.
    """
    imagens_erros = ["excessao_devolucao.png", "excessao_baixada.png"]
    
    start_time = time.time()
    
    # Feedback visual no log
    log_textbox.insert("end", "    -> 🔎 Verificando erros (0.1s)...\n")
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
                log_textbox.insert("end", f"⚠️ Falha ao carregar img {nome_img}: {e}\n")
    
    if not imagens_carregadas:
        # Se não achar as imagens, sai rápido para não atrasar
        return False

    # Loop que roda por X segundos (timeout) procurando a imagem
    while time.time() - start_time < timeout:
        
        for nome_arquivo, imagem_obj in imagens_carregadas:
            # TENTATIVA COM TOLERÂNCIA
            try:
                if pyautogui.locateOnScreen(imagem_obj, confidence=0.7, grayscale=True):
                    try: winsound.Beep(1000, 1000) 
                    except: pass
                    
                    log_textbox.insert("end", f"\n⚠️ ERRO IDENTIFICADO: {nome_arquivo}\n")
                    log_textbox.insert("end", "🛑 PAUSANDO AUTOMATICAMENTE...\n")
                    log_textbox.see("end")
                    utils.PARAR_AUTOMACAO = True
                    return True
            except Exception:
                # TENTATIVA EXATA (FALLBACK)
                try:
                    if pyautogui.locateOnScreen(imagem_obj):
                         try: winsound.Beep(1000, 1000) 
                         except: pass
                         log_textbox.insert("end", f"\n⚠️ ERRO IDENTIFICADO (Exato): {nome_arquivo}\n")
                         utils.PARAR_AUTOMACAO = True
                         return True
                except:
                     pass

        # Pausa mínima
        time.sleep(0.01)
    
    return False

# ####################################################################
# --- FUNÇÕES DE EXECUÇÃO DE AUTOMAÇÃO ---
# ####################################################################

def executar_passos_ciclo(valor_a_colar, log_textbox, delay):
    """
    Executa os passos de automação e verifica erros após o Enter.
    """
    try:
        # Passo 1: Copia o valor para o clipboard
        pyperclip.copy(valor_a_colar)
        time.sleep(0.1) 
        log_textbox.insert("end", f"    -> Copiado: '{valor_a_colar}'\n")

        # Passo 2: Focar no Navegador
        foco_navegador = focar_janela_por_titulo("Opera", log_textbox) or \
                         focar_janela_por_titulo("Google Chrome", log_textbox) or \
                         focar_janela_por_titulo("Microsoft Edge", log_textbox)
        
        if not foco_navegador:
            log_textbox.insert("end", "❌ ERRO: Navegador não encontrado para colar.\n")
            return False

        # Passo 3: Colar (Ctrl+V)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(delay) 
        
        # Passo 4: Enter
        pyautogui.press('enter')
        
        # --- VERIFICAÇÃO INTELIGENTE DE ERRO ---
        # Timeout reduzido para 0.1 segundo conforme solicitado
        erro_encontrado = verificar_erro_e_pausar(log_textbox, timeout=0.1)
        
        if erro_encontrado:
            return True 

        # Ajuste do delay residual
        # Como o tempo de verificação é muito curto (0.1s), aplicamos o delay normal
        # subtraindo apenas esse tempo se o delay total for considerável (> 1.0s)
        if delay > 1.0:
            time.sleep(delay - 0.1)
        
        return True
        
    except Exception as e:
        log_textbox.insert("end", f"❌ ERRO durante a execução do ciclo: {e}\n")
        log_textbox.see("end")
        return False

# ####################################################################
# --- FUNÇÃO PRINCIPAL (CORE) ---
# ####################################################################

def automacao_core(log_textbox, cidade_filtro, delay_inicial, safe_configure_buttons_cb, safe_update_gui_cb):
    """
    Lógica principal da automação.
    """
    sucesso = True
    TOTAL_DE_CICLOS = 0
    try:
        # 1. Leitura de Dados
        dados_filtrados, repetir, msg_log = utils.ler_e_filtrar_dados(
            utils.NOME_ARQUIVO_ALVO, cidade_filtro, log_textbox)
        
        log_textbox.insert("end", msg_log + "\n")
        log_textbox.see("end")

        if repetir == 0:
            log_textbox.insert("end", "⚠️ Nenhuma linha para processar após o filtro.\n")
            safe_update_gui_cb(status="Finalizado", total_ciclos=0, ciclo_atual=0)
            return 

        # 2. Foco Inicial das Janelas
        log_textbox.insert("end", "Iniciando foco nas janelas...\n")
        
        # Foca Excel
        if not focar_janela_por_titulo("Excel", log_textbox):
             log_textbox.insert("end", "⚠️ Aviso: Excel não encontrado.\n")

        # Foca Navegador
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
        
        log_textbox.insert("end", f"Total de ciclos: {repetir}\n")
        log_textbox.see("end")
        
        # FASE 2: LOOP DE REPETIÇÃO
        for i in range(utils.INDICE_ATUAL_DO_CICLO, repetir):
            
            # Atualiza índice global
            utils.INDICE_ATUAL_DO_CICLO = i
            
            # --- LEITURA DINÂMICA DO DELAY ---
            delay_atualizado = utils.DELAY_ATUAL
            # ---------------------------------
            
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
            
            log_textbox.insert("end", 
                               f"\n⚙️ Ciclo {i + 1}/{repetir}: {valor_a_colar}\n")
            log_textbox.see("end")
            
            # Atualiza GUI
            safe_update_gui_cb(ciclo_atual=i + 1) 

            # Executa passos com o DELAY ATUALIZADO
            if not executar_passos_ciclo(valor_a_colar, log_textbox, delay_atualizado):
                log_textbox.insert("end", 
                                   "❌ ERRO: Falha no ciclo de automação. Abortando.\n")
                sucesso = False
                return

            log_textbox.insert("end", "🎉 Ciclo concluído com sucesso!\n")
            log_textbox.see("end")

            # --- VERIFICAÇÃO DE PAUSA (Seja por ESC ou por Erro Visual) ---
            if utils.PARAR_AUTOMACAO:
                log_textbox.insert("end", "--- PAUSA ATIVA. TRANCANDO A THREAD. ---\n")
                log_textbox.see("end")
                safe_update_gui_cb(status="Pausado") 
                
                # Bloqueia a execução aqui até o usuário clicar em Continuar
                if not tratar_pausa_e_retomar_foco(log_textbox, safe_configure_buttons_cb):
                    return 

                safe_update_gui_cb(status="Rodando") 
            # ----------------------------
            
        log_textbox.insert("end", "\n✅ Automação de Cópia/Colagem finalizada.\n")
        log_textbox.see("end")

    except Exception as e:
        log_textbox.insert("end", f"\n❌ ERRO FATAL na automação: {e}\n")
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