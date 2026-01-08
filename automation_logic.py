import pyautogui
import time
import pandas as pd     
import pyperclip
import winsound
import os
import utils
# from PIL import Image <--- Não precisamos mais importar PIL aqui

# ####################################################################
# --- CONFIGURAÇÃO: IMAGENS DE ERRO ---
# ####################################################################

IMAGENS_DE_EXCECAO = [
    "erro_baixada.png"
]

# Cache agora armazena tuplas: (NomeArquivo, CaminhoAbsoluto)
CACHE_IMAGENS = []
CACHE_CARREGADO = False

def _log(msg):
    """Formata a mensagem de log com o horário atual."""
    return f"[{time.strftime('%H:%M:%S')}] {msg}"

# ####################################################################
# --- MÓDULO: DETECÇÃO VISUAL ---
# ####################################################################

def carregar_cache_imagens(log_textbox):
    """
    Resolve os caminhos absolutos das imagens e armazena no cache.
    """
    global CACHE_IMAGENS, CACHE_CARREGADO
    if CACHE_CARREGADO: return

    log_textbox.insert("end", _log("Mapeando imagens de erro...\n"))
    count = 0
    
    for nome_img in IMAGENS_DE_EXCECAO:
        # Usa utils.resource_path para pegar o caminho REAL (seja pasta local ou Temp do EXE)
        caminho_final = utils.resource_path(nome_img)
        
        if os.path.exists(caminho_final):
            # Armazena o CAMINHO (String), não o objeto imagem
            CACHE_IMAGENS.append((nome_img, caminho_final))
            count += 1
        else:
            log_textbox.insert("end", _log(f"⚠️ CRÍTICO: Imagem não encontrada: {caminho_final}\n"))
    
    if count > 0:
        log_textbox.insert("end", _log(f"✅ {count} imagens mapeadas com sucesso.\n"))
    
    CACHE_CARREGADO = True

def verificar_erro_visual_na_tela(log_textbox):
    """
    Verifica se alguma imagem de erro está visível passando o CAMINHO do arquivo.
    """
    if not CACHE_CARREGADO: carregar_cache_imagens(log_textbox)

    for nome_img, caminho_img in CACHE_IMAGENS:
        try:
            # Passamos o CAMINHO (String). O PyAutoGUI lê do disco.
            # O confidence=0.9 exige que o 'opencv-python' esteja no executável (ver build.py)
            if pyautogui.locateOnScreen(caminho_img, grayscale=True, confidence=0.9):
                return True, nome_img
        except Exception:
            # Fallback: Se o opencv falhar ou não estiver presente, tenta busca exata
            try:
                if pyautogui.locateOnScreen(caminho_img, grayscale=True):
                    return True, nome_img
            except: pass
    
    return False, None

# ####################################################################
# --- MÓDULO: CONTROLE DE JANELAS E PAUSA ---
# ####################################################################

def focar_janela_por_titulo(titulo_parcial, log_textbox):
    try:
        janelas = pyautogui.getWindowsWithTitle(titulo_parcial)
        if janelas:
            janela = janelas[0]
            if janela.isMinimized: janela.restore()
            janela.activate()
            time.sleep(0.2)
            return True
    except: pass
    return False

def garantir_foco_navegador(log_textbox):
    return (focar_janela_por_titulo("Opera", log_textbox) or 
            focar_janela_por_titulo("Google Chrome", log_textbox) or 
            focar_janela_por_titulo("Microsoft Edge", log_textbox))

def verificar_e_esperar_limpeza_de_erro(log_textbox):
    MAX_WAIT_TIME = 5.0
    start_time = time.time()
    nome_erro_persistente = "Pop-up"
    
    while (time.time() - start_time) < MAX_WAIT_TIME:
        tem_erro, nome_erro = verificar_erro_visual_na_tela(log_textbox)
        if not tem_erro:
            log_textbox.insert("end", _log("✅ Tela limpa confirmada.\n"))
            return True
        
        nome_erro_persistente = nome_erro 
        time.sleep(0.5)
        
    log_textbox.insert("end", _log(f"❌ ERRO PERSISTE! A janela '{nome_erro_persistente}' ainda está na tela.\n"))
    return False

def tratar_pausa_e_retomar_foco(log_textbox, safe_configure_buttons_cb, is_last_cycle=False):
    try: winsound.Beep(500, 500)
    except: pass
    
    log_textbox.insert("end", _log("⏸️ SISTEMA PAUSADO. Resolva o erro e clique em 'CONTINUAR' ou 'PARAR'.\n"))
    log_textbox.see("end")
    
    safe_configure_buttons_cb(iniciar_state="disabled", continuar_state="normal")
    
    while utils.PARAR_AUTOMACAO:
        if getattr(utils, 'CANCELAR_AUTOMACAO', False):
            focar_janela_por_titulo("Atribuidor", log_textbox) 
            log_textbox.insert("end", _log("⛔ Cancelamento solicitado pelo usuário.\n"))
            return False, True 

        time.sleep(0.5)
    
    if is_last_cycle:
        log_textbox.insert("end", _log("✅ Último item processado. Finalizando automação...\n"))
        safe_configure_buttons_cb(iniciar_state="normal", continuar_state="disabled")
        return True, True

    log_textbox.insert("end", _log("Verificando limpeza da tela e foco...\n"))

    if not verificar_e_esperar_limpeza_de_erro(log_textbox):
        log_textbox.insert("end", _log("⚠️ Pausando novamente. Por favor, feche o erro.\n"))
        utils.PARAR_AUTOMACAO = True
        return tratar_pausa_e_retomar_foco(log_textbox, safe_configure_buttons_cb, is_last_cycle)
    
    focar_janela_por_titulo("Atribuidor", log_textbox) 

    if not garantir_foco_navegador(log_textbox):
        log_textbox.insert("end", _log("❌ Foco perdido. Pausando novamente.\n"))
        utils.PARAR_AUTOMACAO = True
        return tratar_pausa_e_retomar_foco(log_textbox, safe_configure_buttons_cb, is_last_cycle)

    time.sleep(0.2) 
    tem_erro_final, nome_erro_final = verificar_erro_visual_na_tela(log_textbox)
    
    if tem_erro_final:
        log_textbox.insert("end", _log(f"❌ Erro '{nome_erro_final}' reapareceu após focar. Pausando.\n"))
        utils.PARAR_AUTOMACAO = True
        return tratar_pausa_e_retomar_foco(log_textbox, safe_configure_buttons_cb, is_last_cycle) 

    try: winsound.Beep(1000, 300)
    except: pass
    
    log_textbox.insert("end", _log("▶️ Retomando automação...\n"))
    safe_configure_buttons_cb(iniciar_state="disabled", continuar_state="disabled")
    
    return True, False 

def acionar_pausa_sistema(log_textbox, motivo, safe_update_gui_cb, safe_configure_buttons_cb, is_last_cycle=False):
    utils.PARAR_AUTOMACAO = True
    safe_update_gui_cb(status="Pausado")
    
    log_textbox.insert("end", _log(f"🛑 {motivo}\n"))
    log_textbox.see("end")
    
    retomada_ok, deve_finalizar = tratar_pausa_e_retomar_foco(log_textbox, safe_configure_buttons_cb, is_last_cycle)
    
    if deve_finalizar:
        return True 
    
    if retomada_ok:
        safe_update_gui_cb(status="Rodando")
        return True
    
    return False

# ####################################################################
# --- MÓDULO: AÇÃO (INPUT) ---
# ####################################################################

def executar_acao_colar_enter(valor, delay):
    try:
        pyperclip.copy(valor)
        time.sleep(0.05)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(delay)
        pyautogui.press('enter')
        time.sleep(delay)
        return True
    except: return False

# ####################################################################
# --- CORE DA AUTOMAÇÃO ---
# ####################################################################

def automacao_core(log_textbox, cidade_filtro, backlog_filtro, delay_inicial, safe_configure_buttons_cb, safe_update_gui_cb):
    sucesso = True
    try:
        carregar_cache_imagens(log_textbox)

        # 1. Leitura e Filtros
        dados, repetir, msg = utils.ler_e_filtrar_dados(utils.NOME_ARQUIVO_ALVO, cidade_filtro, backlog_filtro, log_textbox)
        log_textbox.insert("end", _log(msg + "\n"))
        
        if repetir == 0:
            safe_update_gui_cb(status="Finalizado")
            return

        focar_janela_por_titulo("Excel", log_textbox)
        garantir_foco_navegador(log_textbox)

        safe_update_gui_cb(status="Rodando", total_ciclos=repetir)
        
        # 2. Loop Principal
        for i in range(utils.INDICE_ATUAL_DO_CICLO, repetir):
            if getattr(utils, 'CANCELAR_AUTOMACAO', False):
                log_textbox.insert("end", _log("⛔ Automação CANCELADA pelo usuário.\n"))
                break

            utils.INDICE_ATUAL_DO_CICLO = i
            linha = dados.iloc[i]
            is_last = (i == repetir - 1)
            
            try:
                if 'Waybill No' in dados.columns: val = str(linha['Waybill No'])
                elif 'Motorista ID' in dados.columns: val = str(linha['Motorista ID'])
                else: val = str(linha.iloc[0])
            except: val = "???"

            log_textbox.insert("end", _log(f"⚙️ Ciclo {i+1}/{repetir}: {val}\n"))
            log_textbox.see("end")
            safe_update_gui_cb(ciclo_atual=i+1)

            finalizar_automacao = False
            
            # --- LOOP DE TENTATIVA (RETRY) ---
            while True:
                if getattr(utils, 'CANCELAR_AUTOMACAO', False):
                    finalizar_automacao = True; break

                # A) Pausa Manual
                if utils.PARAR_AUTOMACAO:
                    deve_parar = acionar_pausa_sistema(log_textbox, "Pausa Manual", safe_update_gui_cb, safe_configure_buttons_cb, is_last_cycle=is_last)
                    
                    if getattr(utils, 'CANCELAR_AUTOMACAO', False):
                        finalizar_automacao = True; break

                    if deve_parar:
                        if is_last:
                            finalizar_automacao = True; break
                        pass 
                    else: 
                        return

                # B) Ação
                if not executar_acao_colar_enter(val, utils.DELAY_ATUAL):
                    deve_parar = acionar_pausa_sistema(log_textbox, "Falha no teclado", safe_update_gui_cb, safe_configure_buttons_cb, is_last_cycle=is_last)
                    if getattr(utils, 'CANCELAR_AUTOMACAO', False):
                        finalizar_automacao = True; break
                    if deve_parar and is_last:
                        finalizar_automacao = True; break
                    continue

                # C) Radar de Erro
                inicio_radar = time.time()
                tem_erro = False
                nome_erro = None
                
                while (time.time() - inicio_radar) < 0.5:
                    tem_erro, nome_erro = verificar_erro_visual_na_tela(log_textbox)
                    if tem_erro: break 
                    time.sleep(0.1)

                # D) Tratamento de Erro
                if tem_erro:
                    if is_last:
                        log_textbox.insert("end", _log(f"🛑 Último item ({val}) com Erro: {nome_erro}. Finalizando.\n"))
                        finalizar_automacao = True
                        break
                    
                    deve_parar = acionar_pausa_sistema(log_textbox, f"Erro Visual: {nome_erro}", safe_update_gui_cb, safe_configure_buttons_cb, is_last_cycle=is_last)
                    
                    if getattr(utils, 'CANCELAR_AUTOMACAO', False):
                        finalizar_automacao = True; break

                    log_textbox.insert("end", _log("⚠️ Item com erro tratado. Pulando para o próximo.\n"))
                    break 

                break 
            
            if finalizar_automacao:
                break

        if not getattr(utils, 'CANCELAR_AUTOMACAO', False):
            winsound.Beep(1000, 500)
            log_textbox.insert("end", _log("✅ Automação Finalizada!\n"))
            safe_update_gui_cb(status="Finalizado")
        else:
            log_textbox.insert("end", _log("⛔ Automação abortada.\n"))
            safe_update_gui_cb(status="Parado")

    except Exception as e:
        log_textbox.insert("end", _log(f"❌ ERRO FATAL: {e}\n"))
        safe_update_gui_cb(status="Erro")
        sucesso = False
        
    finally:
        focar_janela_por_titulo("Atribuidor", log_textbox) 
        utils.CANCELAR_AUTOMACAO = False 
        utils.PARAR_AUTOMACAO = False
        utils.INDICE_ATUAL_DO_CICLO = 0
        safe_configure_buttons_cb(iniciar_state="normal", continuar_state="disabled")