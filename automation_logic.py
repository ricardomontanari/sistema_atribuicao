import pyautogui
import time
import pandas as pd     
import pyperclip
import winsound
import os
import utils
from PIL import Image

# ####################################################################
# --- CONFIGURAÇÃO: IMAGENS DE ERRO ---
# ####################################################################

IMAGENS_DE_EXCECAO = [
    "erro_baixada.png"
]

# Cache para armazenar os objetos de imagem na memória RAM
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
    Carrega as imagens de erro do disco para a memória RAM.
    """
    global CACHE_IMAGENS, CACHE_CARREGADO
    if CACHE_CARREGADO: return

    log_textbox.insert("end", _log("Carregando banco de imagens de erro...\n"))
    count = 0
    
    for nome_img in IMAGENS_DE_EXCECAO:
        
        path = utils.resource_path(nome_img)

        if os.path.exists(path):
            try:
                img_obj = Image.open(path)
                # O .load() garante que a imagem vá para a RAM e não dependa do arquivo aberto
                img_obj.load() 
                CACHE_IMAGENS.append((nome_img, img_obj))
                count += 1
            except Exception as e:
                log_textbox.insert("end", _log(f"❌ Erro ao ler imagem: {e}\n"))
        else:
            # Este log vai te dizer exatamente onde o EXE está procurando a imagem
            log_textbox.insert("end", _log(f"⚠️ CRÍTICO: Imagem não encontrada em: {path}\n"))
    
    if count > 0:
        log_textbox.insert("end", _log(f"✅ {count} imagens carregadas com sucesso.\n"))
    CACHE_CARREGADO = True

def verificar_erro_visual_na_tela(log_textbox):
    """
    Verifica instantaneamente se alguma imagem de erro está visível na tela.
    Retorna (True, Nome da Imagem) se encontrar.
    """
    if not CACHE_CARREGADO: carregar_cache_imagens(log_textbox)

    for nome_img, img_obj in CACHE_IMAGENS:
        try:
            # Tenta com confiança 0.8
            if pyautogui.locateOnScreen(img_obj, grayscale=True, confidence=0.8):
                return True, nome_img
        except Exception:
            # Fallback
            try:
                if pyautogui.locateOnScreen(img_obj, grayscale=True):
                    return True, nome_img
            except: pass
    
    return False, None

# ####################################################################
# --- MÓDULO: CONTROLE DE JANELAS E PAUSA ---
# ####################################################################

def focar_janela_por_titulo(titulo_parcial, log_textbox):
    """
    Tenta trazer para frente uma janela que contenha o título parcial.
    Esta versão é silenciosa no log para evitar poluição em ações rápidas.
    """
    try:
        janelas = pyautogui.getWindowsWithTitle(titulo_parcial)
        if janelas:
            janela = janelas[0]
            if janela.isMinimized: janela.restore()
            janela.activate()
            time.sleep(0.2) # Delay para renderização da janela
            return True
    except: pass
    return False

def garantir_foco_navegador(log_textbox):
    """Tenta focar em um dos navegadores suportados."""
    return (focar_janela_por_titulo("Opera", log_textbox) or 
            focar_janela_por_titulo("Google Chrome", log_textbox) or 
            focar_janela_por_titulo("Microsoft Edge", log_textbox))

def verificar_e_esperar_limpeza_de_erro(log_textbox):
    """
    Verifica se o erro sumiu da tela antes de permitir a retomada.
    """
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
    """
    Bloqueia a execução enquanto o sistema está pausado.
    Retorna:
      (True, False) -> Retomar e Continuar
      (True, True)  -> Retomar e Finalizar (Último ciclo)
      (False, True) -> ABORTAR/CANCELAR AUTOMACAO
    """
    try: winsound.Beep(500, 500)
    except: pass
    
    log_textbox.insert("end", _log("⏸️ SISTEMA PAUSADO. Resolva o erro e clique em 'CONTINUAR' ou 'PARAR'.\n"))
    log_textbox.see("end")
    
    # Sinaliza para a GUI transformar INICIAR em PARAR
    safe_configure_buttons_cb(iniciar_state="disabled", continuar_state="normal")
    
    # === LOOP DE BLOQUEIO ===
    while utils.PARAR_AUTOMACAO:
        # Se uma flag de cancelamento for ativada externamente (ex: Botão Parar)
        if getattr(utils, 'CANCELAR_AUTOMACAO', False):
            # AÇÃO DE FOCO NO CANCELAMENTO (Foco é dado aqui)
            focar_janela_por_titulo("Atribuidor", log_textbox) 
            log_textbox.insert("end", _log("⛔ Cancelamento solicitado pelo usuário.\n"))
            return False, True # Abortar

        time.sleep(0.5)
    
    # --- USUÁRIO CLICOU EM CONTINUAR ---
    
    # Caso Especial: Erro no último ciclo -> Finaliza direto
    if is_last_cycle:
        log_textbox.insert("end", _log("✅ Último item processado. Finalizando automação...\n"))
        safe_configure_buttons_cb(iniciar_state="normal", continuar_state="disabled")
        return True, True # (Retomada OK, Deve Finalizar)

    log_textbox.insert("end", _log("Verificando limpeza da tela e foco...\n"))

    if not verificar_e_esperar_limpeza_de_erro(log_textbox):
        log_textbox.insert("end", _log("⚠️ Pausando novamente. Por favor, feche o erro.\n"))
        utils.PARAR_AUTOMACAO = True
        return tratar_pausa_e_retomar_foco(log_textbox, safe_configure_buttons_cb, is_last_cycle)
    
    # DEVOLVE O FOCO À APLICAÇÃO ANTES DE CONTINUAR A VERIFICAÇÃO DE FOCO DE JANELA
    # O foco é restaurado aqui (após clicar em CONTINUAR)
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

    # Sucesso na retomada
    try: winsound.Beep(1000, 300)
    except: pass
    
    log_textbox.insert("end", _log("▶️ Retomando automação...\n"))
    safe_configure_buttons_cb(iniciar_state="disabled", continuar_state="disabled")
    
    return True, False 

def acionar_pausa_sistema(log_textbox, motivo, safe_update_gui_cb, safe_configure_buttons_cb, is_last_cycle=False):
    """
    Função wrapper que ativa a flag global de pausa e chama o tratador.
    """
    utils.PARAR_AUTOMACAO = True
    safe_update_gui_cb(status="Pausado")
    
    log_textbox.insert("end", _log(f"🛑 {motivo}\n"))
    log_textbox.see("end")
    
    # Captura o retorno (Retomada OK, Deve Finalizar/Abortar)
    retomada_ok, deve_finalizar = tratar_pausa_e_retomar_foco(log_textbox, safe_configure_buttons_cb, is_last_cycle)
    
    if deve_finalizar:
        # Se foi abortado (retomada_ok=False) ou finalizado normal (retomada_ok=True e is_last)
        return True 
    
    if retomada_ok:
        safe_update_gui_cb(status="Rodando")
        return True
    
    # Se chegou aqui, algo falhou gravemente ou foi cancelado
    return False

# ####################################################################
# --- MÓDULO: AÇÃO (INPUT) ---
# ####################################################################

def executar_acao_colar_enter(valor, delay):
    """Executa a sequência: Copiar -> Colar (Ctrl+V) -> Enter."""
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

        # 2. Preparação de Janelas
        focar_janela_por_titulo("Excel", log_textbox)
        garantir_foco_navegador(log_textbox)

        safe_update_gui_cb(status="Rodando", total_ciclos=repetir)
        
        # 3. Loop Principal
        for i in range(utils.INDICE_ATUAL_DO_CICLO, repetir):
            # Verifica cancelamento global antes de começar o ciclo
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

            # === LOOP DE BLINDAGEM (RETRY) ===
            finalizar_automacao = False
            
            while True:
                # Checa cancelamento dentro do loop de retry
                if getattr(utils, 'CANCELAR_AUTOMACAO', False):
                    finalizar_automacao = True; break

                # A) Pausa Manual (ESC)
                if utils.PARAR_AUTOMACAO:
                    # acionar_pausa_sistema agora retorna True se deve parar (cancelar ou finalizar)
                    deve_parar = acionar_pausa_sistema(log_textbox, "Pausa Manual", safe_update_gui_cb, safe_configure_buttons_cb, is_last_cycle=is_last)
                    
                    if getattr(utils, 'CANCELAR_AUTOMACAO', False):
                        finalizar_automacao = True; break

                    if deve_parar:
                        if is_last:
                            finalizar_automacao = True
                            break
                        # Se não foi cancelado e não é o último, apenas CONTINUA (Retomada)
                        pass 
                    else: 
                        # Erro na retomada
                        return

                # B) Ação: Colar + Enter
                if not executar_acao_colar_enter(val, utils.DELAY_ATUAL):
                    deve_parar = acionar_pausa_sistema(log_textbox, "Falha no teclado", safe_update_gui_cb, safe_configure_buttons_cb, is_last_cycle=is_last)
                    if getattr(utils, 'CANCELAR_AUTOMACAO', False):
                        finalizar_automacao = True; break
                    if deve_parar and is_last:
                        finalizar_automacao = True; break
                    continue

                # C) Radar de Erro (0.5 segundos)
                inicio_radar = time.time()
                tem_erro = False
                nome_erro = None
                
                while (time.time() - inicio_radar) < 0.5:
                    tem_erro, nome_erro = verificar_erro_visual_na_tela(log_textbox)
                    if tem_erro: break 
                    time.sleep(0.1)

                # D) Tratamento se houve erro (VISUAL)
                if tem_erro:
                    if is_last:
                        log_textbox.insert("end", _log(f"🛑 Último item ({val}) com Erro: {nome_erro}. Finalizando.\n"))
                        finalizar_automacao = True
                        break
                    
                    deve_parar = acionar_pausa_sistema(log_textbox, f"Erro Visual: {nome_erro}", safe_update_gui_cb, safe_configure_buttons_cb, is_last_cycle=is_last)
                    
                    if getattr(utils, 'CANCELAR_AUTOMACAO', False):
                        finalizar_automacao = True; break

                    # Se retomou com sucesso de um erro visual, PULA para o próximo (conforme regra)
                    log_textbox.insert("end", _log("⚠️ Item com erro tratado. Pulando para o próximo.\n"))
                    break 

                # Se chegou aqui, não houve erro e ação foi feita -> Sucesso do ciclo
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
        # NOVO: Garante que o foco volte para o Atribuidor ao Finalizar/Abortar
        focar_janela_por_titulo("Atribuidor", log_textbox) 
        
        # Reset total do estado
        utils.CANCELAR_AUTOMACAO = False 
        utils.PARAR_AUTOMACAO = False
        
        # Se terminou com sucesso ou foi cancelado, reseta o índice para 0
        utils.INDICE_ATUAL_DO_CICLO = 0
            
        safe_configure_buttons_cb(iniciar_state="normal", continuar_state="disabled")