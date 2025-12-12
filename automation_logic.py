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
    "erro_devolucao.png",
    "erro_baixada.png"
]

# Cache para armazenar os objetos de imagem na memória RAM
CACHE_IMAGENS = []
CACHE_CARREGADO = False

def _log(msg):
    return f"[{time.strftime('%H:%M:%S')}] {msg}"

# ####################################################################
# --- DETECÇÃO VISUAL ---
# ####################################################################

def carregar_cache_imagens(log_textbox):
    """
    Carrega as imagens de erro na memória para detecção rápida.
    """
    global CACHE_IMAGENS, CACHE_CARREGADO
    if CACHE_CARREGADO: return

    log_textbox.insert("end", _log("Carregando imagens de erro...\n"))
    for nome_img in IMAGENS_DE_EXCECAO:
        path = None
        if os.path.exists(nome_img): path = os.path.abspath(nome_img)
        else:
            try: 
                temp = utils.resource_path(nome_img)
                if os.path.exists(temp): path = temp
            except: pass

        if path and os.path.exists(path):
            try:
                img_obj = Image.open(path)
                CACHE_IMAGENS.append((nome_img, img_obj))
            except: pass
        else:
            log_textbox.insert("end", _log(f"⚠️ Imagem não encontrada: '{nome_img}'\n"))
    
    CACHE_CARREGADO = True

def verificar_excecoes_visuais(log_textbox):
    """
    Verifica se alguma imagem de erro está na tela AGORA.
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
# --- CONTROLE DE PAUSA E RETOMADA ---
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

def verificar_e_esperar_limpeza_de_erro(log_textbox):
    """
    Verifica se o erro sumiu da tela. Se não sumiu em 5s, retorna False.
    """
    MAX_WAIT_TIME = 5 # 5 segundos de tolerância
    start_time = time.time()
    nome_erro_persistente = "Pop-up de Erro"
    
    while (time.time() - start_time) < MAX_WAIT_TIME:
        tem_erro, nome_erro = verificar_excecoes_visuais(log_textbox)
        if not tem_erro:
            log_textbox.insert("end", _log("✅ Tela limpa. Retomando segurança.\n"))
            return True # O erro SUMIU, pode seguir.
        
        # O erro persiste, atualiza o nome e espera um pouco
        nome_erro_persistente = nome_erro 
        time.sleep(0.5)
        
    log_textbox.insert("end", _log(f"❌ ERRO PERSISTE! A janela '{nome_erro_persistente}' não foi fechada.\n"))
    return False # O erro NÃO SUMIU após 5 segundos

def tratar_pausa_e_retomar_foco(log_textbox, safe_configure_buttons_cb, is_last_cycle=False):
    """
    TRAVA a execução, garante a limpeza da tela e re-estabelece o foco.
    Adiciona a lógica de finalização se for o último ciclo.
    """
    try: winsound.Beep(500, 500)
    except: pass
    
    log_textbox.insert("end", _log("⏸️ SISTEMA PAUSADO. Resolva o erro e clique em 'CONTINUAR'.\n"))
    log_textbox.see("end")
    
    safe_configure_buttons_cb(iniciar_state="disabled", continuar_state="normal")
    
    # LOOP DE ESPERA PELO CLIQUE DO USUÁRIO
    while utils.PARAR_AUTOMACAO:
        time.sleep(0.5)
    
    # --- USUÁRIO CLICOU EM CONTINUAR ---
    
    if is_last_cycle:
        # Se é o último ciclo e o usuário clicou em continuar após o erro, FINALIZA.
        # Não precisa de mais verificações de foco/limpeza se for terminar
        log_textbox.insert("end", _log("✅ Último item processado. Finalizando automação...\n"))
        safe_configure_buttons_cb(iniciar_state="normal", continuar_state="disabled")
        return True, True # Retorna True para retomada e True para finalização

    log_textbox.insert("end", _log("Aguardando confirmação de que o erro foi fechado e o foco...\n"))

    # 1. VERIFICA SE O ERRO FOI FECHADO
    if not verificar_e_esperar_limpeza_de_erro(log_textbox):
        # Se o erro PERSISTE, força a pausa de volta (Recursão)
        log_textbox.insert("end", _log("⚠️ O erro não foi fechado. Pausando novamente. Por favor, feche o popup.\n"))
        utils.PARAR_AUTOMACAO = True
        return tratar_pausa_e_retomar_foco(log_textbox, safe_configure_buttons_cb, is_last_cycle) # Volta para o modo pausado
    
    # 2. SE A TELA ESTIVER LIMPA, RECUPERA O FOCO DO NAVEGADOR
    focar_navegador = focar_janela_por_titulo("Opera", log_textbox) or \
                      focar_janela_por_titulo("Google Chrome", log_textbox) or \
                      focar_janela_por_titulo("Microsoft Edge", log_textbox)
    
    if not focar_navegador:
        log_textbox.insert("end", _log("❌ Foco no navegador perdido. Pausando novamente.\n"))
        utils.PARAR_AUTOMACAO = True
        return tratar_pausa_e_retomar_foco(log_textbox, safe_configure_buttons_cb, is_last_cycle) # Volta para o modo pausado

    # 3. VERIFICAÇÃO DE SAÍDA APÓS FOCO
    # Delay de 0.2s para renderização
    time.sleep(0.2) 
    tem_erro_final, nome_erro_final = verificar_excecoes_visuais(log_textbox)
    
    if tem_erro_final:
        log_textbox.insert("end", _log(f"❌ ERRO REAPARECEU após focar! '{nome_erro_final}'. Pausando novamente.\n"))
        utils.PARAR_AUTOMACAO = True
        return tratar_pausa_e_retomar_foco(log_textbox, safe_configure_buttons_cb, is_last_cycle) 


    # Se tudo sumiu e o foco foi recuperado, pode seguir
    try: winsound.Beep(1000, 300)
    except: pass
    
    log_textbox.insert("end", _log("▶️ Retomando automação...\n"))
    safe_configure_buttons_cb(iniciar_state="disabled", continuar_state="disabled")
    
    return True, False # Retorna True para retomada e False para não finalização

def lidar_com_erro_e_pausar(log_textbox, msg_erro, safe_update_gui_cb, safe_configure_buttons_cb, is_last_cycle=False):
    """
    Função central que ativa a pausa.
    """
    utils.PARAR_AUTOMACAO = True
    safe_update_gui_cb(status="Pausado")
    
    log_textbox.insert("end", _log(f"🛑 {msg_erro}\n"))
    log_textbox.see("end")
    
    # Passa o flag is_last_cycle para a função de tratamento de pausa
    retomada_ok, deve_finalizar = tratar_pausa_e_retomar_foco(log_textbox, safe_configure_buttons_cb, is_last_cycle)
    
    if deve_finalizar:
        return True # Sinaliza para o loop principal finalizar
    
    if retomada_ok:
        safe_update_gui_cb(status="Rodando")
        return True
    return False

# ####################################################################
# --- AÇÃO (CTRL+V / ENTER) ---
# ####################################################################

def executar_passos_ciclo(valor, delay):
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

        # 1. Leitura
        dados, repetir, msg = utils.ler_e_filtrar_dados(utils.NOME_ARQUIVO_ALVO, cidade_filtro, backlog_filtro, log_textbox)
        log_textbox.insert("end", _log(msg + "\n"))
        
        if repetir == 0:
            safe_update_gui_cb(status="Finalizado")
            return

        # 2. Foco Inicial
        focar_janela_por_titulo("Excel", log_textbox)
        # Foco Final no Navegador para começar a colar
        focar_janela_por_titulo("Opera", log_textbox) or focar_janela_por_titulo("Google Chrome", log_textbox) or focar_janela_por_titulo("Microsoft Edge", log_textbox)

        safe_update_gui_cb(status="Rodando", total_ciclos=repetir)
        
        # 3. Loop Principal
        for i in range(utils.INDICE_ATUAL_DO_CICLO, repetir):
            utils.INDICE_ATUAL_DO_CICLO = i
            linha = dados.iloc[i]
            
            # Checa se é o último ciclo
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
                
                # A) Checa Pausa Manual (ESC)
                if utils.PARAR_AUTOMACAO:
                    if lidar_com_erro_e_pausar(log_textbox, "Pausa Manual", safe_update_gui_cb, safe_configure_buttons_cb, is_last_cycle=is_last):
                        finalizar_automacao = True # Sinaliza para finalizar o for
                        break 
                    else: return

                # B) Executa: Colar + Enter
                if not executar_passos_ciclo(val, utils.DELAY_ATUAL):
                    if lidar_com_erro_e_pausar(log_textbox, "Falha no teclado", safe_update_gui_cb, safe_configure_buttons_cb, is_last_cycle=is_last):
                        continue
                    else: return

                # C) DETECÇÃO PÓS-AÇÃO (Radar de 0.5 segundos)
                inicio_radar = time.time()
                tem_erro = False
                nome_erro = None
                
                while (time.time() - inicio_radar) < 0.5:
                    tem_erro, nome_erro = verificar_excecoes_visuais(log_textbox)
                    if tem_erro:
                        break 
                    time.sleep(0.1)

                if tem_erro:
                    # --- NOVO AJUSTE AQUI ---
                    if is_last:
                        # Se é o ÚLTIMO ITEM e detectou erro, APENAS LOGA e FINALIZA, SEM PAUSAR.
                        log_textbox.insert("end", _log(f"🛑 Último item ({val}) com Erro Visual: {nome_erro}. Finalizando.\n"))
                        finalizar_automacao = True
                        break
                    
                    # ITEM NORMAL -> PAUSA
                    retomou_sucesso = lidar_com_erro_e_pausar(log_textbox, f"Erro Visual: {nome_erro}", safe_update_gui_cb, safe_configure_buttons_cb, is_last_cycle=is_last)
                    
                    if retomou_sucesso:
                        # Item normal: Pula para o próximo.
                        log_textbox.insert("end", _log("⚠️ Item com erro tratado. Pulando para o próximo.\n"))
                        break # Sai do WHILE True loop (item handled)
                    else: return # Falha na retomada, aborta tudo.

                # Se chegou aqui, não tem erro.
                break 
            
            if finalizar_automacao:
                break


        winsound.Beep(1000, 500)
        log_textbox.insert("end", _log("✅ Automação Finalizada!\n"))
        safe_update_gui_cb(status="Finalizado")

    except Exception as e:
        log_textbox.insert("end", _log(f"❌ ERRO FATAL: {e}\n"))
        safe_update_gui_cb(status="Erro")
        sucesso = False
        
    finally:
        # AQUI FOI AJUSTADO: Reseta o índice se a automação foi concluída com sucesso OU se finalizou no último ciclo.
        if (sucesso and utils.INDICE_ATUAL_DO_CICLO == repetir - 1) or (utils.INDICE_ATUAL_DO_CICLO == repetir - 1):
            utils.INDICE_ATUAL_DO_CICLO = 0
            
        safe_configure_buttons_cb(iniciar_state="normal", continuar_state="disabled")
        utils.PARAR_AUTOMACAO = False