import pyautogui
import time
import pandas as pd     
import pyperclip
import winsound
import os
import utils
from PIL import Image

# ####################################################################
# --- CONFIGURAÇÃO: IMAGENS QUE CAUSAM PAUSA ---
# ####################################################################

IMAGENS_DE_EXCECAO = [
    "erro_devolucao.png",
    "erro_baixada.png"
]

# Cache para armazenar as imagens na memória RAM (Máxima velocidade)
CACHE_IMAGENS = []
CACHE_CARREGADO = False

# Função auxiliar para log com horário (evita erros se utils não tiver log_time)
def _log(msg):
    return f"[{time.strftime('%H:%M:%S')}] {msg}"

# ####################################################################
# --- FUNÇÕES DE VERIFICAÇÃO VISUAL (CACHE + INSTANTÂNEA) ---
# ####################################################################

def carregar_cache_imagens(log_textbox):
    """
    Carrega imagens do disco para a memória uma única vez no início.
    """
    global CACHE_IMAGENS, CACHE_CARREGADO
    if CACHE_CARREGADO: return

    log_textbox.insert("end", _log("Carregando banco de imagens...\n"))
    for nome_img in IMAGENS_DE_EXCECAO:
        # Resolve o caminho do arquivo (seja script ou executável)
        if os.path.exists(nome_img): path = nome_img
        else:
            try: path = utils.resource_path(nome_img)
            except: continue

        if os.path.exists(path):
            try:
                img_obj = Image.open(path)
                CACHE_IMAGENS.append((nome_img, img_obj))
            except: pass
    
    CACHE_CARREGADO = True

def verificar_excecoes_visuais(log_textbox):
    """
    Varredura IMEDIATA da tela.
    Não espera tempo fixo. Retorna assim que encontra o primeiro erro.
    Isso garante que o sistema rode rápido (0.05s) quando não há erros.
    """
    if not CACHE_CARREGADO: carregar_cache_imagens(log_textbox)

    for nome_img, img_obj in CACHE_IMAGENS:
        try:
            # Tenta localizar com confiança 0.8 (Detecta mesmo com leve transparência/cor)
            if pyautogui.locateOnScreen(img_obj, grayscale=True, confidence=0.8):
                log_textbox.insert("end", _log(f"👁️ Erro visual detectado: '{nome_img}'\n"))
                return True, nome_img
        except Exception:
            # Fallback para busca exata (caso a biblioteca opencv não esteja carregada)
            try:
                if pyautogui.locateOnScreen(img_obj, grayscale=True):
                    log_textbox.insert("end", _log(f"👁️ Erro visual (Exato): '{nome_img}'\n"))
                    return True, nome_img
            except: pass
    
    return False, None

# ####################################################################
# --- FUNÇÕES DE CONTROLE E PAUSA ---
# ####################################################################

def focar_janela_por_titulo(titulo_parcial, log_textbox):
    """Tenta focar na janela com 3 tentativas rápidas."""
    MAX_TRIES = 3
    for attempt in range(MAX_TRIES):
        janelas = pyautogui.getWindowsWithTitle(titulo_parcial)
        if janelas:
            try:
                janela = janelas[0]
                if janela.isMinimized: janela.restore()
                janela.activate()
                time.sleep(0.2) 
                return True
            except: pass
        time.sleep(0.2)
    log_textbox.insert("end", _log(f"⚠️ Janela '{titulo_parcial}' não encontrada.\n"))
    return False

def tratar_pausa_e_retomar_foco(log_textbox, safe_configure_buttons_cb):
    """
    TRAVA a execução até o usuário clicar em CONTINUAR na interface.
    """
    try: winsound.Beep(500, 500)
    except: pass
    
    log_textbox.insert("end", _log("⏸️ PAUSADO. Resolva o erro e clique em 'CONTINUAR'.\n"))
    log_textbox.see("end")
    
    # Habilita o botão 'Continuar'
    safe_configure_buttons_cb(iniciar_state="disabled", continuar_state="normal")
    
    # --- LOOP DE BLOQUEIO ---
    # O código fica preso aqui eternamente até utils.PARAR_AUTOMACAO virar False
    while utils.PARAR_AUTOMACAO:
        time.sleep(0.5)
    
    # Usuário clicou em continuar
    try: winsound.Beep(1000, 300)
    except: pass
    
    log_textbox.insert("end", _log("▶️ Retomando...\n"))
    safe_configure_buttons_cb(iniciar_state="disabled", continuar_state="disabled")

    # Tenta devolver o foco para o navegador
    recuperou = focar_janela_por_titulo("Opera", log_textbox) or \
                focar_janela_por_titulo("Google Chrome", log_textbox) or \
                focar_janela_por_titulo("Microsoft Edge", log_textbox)
                
    if not recuperou:
        log_textbox.insert("end", _log("❌ Foco perdido. Pausando novamente.\n"))
        utils.PARAR_AUTOMACAO = True
        return tratar_pausa_e_retomar_foco(log_textbox, safe_configure_buttons_cb) # Recursivo
        
    return True

def lidar_com_erro_e_pausar(log_textbox, msg_erro, safe_update_gui_cb, safe_configure_buttons_cb):
    """Ativa a flag de pausa e chama a rotina de travamento."""
    utils.PARAR_AUTOMACAO = True
    safe_update_gui_cb(status="Pausado")
    
    log_textbox.insert("end", _log(f"🛑 {msg_erro}\n"))
    log_textbox.see("end")
    
    if tratar_pausa_e_retomar_foco(log_textbox, safe_configure_buttons_cb):
        safe_update_gui_cb(status="Rodando")
        return True
    return False

# ####################################################################
# --- EXECUÇÃO DOS PASSOS ---
# ####################################################################

def executar_passos_ciclo(valor, delay):
    try:
        pyperclip.copy(valor)
        time.sleep(0.05) # Delay mínimo para clipboard
        
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
        # Carrega imagens para RAM
        carregar_cache_imagens(log_textbox)

        # 1. Carregar Dados
        dados, repetir, msg = utils.ler_e_filtrar_dados(utils.NOME_ARQUIVO_ALVO, cidade_filtro, backlog_filtro, log_textbox)
        log_textbox.insert("end", _log(msg + "\n"))
        
        if repetir == 0:
            safe_update_gui_cb(status="Finalizado")
            return

        # 2. Foco Inicial
        focar_janela_por_titulo("Excel", log_textbox)
        focar_janela_por_titulo("Opera", log_textbox) or focar_janela_por_titulo("Google Chrome", log_textbox) or focar_janela_por_titulo("Microsoft Edge", log_textbox)

        safe_update_gui_cb(status="Rodando", total_ciclos=repetir)
        
        for i in range(utils.INDICE_ATUAL_DO_CICLO, repetir):
            utils.INDICE_ATUAL_DO_CICLO = i
            linha = dados.iloc[i]
            
            try:
                if 'Waybill No' in dados.columns: val = str(linha['Waybill No'])
                elif 'Motorista ID' in dados.columns: val = str(linha['Motorista ID'])
                else: val = str(linha.iloc[0])
            except: val = "???"

            log_textbox.insert("end", _log(f"⚙️ Ciclo {i+1}/{repetir}: {val}\n"))
            log_textbox.see("end")
            safe_update_gui_cb(ciclo_atual=i+1)

            # === LOOP DE BLINDAGEM (RETRY) ===
            while True:
                # 1. Pausa Manual (ESC)
                if utils.PARAR_AUTOMACAO:
                    lidar_com_erro_e_pausar(log_textbox, "Pausa Manual solicitada", safe_update_gui_cb, safe_configure_buttons_cb)

                # 2. Erro Visual (ANTES DA AÇÃO)
                tem_erro, nome_erro = verificar_excecoes_visuais(log_textbox)
                if tem_erro:
                    # Pausa e espera você clicar em Continuar
                    if lidar_com_erro_e_pausar(log_textbox, f"Erro na tela: {nome_erro}", safe_update_gui_cb, safe_configure_buttons_cb):
                        continue # Retomou? Verifica de novo (continue loop)
                    else: return

                # 3. Garante Foco
                focar_janela_por_titulo("Opera", log_textbox) or focar_janela_por_titulo("Google Chrome", log_textbox) or focar_janela_por_titulo("Microsoft Edge", log_textbox)

                # 4. Ação (Usa delay configurado na GUI)
                if not executar_passos_ciclo(val, utils.DELAY_ATUAL):
                    if lidar_com_erro_e_pausar(log_textbox, "Falha teclado", safe_update_gui_cb, safe_configure_buttons_cb):
                        continue
                    else: return

                # 5. Erro Visual (DEPOIS DA AÇÃO) - Verifica se apareceu algo após o Enter
                tem_erro_pos, nome_erro_pos = verificar_excecoes_visuais(log_textbox)
                if tem_erro_pos:
                     if lidar_com_erro_e_pausar(log_textbox, f"Erro pós-ação: {nome_erro_pos}", safe_update_gui_cb, safe_configure_buttons_cb):
                        continue # Se apareceu erro, repete o ciclo
                     else: return

                break # Sucesso, sai do loop de retry e vai para próxima linha

        winsound.Beep(1000, 500)
        log_textbox.insert("end", _log("✅ Finalizado!\n"))
        safe_update_gui_cb(status="Finalizado")

    except Exception as e:
        log_textbox.insert("end", _log(f"❌ ERRO FATAL: {e}\n"))
        safe_update_gui_cb(status="Erro")
        sucesso = False
        
    finally:
        if sucesso and utils.INDICE_ATUAL_DO_CICLO == repetir - 1:
            utils.INDICE_ATUAL_DO_CICLO = 0
        safe_configure_buttons_cb(iniciar_state="normal", continuar_state="disabled")
        utils.PARAR_AUTOMACAO = False