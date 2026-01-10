import pyautogui
import time
import pandas as pd     
import pyperclip
import winsound
import os
import utils
import ctypes

# --- AJUSTES DE SISTEMA ---
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    pyautogui.PAUSE = 0.02 
except:
    pass

# ####################################################################
# --- CONFIGURAÇÃO DE RECURSOS ---
# ####################################################################

# Carrega imagens da pasta assets
IMAGENS_DE_EXCECAO = []
assets_dir = utils.resource_path("assets")
if os.path.exists(assets_dir):
    for f in os.listdir(assets_dir):
        if f.lower().endswith(".png"):
            IMAGENS_DE_EXCECAO.append(os.path.join("assets", f))

# Palavras-chave de Título (Janelas de Erro)
PALAVRAS_TITULO_ERRO = [
    "ERRO", "FALHA", "ATENÇÃO", "AVISO", "ERROR", "PROBLEM", "ALERT", 
    "MENSAGEM DA PÁGINA", "CONFIRMAÇÃO"
]

CACHE_CAMINHOS = []
CACHE_CARREGADO = False

def _log(msg):
    return f"[{time.strftime('%H:%M:%S')}] {msg}"

# ####################################################################
# --- CONTROLE DE JANELAS (SILENCIOSO) ---
# ####################################################################

def focar_janela_por_titulo(titulo_parcial, log_textbox):
    try:
        janelas = pyautogui.getWindowsWithTitle(titulo_parcial)
        if janelas:
            janela = janelas[0]
            if janela.isMinimized: janela.restore()
            janela.activate()
            time.sleep(0.1) 
            return True
    except: pass
    return False

def garantir_foco_navegador(log_textbox):
    navegadores = ["Opera", "Google Chrome", "Microsoft Edge", "Firefox", "Brave"]
    for nav in navegadores:
        if focar_janela_por_titulo(nav, log_textbox):
            return True
    log_textbox.insert("end", _log("⚠️ Aviso: Navegador não detectado (usando janela ativa).\n"))
    return True 

# ####################################################################
# --- DETECÇÃO VISUAL (SILENCIOSA) ---
# ####################################################################

def carregar_recursos_detecao(log_textbox):
    global CACHE_CAMINHOS, CACHE_CARREGADO
    if CACHE_CARREGADO: return

    for caminho_relativo in IMAGENS_DE_EXCECAO:
        path = utils.resource_path(caminho_relativo)
        if os.path.exists(path):
            CACHE_CAMINHOS.append(path)
        else:
            path_root = utils.resource_path(os.path.basename(caminho_relativo))
            if os.path.exists(path_root): CACHE_CAMINHOS.append(path_root)
            
    CACHE_CARREGADO = True

def verificar_presenca_erro():
    """
    Verifica erro de forma 100% VISUAL e INVISÍVEL.
    """
    # 1. Título da Janela
    try:
        titulo_ativo = pyautogui.getActiveWindowTitle()
        if titulo_ativo:
            titulo_upper = titulo_ativo.upper()
            for palavra in PALAVRAS_TITULO_ERRO:
                if palavra in titulo_upper:
                    return True, f"Janela: {palavra}"
    except: pass

    # 2. Imagem (Passivo)
    for caminho_img in CACHE_CAMINHOS:
        try:
            if pyautogui.locateOnScreen(caminho_img, grayscale=True, confidence=0.8):
                nome_erro = os.path.basename(caminho_img).replace('.png', '')
                return True, f"Imagem: {nome_erro}"
        except: pass

    return False, None

# ####################################################################
# --- TRATAMENTO DE PAUSA ---
# ####################################################################

def lidar_com_erro_e_pausar(log_textbox, motivo, safe_update_gui_cb, safe_configure_buttons_cb):
    winsound.Beep(800, 500)
    utils.PARAR_AUTOMACAO = True
    
    log_textbox.insert("end", _log(f"🛑 BLOQUEIO DETECTADO ({motivo}).\n"))
    log_textbox.insert("end", "   -> Resolva no navegador e clique em CONTINUAR.\n")
    log_textbox.see("end")
    
    safe_update_gui_cb(status="Pausado")
    safe_configure_buttons_cb(iniciar_state="disabled", continuar_state="normal")
    
    while utils.PARAR_AUTOMACAO:
        if getattr(utils, 'CANCELAR_AUTOMACAO', False): 
            return False 
        time.sleep(0.5)
    
    if getattr(utils, 'CANCELAR_AUTOMACAO', False): 
        return False

    log_textbox.insert("end", _log("▶️ Retomando operação...\n"))
    safe_update_gui_cb(status="Rodando")
    safe_configure_buttons_cb(iniciar_state="disabled", continuar_state="disabled")
    
    garantir_foco_navegador(log_textbox)
    return True

# ####################################################################
# --- CORE DA AUTOMAÇÃO ---
# ####################################################################

def automacao_core(log_textbox, cidade_filtro, backlog_filtro, delay_inicial, safe_configure_buttons_cb, safe_update_gui_cb):
    try:
        carregar_recursos_detecao(log_textbox)
        utils.DELAY_ATUAL = delay_inicial
        
        dados, repetir, msg = utils.ler_e_filtrar_dados(utils.NOME_ARQUIVO_ALVO, cidade_filtro, backlog_filtro, log_textbox)
        log_textbox.insert("end", _log(f"{msg}\n"))
        
        if repetir == 0:
            safe_update_gui_cb(status="Finalizado")
            return

        focar_janela_por_titulo("Excel", log_textbox)
        garantir_foco_navegador(log_textbox)
        
        safe_update_gui_cb(status="Rodando", total_ciclos=repetir)
        
        for i in range(utils.INDICE_ATUAL_DO_CICLO, repetir):
            
            if getattr(utils, 'CANCELAR_AUTOMACAO', False): 
                log_textbox.insert("end", _log("⛔ Operação cancelada.\n"))
                break
            
            utils.INDICE_ATUAL_DO_CICLO = i
            
            try:
                linha = dados.iloc[i]
                val = ""
                if 'Waybill No' in dados.columns: val = str(linha['Waybill No']).strip()
                elif 'Motorista ID' in dados.columns: val = str(linha['Motorista ID']).strip()
                else: val = str(linha.iloc[0]).strip()
                
                if val.lower() == 'nan' or val == '' or val.lower() == 'nat':
                    log_textbox.insert("end", _log(f"⚠️ Linha {i+1} inválida. Pulando.\n"))
                    continue
            except Exception as e:
                log_textbox.insert("end", _log(f"⚠️ Erro dados linha {i+1}: {e}\n"))
                continue

            log_textbox.insert("end", _log(f"Ciclo {i+1}/{repetir}: {val}\n"))
            log_textbox.see("end")
            safe_update_gui_cb(ciclo_atual=i+1)

            while True:
                if getattr(utils, 'CANCELAR_AUTOMACAO', False): break
                
                if utils.PARAR_AUTOMACAO:
                    if not lidar_com_erro_e_pausar(log_textbox, "Pausa Manual", safe_update_gui_cb, safe_configure_buttons_cb):
                        break
                    continue

                # AÇÃO
                try:
                    pyperclip.copy(val)
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(utils.DELAY_ATUAL)
                    pyautogui.press('enter')
                except Exception as e:
                    if not lidar_com_erro_e_pausar(log_textbox, f"Erro Teclado: {e}", safe_update_gui_cb, safe_configure_buttons_cb):
                        break
                    continue

                # --- RADAR OTIMIZADO (VELOCIDADE MÁXIMA) ---
                # Reduzido de 2.0s para 0.6s. 
                # Se o erro não aparecer em meio segundo, assumimos sucesso.
                tempo_limite = time.time() + 0.6
                tem_erro = False
                motivo = None
                
                while time.time() < tempo_limite:
                    if utils.PARAR_AUTOMACAO: break 
                    
                    tem_erro, motivo = verificar_presenca_erro()
                    if tem_erro: break
                    time.sleep(0.05) # Polling ultra-rápido

                if tem_erro:
                    sucesso_retomada = lidar_com_erro_e_pausar(log_textbox, motivo, safe_update_gui_cb, safe_configure_buttons_cb)
                    if sucesso_retomada:
                        log_textbox.insert("end", _log("▶️ Erro tratado. Próximo registro.\n"))
                        break 
                    else:
                        break 
                
                break 

            # Sem sleep extra no final para maximizar velocidade

        if not getattr(utils, 'CANCELAR_AUTOMACAO', False):
            log_textbox.insert("end", _log("✅ Finalizado com Sucesso.\n"))
            safe_update_gui_cb(status="Finalizado")
        else:
            safe_update_gui_cb(status="Parado")

    except Exception as e:
        log_textbox.insert("end", _log(f"❌ ERRO CRÍTICO: {e}\n"))
        safe_update_gui_cb(status="Erro")
    finally:
        focar_janela_por_titulo("Atribuidor", log_textbox)
        utils.INDICE_ATUAL_DO_CICLO = 0
        safe_configure_buttons_cb(iniciar_state="normal", continuar_state="disabled")