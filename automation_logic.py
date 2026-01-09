import pyautogui
import time
import pandas as pd     
import pyperclip
import winsound
import os
import utils
import ctypes

# --- AJUSTES DE SISTEMA (DPI & VELOCIDADE) ---
try:
    # Garante que o robô veja a tela na resolução real
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    # Velocidade alta, mas segura para o Windows não ignorar teclas
    pyautogui.PAUSE = 0.02 
except:
    pass

# ####################################################################
# --- CONFIGURAÇÃO DE RECURSOS ---
# ####################################################################

# Caminhos ajustados para a pasta assets (compatível com seu build.py)
IMAGENS_DE_EXCECAO = [
    os.path.join("assets", "erro_baixada.png"),
    os.path.join("assets", "erro_devolucao.png")
]

# Lista de palavras-chave para o Radar de Texto (Baseado no log 1352)
PALAVRAS_BASE_ERRO = [
    "1352", "ASSINATURA", "NÃO SUPORTA", "ERRO", 
    "FALHA", "BAIXADA", "INVALID", "JA FOI", "JÁ FOI", "CANCELAR", "FECHAR"
]

CACHE_CAMINHOS = []
CACHE_CARREGADO = False
LISTA_DINAMICA_ERROS = []

def _log(msg):
    """Formata o log com timestamp."""
    return f"[{time.strftime('%H:%M:%S')}] {msg}"

# ####################################################################
# --- MÓDULO: CONTROLE DE JANELAS ---
# ####################################################################

def focar_janela_por_titulo(titulo_parcial, log_textbox):
    """
    Tenta encontrar e trazer para frente uma janela específica.
    """
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
    """
    Varre os navegadores conhecidos para garantir o foco antes de digitar.
    """
    navegadores = ["Opera", "Google Chrome", "Microsoft Edge", "Firefox", "Brave"]
    for nav in navegadores:
        if focar_janela_por_titulo(nav, log_textbox):
            return True
    
    # Se não achar, loga aviso mas não trava
    log_textbox.insert("end", _log("⚠️ Aviso: Navegador não detectado. Usando janela ativa.\n"))
    return True 

# ####################################################################
# --- MÓDULO: DETECÇÃO (RADAR) ---
# ####################################################################

def carregar_recursos_detecao(log_textbox):
    """Carrega imagens e lista de erros apenas uma vez."""
    global CACHE_CAMINHOS, CACHE_CARREGADO, LISTA_DINAMICA_ERROS
    if CACHE_CARREGADO: return

    # 1. Carrega Imagens
    for caminho_relativo in IMAGENS_DE_EXCECAO:
        path = utils.resource_path(caminho_relativo)
        if os.path.exists(path):
            CACHE_CAMINHOS.append(path)
        else:
            # Fallback: Tenta na raiz
            path_root = utils.resource_path(os.path.basename(caminho_relativo))
            if os.path.exists(path_root): CACHE_CAMINHOS.append(path_root)

    # 2. Carrega Palavras do txt externo
    LISTA_DINAMICA_ERROS = PALAVRAS_BASE_ERRO.copy()
    caminho_txt = utils.get_external_path("palavras_erro.txt")
    if os.path.exists(caminho_txt):
        try:
            with open(caminho_txt, 'r', encoding='utf-8') as f:
                linhas = [l.strip().upper() for l in f.readlines() if l.strip() and not l.startswith("#")]
                for l in linhas:
                    if l not in LISTA_DINAMICA_ERROS: LISTA_DINAMICA_ERROS.append(l)
        except Exception as e:
            log_textbox.insert("end", _log(f"⚠️ Erro ao ler palavras_erro.txt: {e}\n"))
            
    CACHE_CARREGADO = True

def verificar_presenca_erro():
    """
    Verifica erro via Clipboard (Rápido) e depois Visual (Lento).
    """
    # 1. TENTATIVA RÁPIDA: TEXTO (Clipboard)
    try:
        pyperclip.copy("") 
        
        # Sequência para capturar texto do modal
        pyautogui.hotkey('ctrl', 'a') 
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.05) 
        
        conteudo = str(pyperclip.paste()).upper()
        if conteudo:
            for palavra in LISTA_DINAMICA_ERROS:
                if palavra in conteudo:
                    pyautogui.press('esc') 
                    return True, f"Texto: {palavra}"
        
        pyautogui.press('right') 
    except: pass

    # 2. TENTATIVA VISUAL (Fallback)
    for caminho_img in CACHE_CAMINHOS:
        try:
            if pyautogui.locateOnScreen(caminho_img, grayscale=True, confidence=0.75):
                return True, f"Imagem Detectada"
        except: pass

    return False, None

# ####################################################################
# --- MÓDULO: TRATAMENTO DE PAUSA (SEGURANÇA) ---
# ####################################################################

def lidar_com_erro_e_pausar(log_textbox, motivo, safe_update_gui_cb, safe_configure_buttons_cb):
    """
    Função centralizada para pausar o sistema quando algo dá errado.
    Retorna True se o usuário pediu para Continuar.
    Retorna False se o usuário pediu para Cancelar/Parar.
    """
    winsound.Beep(800, 500)
    utils.PARAR_AUTOMACAO = True
    
    log_textbox.insert("end", _log(f"🛑 BLOQUEIO DETECTADO ({motivo}).\n"))
    log_textbox.insert("end", "   -> Resolva no navegador e clique em CONTINUAR.\n")
    log_textbox.see("end")
    
    # Atualiza GUI para estado de Pausa
    safe_update_gui_cb(status="Pausado")
    safe_configure_buttons_cb(iniciar_state="disabled", continuar_state="normal")
    
    # Loop de Travamento (Espera o usuário)
    while utils.PARAR_AUTOMACAO:
        if getattr(utils, 'CANCELAR_AUTOMACAO', False): 
            return False # Usuário cancelou
        time.sleep(0.5)
    
    # Se saiu do loop sem cancelar, é porque vai continuar
    if getattr(utils, 'CANCELAR_AUTOMACAO', False): 
        return False

    # Retomada
    log_textbox.insert("end", _log("▶️ Retomando operação...\n"))
    safe_update_gui_cb(status="Rodando")
    safe_configure_buttons_cb(iniciar_state="disabled", continuar_state="disabled")
    
    # Garante o foco antes de devolver o controle
    garantir_foco_navegador(log_textbox)
    return True

# ####################################################################
# --- CORE DA AUTOMAÇÃO ---
# ####################################################################

def automacao_core(log_textbox, cidade_filtro, backlog_filtro, delay_inicial, safe_configure_buttons_cb, safe_update_gui_cb):
    try:
        carregar_recursos_detecao(log_textbox)
        utils.DELAY_ATUAL = delay_inicial
        
        # 1. LEITURA DE DADOS (PANDAS)
        dados, repetir, msg = utils.ler_e_filtrar_dados(utils.NOME_ARQUIVO_ALVO, cidade_filtro, backlog_filtro, log_textbox)
        log_textbox.insert("end", _log(f"{msg}\n"))
        
        if repetir == 0:
            safe_update_gui_cb(status="Finalizado")
            return

        # 2. PREPARAÇÃO
        focar_janela_por_titulo("Excel", log_textbox)
        garantir_foco_navegador(log_textbox)
        
        safe_update_gui_cb(status="Rodando", total_ciclos=repetir)
        
        # 3. LOOP PRINCIPAL
        for i in range(utils.INDICE_ATUAL_DO_CICLO, repetir):
            
            # Verifica Cancelamento Global
            if getattr(utils, 'CANCELAR_AUTOMACAO', False): 
                log_textbox.insert("end", _log("⛔ Operação cancelada pelo usuário.\n"))
                break
            
            utils.INDICE_ATUAL_DO_CICLO = i
            
            # --- VALIDAÇÃO ROBUSTA DE DADOS (PANDAS) ---
            try:
                linha = dados.iloc[i]
                val = ""
                
                # Prioridade: Waybill > Motorista ID > Coluna 0
                if 'Waybill No' in dados.columns: 
                    val = str(linha['Waybill No']).strip()
                elif 'Motorista ID' in dados.columns: 
                    val = str(linha['Motorista ID']).strip()
                else: 
                    val = str(linha.iloc[0]).strip()
                
                # Validação de Nulos/Vazios
                if val.lower() == 'nan' or val == '' or val.lower() == 'nat':
                    log_textbox.insert("end", _log(f"⚠️ Linha {i+1} inválida/vazia no Excel. Pulando.\n"))
                    continue
                    
            except Exception as e:
                log_textbox.insert("end", _log(f"⚠️ Erro ao ler dados da linha {i+1}: {e}\n"))
                continue

            log_textbox.insert("end", _log(f"Ciclo {i+1}/{repetir}: {val}\n"))
            log_textbox.see("end")
            safe_update_gui_cb(ciclo_atual=i+1)

            # --- LOOP DE TENTATIVA (RETRY) ---
            while True:
                # 1. Verifica Cancelamento
                if getattr(utils, 'CANCELAR_AUTOMACAO', False): break
                
                # 2. Verifica Pausa Manual (ESC)
                if utils.PARAR_AUTOMACAO:
                    if not lidar_com_erro_e_pausar(log_textbox, "Pausa Manual [ESC]", safe_update_gui_cb, safe_configure_buttons_cb):
                        break # Cancelou
                    continue # Retomou, tenta colar de novo

                # 3. AÇÃO: COPIAR E COLAR
                try:
                    pyperclip.copy(val)
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(utils.DELAY_ATUAL)
                    pyautogui.press('enter')
                except Exception as e:
                    # Se falhar o teclado, pausa e pede ajuda
                    if not lidar_com_erro_e_pausar(log_textbox, f"Erro Teclado: {e}", safe_update_gui_cb, safe_configure_buttons_cb):
                        break
                    continue

                # 4. RADAR DE PERSISTÊNCIA (2.0s - Vigília)
                tempo_limite = time.time() + 2.0
                tem_erro = False
                motivo = None
                
                while time.time() < tempo_limite:
                    if utils.PARAR_AUTOMACAO: break # Pausa durante vigília
                    
                    tem_erro, motivo = verificar_presenca_erro()
                    if tem_erro: break
                    time.sleep(0.15)

                # 5. TRATAMENTO DO ERRO
                if tem_erro:
                    # Chama a função centralizada de tratamento
                    sucesso_retomada = lidar_com_erro_e_pausar(log_textbox, motivo, safe_update_gui_cb, safe_configure_buttons_cb)
                    
                    if sucesso_retomada:
                        # SE O USUÁRIO CLICOU EM CONTINUAR APÓS UM ERRO:
                        # Assumimos que ele resolveu/ignorou e quer ir para o PRÓXIMO.
                        log_textbox.insert("end", _log("▶️ Erro tratado. Pulando para o próximo registro.\n"))
                        break # Sai do 'while True', vai para o próximo 'i'
                    else:
                        break # Cancelou, sai do 'while True' e vai parar o loop 'for'
                
                # Se não houve erro, sucesso! Vai para o próximo 'i'
                break 

            # Pequeno respiro
            time.sleep(0.05)

        # FIM DO LOOP
        if not getattr(utils, 'CANCELAR_AUTOMACAO', False):
            log_textbox.insert("end", _log("✅ Lista Finalizada com Sucesso.\n"))
            safe_update_gui_cb(status="Finalizado")
        else:
            safe_update_gui_cb(status="Parado")

    except Exception as e:
        log_textbox.insert("end", _log(f"❌ ERRO CRÍTICO NO CORE: {e}\n"))
        log_textbox.see("end")
        safe_update_gui_cb(status="Erro")
    finally:
        focar_janela_por_titulo("Atribuidor", log_textbox)
        utils.INDICE_ATUAL_DO_CICLO = 0
        safe_configure_buttons_cb(iniciar_state="normal", continuar_state="disabled")