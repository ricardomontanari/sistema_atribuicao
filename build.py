import PyInstaller.__main__
import os
import customtkinter
import CTkMessagebox

# --- CONFIGURAÇÕES DO BUILD ---
NOME_EXECUTAVEL = "Atribuidor"
SCRIPT_PRINCIPAL = "main.py" 
# Caminho do ícone dentro da pasta assets
ICONE = os.path.join("assets", "icone.ico") 

def obter_caminho_lib(lib):
    """Retorna o diretório de instalação da biblioteca."""
    return os.path.dirname(lib.__file__)

def criar_executavel():
    print(f"🚀 Build iniciado para '{NOME_EXECUTAVEL}'...")

    # Verifica se o ícone existe antes de tentar usar
    if not os.path.exists(ICONE):
        print(f"⚠️ Aviso: Ícone não encontrado em {ICONE}. O build seguirá sem ícone personalizado.")

    # Localiza caminhos das libs gráficas
    ctk_path = obter_caminho_lib(customtkinter)
    msg_path = obter_caminho_lib(CTkMessagebox)
    
    # Define separador de arquivos (Windows usa ';', Linux usa ':')
    sep = ";" if os.name == 'nt' else ":"

    args = [
        SCRIPT_PRINCIPAL,
        f'--name={NOME_EXECUTAVEL}',
        '--onefile',       # Gera um único arquivo .exe
        '--windowed',      # Executa sem abrir o console preto (CMD)
        '--clean',         # Limpa caches anteriores
        '--noconfirm',     # Sobrescreve sem perguntar
        
        # --- INCLUSÃO DE ASSETS (Pasta Inteira) ---
        # Copia a pasta 'assets' local para 'assets' dentro do executável
        f'--add-data=assets{sep}assets',
        
        # --- BIBLIOTECAS GRÁFICAS ---
        f'--add-data={ctk_path}{sep}customtkinter',
        f'--add-data={msg_path}{sep}CTkMessagebox',
        
        # --- DEPENDÊNCIAS DO RADAR (CRÍTICO) ---
        # 'collect-all' garante que binários do OpenCV e PyAutoGUI sejam copiados
        '--collect-all=cv2',
        '--collect-all=pyautogui',
        
        # --- IMPORTS ESCONDIDOS ---
        # Ajuda o PyInstaller a encontrar módulos importados dinamicamente
        '--hidden-import=PIL._tkinter_finder',
        '--hidden-import=pandas',
        '--hidden-import=sqlite3',
        '--hidden-import=numpy',
        '--hidden-import=pyscreeze',
        '--hidden-import=pyautogui',
    ]

    # Adiciona ícone se existir
    if os.path.exists(ICONE):
        args.append(f'--icon={ICONE}')

    print("📦 Empacotando... Isso pode levar alguns minutos.")
    
    # Executa o PyInstaller
    PyInstaller.__main__.run(args)
    
    print(f"✅ Sucesso! Seu executável está na pasta: {os.path.abspath('dist')}")

if __name__ == "__main__":
    criar_executavel()