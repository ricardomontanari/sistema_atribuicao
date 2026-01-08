import PyInstaller.__main__
import os
import customtkinter
import CTkMessagebox

# --- CONFIGURAÇÕES DO BUILD ---
NOME_EXECUTAVEL = "Atribuidor"
SCRIPT_PRINCIPAL = "main.py" 
ICONE = "app.ico" # Se tiver um ícone .ico, ele será usado. Caso contrário, ignora.

# Lista de imagens críticas para verificar antes de começar o build
# Isso evita criar um executável quebrado se a imagem não estiver na pasta.
IMAGENS_CRITICAS = ["erro_baixada.png"]

def obter_caminho_lib(lib):
    """Retorna o caminho da pasta da biblioteca instalada."""
    return os.path.dirname(lib.__file__)

def criar_executavel():
    print(f"🚀 Iniciando build do '{NOME_EXECUTAVEL}'...")

    # 1. VERIFICAÇÃO PRÉ-BUILD
    # Garante que as imagens existem antes de empacotar
    for img in IMAGENS_CRITICAS:
        if not os.path.exists(img):
            print(f"❌ ERRO CRÍTICO: A imagem '{img}' não está na pasta!")
            print("   O executável não vai funcionar sem ela.")
            return

    # 2. Localiza caminhos de bibliotecas externas (CustomTkinter precisa disso)
    ctk_path = obter_caminho_lib(customtkinter)
    msg_path = obter_caminho_lib(CTkMessagebox)
    sep = ";" if os.name == 'nt' else ":"

    # 3. Argumentos para o PyInstaller
    args = [
        SCRIPT_PRINCIPAL,
        f'--name={NOME_EXECUTAVEL}',
        '--onefile',       # Gera um único arquivo .exe (portátil)
        '--windowed',      # Executa sem abrir o console preto (CMD)
        '--clean',         # Limpa caches de compilações anteriores
        '--noconfirm',     # Sobrescreve a pasta dist sem perguntar
        
        # --- INCLUSÃO DE ARQUIVOS (Origem;Destino) ---
        
        # Inclui os temas e json do CustomTkinter
        f'--add-data={ctk_path}{sep}customtkinter',
        
        # Inclui assets do CTkMessagebox
        f'--add-data={msg_path}{sep}CTkMessagebox',
        
        # Inclui TODAS as imagens PNG na raiz do executável
        # Isso garante que erro_baixada.png vá junto
        f'--add-data=*.png{sep}.',
        
        # --- IMPORTS ESCONDIDOS (Hidden Imports) ---
        # Garante que o PyInstaller encontre módulos que ele geralmente esquece
        '--hidden-import=PIL._tkinter_finder',
        '--hidden-import=pandas',
        '--hidden-import=openpyxl',
        '--hidden-import=sqlite3',
        '--hidden-import=pyautogui',
        '--hidden-import=pyscreeze', # Necessário para locateOnScreen
        
        # CRÍTICO: Necessários para o parâmetro 'confidence' funcionar
        '--hidden-import=cv2',    
        '--hidden-import=numpy',  
    ]

    # Adiciona ícone se o arquivo existir
    if ICONE and os.path.exists(ICONE):
        args.append(f'--icon={ICONE}')
        print(f"🎨 Ícone incluído: {ICONE}")

    # 4. Executa o PyInstaller
    print("📦 Empacotando... (Isso pode levar alguns minutos)")
    PyInstaller.__main__.run(args)

    print("\n" + "="*50)
    print("✅ SUCESSO! Executável criado.")
    print(f"📂 Localização: {os.path.abspath('dist')}")
    print("="*50)

if __name__ == "__main__":
    criar_executavel()