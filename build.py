import PyInstaller.__main__
import os
import customtkinter
import CTkMessagebox

NOME_EXECUTAVEL = "Atribuidor"
SCRIPT_PRINCIPAL = "main.py" 
# Aponta para a pasta assets
ICONE = os.path.join("assets", "icone.ico") 

def obter_caminho_lib(lib):
    return os.path.dirname(lib.__file__)

def criar_executavel():
    print(f"🚀 Build iniciado para '{NOME_EXECUTAVEL}'...")

    # Verifica ícone
    if not os.path.exists(ICONE):
        print(f"⚠️ Aviso: Ícone não encontrado em {ICONE}")

    ctk_path = obter_caminho_lib(customtkinter)
    msg_path = obter_caminho_lib(CTkMessagebox)
    sep = ";" if os.name == 'nt' else ":"

    args = [
        SCRIPT_PRINCIPAL,
        f'--name={NOME_EXECUTAVEL}',
        '--onefile',       
        '--windowed',      
        '--clean',         
        '--noconfirm',     
        
        # --- ASSETS ---
        # Copia a pasta assets inteira para dentro do executável
        f'--add-data=assets{sep}assets',
        
        # Bibliotecas gráficas
        f'--add-data={ctk_path}{sep}customtkinter',
        f'--add-data={msg_path}{sep}CTkMessagebox',
        
        # --- IMPORTANTE: DEPENDÊNCIAS DO RADAR ---
        '--collect-all=cv2',
        '--collect-all=pyautogui',
        '--hidden-import=PIL._tkinter_finder',
        '--hidden-import=pandas',
        '--hidden-import=sqlite3',
        '--hidden-import=numpy',  
    ]

    if os.path.exists(ICONE):
        args.append(f'--icon={ICONE}')

    print("📦 Empacotando...")
    PyInstaller.__main__.run(args)
    print(f"✅ Sucesso! Executável em: {os.path.abspath('dist')}")

if __name__ == "__main__":
    criar_executavel()