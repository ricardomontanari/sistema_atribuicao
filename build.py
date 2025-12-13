import PyInstaller.__main__
import os
import customtkinter
import CTkMessagebox

# --- CONFIGURAÇÕES DO PROJETO ---
NOME_EXECUTAVEL = "Atribuidor"
SCRIPT_PRINCIPAL = "main.py"
ICONE = "icone.ico"  # Opcional: Se não tiver um ícone, o script ignora

def obter_caminho_biblioteca(lib):
    """Retorna o diretório de instalação da biblioteca para inclusão de dados."""
    return os.path.dirname(lib.__file__)

def criar_executavel():
    print(f"🚀 Iniciando build do '{NOME_EXECUTAVEL}'...")

    # 1. Localiza os caminhos das bibliotecas que possuem arquivos de dados (temas/json)
    ctk_path = obter_caminho_biblioteca(customtkinter)
    msg_path = obter_caminho_biblioteca(CTkMessagebox)

    # 2. Define o separador de caminhos correto para o SO (Windows usa ';')
    sep = ";" if os.name == 'nt' else ":"

    # 3. Monta os argumentos para o PyInstaller
    args = [
        SCRIPT_PRINCIPAL,
        f'--name={NOME_EXECUTAVEL}',
        '--onefile',       # Gera um único arquivo .exe
        '--windowed',      # Executa sem abrir o console preto (CMD)
        '--clean',         # Limpa caches de compilações anteriores
        '--noconfirm',     # Sobrescreve a pasta dist sem perguntar
        
        # --- INCLUSÃO DE ARQUIVOS (Origem;Destino) ---
        
        # Inclui temas do CustomTkinter
        f'--add-data={ctk_path}{sep}customtkinter',
        
        # Inclui assets do CTkMessagebox
        f'--add-data={msg_path}{sep}CTkMessagebox',
        
        # Inclui TODAS as imagens PNG (erro_devolucao.png, etc.) na raiz do executável
        f'--add-data=*.png{sep}.',
        
        # --- IMPORTAÇÕES FORÇADAS (Hidden Imports) ---
        # Garante que o PyInstaller encontre módulos que às vezes passam despercebidos
        '--hidden-import=PIL._tkinter_finder',
        '--hidden-import=pandas',
        '--hidden-import=sqlite3',
        '--hidden-import=babel.numbers',
        '--hidden-import=pyautogui',
    ]

    # Adiciona ícone apenas se o arquivo existir
    if ICONE and os.path.exists(ICONE):
        args.append(f'--icon={ICONE}')
        print(f"🎨 Ícone '{ICONE}' detectado e incluído.")

    # 4. Executa o PyInstaller
    print("📦 Empacotando arquivos... (Isso pode levar alguns minutos)")
    PyInstaller.__main__.run(args)

    print("\n" + "="*50)
    print("✅ SUCESSO! O Executável foi criado.")
    print(f"📂 Localização: {os.path.abspath('dist')}")
    print("="*50)

if __name__ == "__main__":
    criar_executavel()