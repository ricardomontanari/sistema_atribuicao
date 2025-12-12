import customtkinter as ctk
from gui import App

# Configuração Global de Tema
# Modos: "System" (Padrão do OS), "Dark" (Escuro), "Light" (Claro)
ctk.set_appearance_mode("Dark") 
# Temas de Cor: "blue" (Padrão), "green", "dark-blue"
ctk.set_default_color_theme("green")

if __name__ == "__main__":
    # Inicializa a Aplicação Unificada
    app = App()
    
    # Previne fechamento acidental de threads ao fechar a janela principal
    try:
        app.mainloop()
    except KeyboardInterrupt:
        print("Aplicação encerrada via terminal.")