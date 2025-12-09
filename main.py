import customtkinter as ctk
# Importamos as duas classes que estarão no gui.py
from gui import LoginWindow, MainWindow 

if __name__ == "__main__":
    # Configuração inicial do tema
    ctk.set_appearance_mode("Dark")  # Modes: "System" (default), "Dark", "Light"
    ctk.set_default_color_theme("green")  # Themes: "blue" (default), "green", "dark-blue"
    
    # Inicia a janela de login
    login_app = LoginWindow()
    login_app.mainloop()