import customtkinter as ctk
import threading 
import utils
import pyautogui
import json 
import os   
import time
from CTkMessagebox import CTkMessagebox
from automation_logic import automacao_core
from db_manager import (
    setup_database, 
    adicionar_cidade, 
    buscar_nomes_cidades, 
    listar_cidades,
    excluir_cidade,
    adicionar_usuario,     
    verificar_credenciais,
    buscar_nome_cidade_por_id,
    listar_usuarios,
    excluir_usuario
)

# Arquivo para salvar a preferência do usuário (localmente)
CONFIG_LOGIN_FILE = "login_config.json"

# --- CONSTANTES DE ESTILO (PADRONIZAÇÃO) ---
BTN_HEIGHT_DEFAULT = 35
BTN_HEIGHT_MAIN = 40

COLOR_SUCCESS = "green"       # Iniciar, Adicionar, Salvar
COLOR_SUCCESS_HOVER = "darkgreen"

COLOR_DANGER = "red"          # Excluir, Parar, Remover
COLOR_DANGER_HOVER = "darkred"

# Azul Real (Royal Blue) para Continuar
COLOR_WARNING = "#1565C0"      # Continuar
COLOR_WARNING_HOVER = "#0D47A1"

# NOVA COR: Teal (Verde-petróleo) para Excel/Arquivo
# Distinto do Azul e do Verde padrão
COLOR_NEUTRAL = "#00796B"     # Excel
COLOR_NEUTRAL_HOVER = "#004D40"

# ####################################################################
# --- COMPONENTE PERSONALIZADO: TEXTBOX SOMENTE LEITURA ---
# ####################################################################

class ReadOnlyTextbox(ctk.CTkTextbox):
    """
    Uma caixa de texto que impede o usuário de digitar (state='disabled'),
    mas permite que o código insira/delete texto interceptando os métodos.
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.configure(state="disabled")

    def insert(self, index, text, tags=None):
        self.configure(state="normal")
        super().insert(index, text, tags)
        self.configure(state="disabled")

    def delete(self, index1, index2=None):
        self.configure(state="normal")
        super().delete(index1, index2)
        self.configure(state="disabled")

# ####################################################################
# --- FUNÇÕES AUXILIARES GLOBAIS PARA POP-UPS ---
# ####################################################################

def _aplicar_estilo_padrao(msg_box):
    try:
        msg_box.icon_label.grid_configure(padx=(25, 25))
    except Exception:
        pass 
    msg_box.after(10, lambda: msg_box.focus_force())

def exibir_popup(title, message, icon="info"):
    msg_box = CTkMessagebox(title=title, message=message, icon=icon)
    _aplicar_estilo_padrao(msg_box)
    msg_box.bind("<Return>", lambda event: msg_box.button_event("OK"))
    return msg_box.get()

def exibir_confirmacao(title, message, icon="question", option_ok="Confirmar", option_cancel="Cancelar"):
    msg_box = CTkMessagebox(title=title, message=message, icon=icon, 
                            option_1=option_cancel, option_2=option_ok)
    _aplicar_estilo_padrao(msg_box)
    msg_box.bind("<Return>", lambda event: msg_box.button_event(option_ok))
    msg_box.bind("<Escape>", lambda event: msg_box.button_event(option_cancel))
    return msg_box.get()

# ####################################################################
# --- CLASSE 1: JANELA DE LOGIN (APP INICIAL) ---
# ####################################################################

class LoginWindow(ctk.CTk):
    def __init__(self):
        super().__init__()
        setup_database() 
        
        self.title("Login") 
        window_width = 350
        window_height = 380
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = int((screen_width - window_width) / 2)
        y = int((screen_height - window_height) / 2)
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.resizable(False, False)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1) 

        login_frame = ctk.CTkFrame(self)
        login_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        login_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(login_frame, text="Acesso ao Atribuidor", font=ctk.CTkFont(size=20, weight="bold")).grid(row=0, column=0, pady=(30, 20), padx=20)

        self.username_entry = ctk.CTkEntry(login_frame, placeholder_text="Usuário", width=200, height=35)
        self.username_entry.grid(row=1, column=0, pady=10, padx=20)

        self.password_entry = ctk.CTkEntry(login_frame, placeholder_text="Senha", show="*", width=200, height=35)
        self.password_entry.grid(row=2, column=0, pady=10, padx=20)

        self.lembrar_var = ctk.BooleanVar(value=False)
        self.lembrar_checkbox = ctk.CTkCheckBox(
            login_frame, text="Lembrar usuário", variable=self.lembrar_var,
            font=ctk.CTkFont(size=11), checkbox_width=14, checkbox_height=14, border_width=1       
        )
        self.lembrar_checkbox.grid(row=3, column=0, pady=(5, 10), padx=20)

        # Botão de Login (Texto Branco - Cor Neutral/Teal)
        self.login_button = ctk.CTkButton(
            login_frame, text="Acessar", command=self.attempt_login, width=200, 
            height=BTN_HEIGHT_MAIN, text_color="white",
            fg_color=COLOR_NEUTRAL, hover_color=COLOR_NEUTRAL_HOVER
        )
        self.login_button.grid(row=4, column=0, pady=20, padx=20)
        
        self.bind('<Return>', lambda event: self.attempt_login())
        self.carregar_usuario_salvo()
        self.after(100, lambda: self.password_entry.focus_set() if self.username_entry.get() else self.username_entry.focus_set())
        
    def carregar_usuario_salvo(self):
        if os.path.exists(CONFIG_LOGIN_FILE):
            try:
                with open(CONFIG_LOGIN_FILE, 'r') as f:
                    data = json.load(f)
                    saved_user = data.get("last_user", "")
                    if saved_user:
                        self.username_entry.insert(0, saved_user)
                        self.lembrar_var.set(True)
            except Exception: pass 

    def salvar_preferencia_usuario(self, username):
        data = {"last_user": username if self.lembrar_var.get() else ""}
        try:
            with open(CONFIG_LOGIN_FILE, 'w') as f: json.dump(data, f)
        except Exception: pass

    def attempt_login(self):
        user = self.username_entry.get()
        password = self.password_entry.get()
        if verificar_credenciais(user, password):
            self.salvar_preferencia_usuario(user)
            self.withdraw()
            main_app = MainWindow()
            main_app.mainloop()
            self.destroy()
        else:
            exibir_popup(title="Erro de Acesso", message="Usuário ou Senha inválidos.", icon="cancel")

# ####################################################################
# --- CLASSE 2: APLICAÇÃO PRINCIPAL (MAIN WINDOW) ---
# ####################################################################

class MainWindow(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"Sistema de Automação ({utils.VERSAO_SISTEMA})") 
        window_width = 650
        window_height = 700
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = int((screen_width - window_width) / 2)
        y = int((screen_height - window_height) / 2)
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.status_text = ctk.StringVar(value="PRONTO")
        self.status_color = ctk.StringVar(value="green")
        self.progresso_atual = ctk.DoubleVar(value=0)
        self.total_de_ciclos_var = ctk.IntVar(value=0)
        
        self.todas_cidades = []
        self.monitor_thread_started = False

        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.tabview.add("Automação")
        self.tabview.add("Cadastro")
        self.tabview.add("Usuários") 
        
        self.setup_automacao_tab()
        self.setup_cadastro_tab()
        self.setup_usuarios_tab() 
        self.carregar_cidades_db()

    def _safe_update_gui(self, status=None, total_ciclos=None, ciclo_atual=None):
        def update():
            if status:
                self.status_text.set(status.upper())
                colors = {"Rodando": "blue", "Pausado": "orange", "Erro": "red", "Finalizado": "green", "PRONTO": "gray"}
                self.status_color.set(colors.get(status, self.status_color.get()))
                self.status_lbl.configure(fg_color=self.status_color.get())
            if total_ciclos is not None:
                self.total_de_ciclos_var.set(total_ciclos)
            if ciclo_atual is not None:
                total = self.total_de_ciclos_var.get()
                if total > 0:
                    self.progresso_contador.configure(text=f"{ciclo_atual}/{total}")
                    self.progresso_bar.set(ciclo_atual / total)
                else:
                    self.progresso_contador.configure(text="0/0")
                    self.progresso_bar.set(0)
        self.after(0, update)

    def _safe_configure_buttons(self, iniciar_state, continuar_state):
        self.after(0, lambda: self._unsafe_btns(iniciar_state, continuar_state))

    def _unsafe_btns(self, i_state, c_state):
        self.iniciar_btn.configure(state=i_state)
        self.continuar_btn.configure(state=c_state)

    def setup_automacao_tab(self):
        tab = self.tabview.tab("Automação")
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(4, weight=1) # Log expande

        ctrl_frame = ctk.CTkFrame(tab)
        ctrl_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        ctrl_frame.columnconfigure((1, 3), weight=1) 
        
        # Campos de Cidades (1 a 5)
        ctk.CTkLabel(ctrl_frame, text="Cidade 1 (Principal):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.cidades_combobox_1 = ctk.CTkComboBox(ctrl_frame, width=200, command=self.atualizar_listas_exclusivas)
        self.cidades_combobox_1.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        ctk.CTkLabel(ctrl_frame, text="Delay(s):").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.delay_combobox = ctk.CTkComboBox(ctrl_frame, values=["0.05", "0.1", "0.2", "0.5", "1.0", "1.5", "2.0"], width=80)
        self.delay_combobox.set("0.5")
        self.delay_combobox.grid(row=0, column=3, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(ctrl_frame, text="Cidade 2 (Opcional):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.cidades_combobox_2 = ctk.CTkComboBox(ctrl_frame, width=200, command=self.atualizar_listas_exclusivas)
        self.cidades_combobox_2.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(ctrl_frame, text="Status:").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.status_lbl = ctk.CTkLabel(ctrl_frame, textvariable=self.status_text, text_color="white", corner_radius=5, width=100)
        self.status_lbl.grid(row=1, column=3, padx=5, pady=5, sticky="ew")
        self.status_color.trace_add("write", lambda *args: self.status_lbl.configure(fg_color=self.status_color.get()))
        self.status_color.set("gray") 

        ctk.CTkLabel(ctrl_frame, text="Cidade 3 (Opcional):").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.cidades_combobox_3 = ctk.CTkComboBox(ctrl_frame, width=200, command=self.atualizar_listas_exclusivas)
        self.cidades_combobox_3.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        
        ctk.CTkLabel(ctrl_frame, text="Backlog:").grid(row=2, column=2, padx=5, pady=5, sticky="e")
        self.backlog_combobox = ctk.CTkComboBox(ctrl_frame, values=["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "12", "14", "18", "26"], width=80)
        self.backlog_combobox.set("1") 
        self.backlog_combobox.grid(row=2, column=3, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(ctrl_frame, text="Cidade 4 (Opcional):").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.cidades_combobox_4 = ctk.CTkComboBox(ctrl_frame, width=200, command=self.atualizar_listas_exclusivas)
        self.cidades_combobox_4.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(ctrl_frame, text="Cidade 5 (Opcional):").grid(row=4, column=0, padx=5, pady=5, sticky="e")
        self.cidades_combobox_5 = ctk.CTkComboBox(ctrl_frame, width=200, command=self.atualizar_listas_exclusivas)
        self.cidades_combobox_5.grid(row=4, column=1, padx=5, pady=5, sticky="ew")

        # --- BOTÕES DE AÇÃO PRINCIPAL ---
        btn_frame = ctk.CTkFrame(tab, fg_color="transparent")
        btn_frame.grid(row=1, column=0, padx=10, pady=0, sticky="ew")
        
        # INICIAR: Texto Branco
        self.iniciar_btn = ctk.CTkButton(
            btn_frame, text="INICIAR", command=self.iniciar_automacao, 
            height=BTN_HEIGHT_MAIN, text_color="white",
            fg_color=COLOR_SUCCESS, hover_color=COLOR_SUCCESS_HOVER
        )
        self.iniciar_btn.pack(side="left", fill="x", expand=True, padx=5)
        
        # CONTINUAR: Texto Branco (Azul Real)
        self.continuar_btn = ctk.CTkButton(
            btn_frame, text="CONTINUAR", command=self.continuar_automacao, state="disabled", 
            height=BTN_HEIGHT_MAIN, text_color="white",
            fg_color=COLOR_WARNING, hover_color=COLOR_WARNING_HOVER
        )
        self.continuar_btn.pack(side="left", fill="x", expand=True, padx=5)

        # EXCEL: Texto Branco (Agora TEAL)
        self.abrir_btn = ctk.CTkButton(
            btn_frame, text="Excel", command=self.abrir_excel, width=80, 
            height=BTN_HEIGHT_DEFAULT, text_color="white",
            fg_color=COLOR_NEUTRAL, hover_color=COLOR_NEUTRAL_HOVER
        )
        self.abrir_btn.pack(side="right", padx=5)

        # Progresso
        prog_frame = ctk.CTkFrame(tab, fg_color="transparent")
        prog_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        self.progresso_bar = ctk.CTkProgressBar(prog_frame)
        self.progresso_bar.pack(side="left", fill="x", expand=True, padx=(5, 10))
        self.progresso_bar.set(0)
        self.progresso_contador = ctk.CTkLabel(prog_frame, text="0/0", font=("Arial", 12, "bold"))
        self.progresso_contador.pack(side="right", padx=5)

        # Log
        ctk.CTkLabel(tab, text="Log de Execução:", anchor="w").grid(row=3, column=0, padx=10, pady=(10, 0), sticky="nw")
        self.log_textbox = ReadOnlyTextbox(tab) 
        self.log_textbox.grid(row=4, column=0, padx=10, pady=(0, 10), sticky="nsew")

    def setup_cadastro_tab(self):
        tab = self.tabview.tab("Cadastro")
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1) 

        # Add Frame
        add_frame = ctk.CTkFrame(tab)
        add_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        self.nova_cidade_entry = ctk.CTkEntry(add_frame, placeholder_text="Nome da Cidade", height=35)
        self.nova_cidade_entry.pack(side="left", fill="x", expand=True, padx=10, pady=10)
        self.nova_cidade_entry.bind('<Return>', lambda event: self.add_cidade_ui())
        
        # ADICIONAR: Texto Branco
        ctk.CTkButton(
            add_frame, text="Adicionar", width=100, height=BTN_HEIGHT_DEFAULT,
            fg_color=COLOR_SUCCESS, hover_color=COLOR_SUCCESS_HOVER,
            text_color="white",
            command=self.add_cidade_ui
        ).pack(side="right", padx=10, pady=10)

        # List
        self.lista_cidades_txt = ReadOnlyTextbox(tab) 
        self.lista_cidades_txt.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")

        # Del Frame
        del_frame = ctk.CTkFrame(tab)
        del_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        self.del_id_entry = ctk.CTkEntry(del_frame, placeholder_text="ID para Excluir", height=35)
        self.del_id_entry.pack(side="left", fill="x", expand=True, padx=10, pady=10)
        self.del_id_entry.bind('<Return>', lambda event: self.del_cidade_ui())
        
        # EXCLUIR: Texto Branco
        ctk.CTkButton(
            del_frame, text="Excluir", width=100, height=BTN_HEIGHT_DEFAULT,
            fg_color=COLOR_DANGER, hover_color=COLOR_DANGER_HOVER,
            text_color="white",
            command=self.del_cidade_ui
        ).pack(side="right", padx=10, pady=10)

    def setup_usuarios_tab(self):
        tab = self.tabview.tab("Usuários")
        tab.columnconfigure(0, weight=1)
        tab.columnconfigure(1, weight=1)
        tab.rowconfigure(0, weight=1)
        
        # --- COLUNA ESQUERDA: CADASTRO ---
        left_frame = ctk.CTkFrame(tab)
        left_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        
        ctk.CTkLabel(left_frame, text="Novo Operador", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(20, 10))
        
        self.new_username_entry = ctk.CTkEntry(left_frame, placeholder_text="Usuário (Ex: operador)", height=35)
        self.new_username_entry.pack(pady=10, padx=20, fill="x")

        self.new_password_entry = ctk.CTkEntry(left_frame, show="*", placeholder_text="Senha", height=35)
        self.new_password_entry.pack(pady=10, padx=20, fill="x")

        # CADASTRAR: Texto Branco
        ctk.CTkButton(
            left_frame, text="Cadastrar", height=BTN_HEIGHT_DEFAULT,
            fg_color=COLOR_SUCCESS, hover_color=COLOR_SUCCESS_HOVER,
            text_color="white",
            command=self.add_usuario_ui
        ).pack(pady=20, padx=20, fill="x")
        
        # --- COLUNA DIREITA: LISTAGEM ---
        right_frame = ctk.CTkFrame(tab)
        right_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        right_frame.rowconfigure(1, weight=1) 
        right_frame.columnconfigure(0, weight=1)

        ctk.CTkLabel(right_frame, text="Usuários Ativos", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, pady=(20, 10))

        self.lista_usuarios_txt = ReadOnlyTextbox(right_frame)
        self.lista_usuarios_txt.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")
        
        # Área de Exclusão
        del_user_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        del_user_frame.grid(row=2, column=0, pady=10, padx=10, sticky="ew")
        
        self.del_user_entry = ctk.CTkEntry(del_user_frame, placeholder_text="Digite o usuário para excluir", height=35)
        self.del_user_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        # REMOVER: Texto Branco
        ctk.CTkButton(
            del_user_frame, text="Remover", width=80, height=BTN_HEIGHT_DEFAULT,
            fg_color=COLOR_DANGER, hover_color=COLOR_DANGER_HOVER,
            text_color="white",
            command=self.del_usuario_ui
        ).pack(side="right")
        
        self.carregar_lista_usuarios()

    # --- LÓGICA DE NEGÓCIO DA GUI ---

    def carregar_lista_usuarios(self):
        users = listar_usuarios()
        self.lista_usuarios_txt.delete("0.0", "end")
        if not users:
            self.lista_usuarios_txt.insert("end", "Nenhum usuário encontrado.")
        else:
            for u in users:
                prefixo = "👑 " if u == "admin" else "👤 "
                self.lista_usuarios_txt.insert("end", f"{prefixo}{u}\n")

    def add_usuario_ui(self):
        user = self.new_username_entry.get()
        pwd = self.new_password_entry.get()
        msg, ok = adicionar_usuario(user, pwd)
        if ok:
            self.new_username_entry.delete(0, "end")
            self.new_password_entry.delete(0, "end")
            self.carregar_lista_usuarios()
        exibir_popup(title="Cadastro", message=msg, icon="check" if ok else "cancel")

    def del_usuario_ui(self):
        user_to_del = self.del_user_entry.get().strip()
        if not user_to_del:
            exibir_popup("Aviso", "Digite o nome do usuário para excluir.", "warning")
            return
        if exibir_confirmacao("Confirmar Exclusão", f"Tem certeza que deseja remover o usuário '{user_to_del}'?", "warning", option_ok="Excluir") == "Excluir":
            msg, ok = excluir_usuario(user_to_del)
            if ok:
                self.del_user_entry.delete(0, "end")
                self.carregar_lista_usuarios()
            exibir_popup("Exclusão", msg, "check" if ok else "cancel")

    def atualizar_listas_exclusivas(self, choice=None):
        sel1 = self.cidades_combobox_1.get()
        sel2 = self.cidades_combobox_2.get()
        sel3 = self.cidades_combobox_3.get()
        sel4 = self.cidades_combobox_4.get()
        sel5 = self.cidades_combobox_5.get()
        ignorar = ["NENHUM FILTRO", "SELECIONE", "", None]
        def get_opcoes(excluir_list):
            excluir_limpa = [x for x in excluir_list if x not in ignorar]
            return ["NENHUM FILTRO"] + [c for c in self.todas_cidades if c not in excluir_limpa]
        self.cidades_combobox_1.configure(values=get_opcoes([sel2, sel3, sel4, sel5]))
        self.cidades_combobox_1.set(sel1)
        self.cidades_combobox_2.configure(values=get_opcoes([sel1, sel3, sel4, sel5]))
        self.cidades_combobox_2.set(sel2)
        self.cidades_combobox_3.configure(values=get_opcoes([sel1, sel2, sel4, sel5]))
        self.cidades_combobox_3.set(sel3)
        self.cidades_combobox_4.configure(values=get_opcoes([sel1, sel2, sel3, sel5]))
        self.cidades_combobox_4.set(sel4)
        self.cidades_combobox_5.configure(values=get_opcoes([sel1, sel2, sel3, sel4]))
        self.cidades_combobox_5.set(sel5)

    def carregar_cidades_db(self):
        nomes = buscar_nomes_cidades()
        self.todas_cidades = nomes 
        opcoes_padrao = ["NENHUM FILTRO"] + nomes
        self.cidades_combobox_1.configure(values=opcoes_padrao)
        self.cidades_combobox_1.set("NENHUM FILTRO")
        self.cidades_combobox_2.configure(values=opcoes_padrao)
        self.cidades_combobox_2.set("NENHUM FILTRO")
        self.cidades_combobox_3.configure(values=opcoes_padrao)
        self.cidades_combobox_3.set("NENHUM FILTRO")
        self.cidades_combobox_4.configure(values=opcoes_padrao)
        self.cidades_combobox_4.set("NENHUM FILTRO")
        self.cidades_combobox_5.configure(values=opcoes_padrao)
        self.cidades_combobox_5.set("NENHUM FILTRO")
        lista = listar_cidades()
        self.lista_cidades_txt.delete("0.0", "end")
        if not lista:
            self.lista_cidades_txt.insert("end", "Nenhuma cidade cadastrada.\n")
        for id_c, nome in lista:
            self.lista_cidades_txt.insert("end", f"[{id_c}] {nome}\n")

    def iniciar_automacao(self):
        try:
            nome_busca = utils.NOME_ARQUIVO_ALVO.split('.')[0] 
            janelas_encontradas = pyautogui.getWindowsWithTitle(nome_busca)
            planilha_detectada = any(nome_busca.lower() in str(j.title).lower() for j in janelas_encontradas)
            if not planilha_detectada:
                 exibir_popup(title="Planilha Não Detectada", message=f"⚠️ A automação requer que a planilha '{utils.NOME_ARQUIVO_ALVO}' esteja aberta.\n\nPor favor, abra o arquivo Excel e tente novamente.", icon="cancel")
                 return
            c1 = self.cidades_combobox_1.get()
            c2 = self.cidades_combobox_2.get()
            c3 = self.cidades_combobox_3.get()
            c4 = self.cidades_combobox_4.get()
            c5 = self.cidades_combobox_5.get()
            backlog_val = self.backlog_combobox.get()
            delay_str = self.delay_combobox.get()
            def limpar_valor(val):
                if val and val.upper() not in ["NENHUM FILTRO", "SELECIONE", ""]: return val
                return None
            if not limpar_valor(c1):
                exibir_popup(title="Campo Obrigatório", message="⚠️ É obrigatório selecionar a 'Cidade 1 (Principal)' para iniciar a automação.", icon="warning")
                return 
            cidades_validas = []
            if limpar_valor(c1): cidades_validas.append(limpar_valor(c1))
            if limpar_valor(c2): cidades_validas.append(limpar_valor(c2))
            if limpar_valor(c3): cidades_validas.append(limpar_valor(c3))
            if limpar_valor(c4): cidades_validas.append(limpar_valor(c4))
            if limpar_valor(c5): cidades_validas.append(limpar_valor(c5))
            cidade_final = ",".join(cidades_validas)
            msg_filtro = f"para: {cidade_final}"
            delay = utils.validar_e_obter_delay(delay_str)
            utils.DELAY_ATUAL = delay 
            utils.INDICE_ATUAL_DO_CICLO = 0
            utils.PARAR_AUTOMACAO = False
            self.log_textbox.delete("1.0", "end")
            self._safe_update_gui(status="Rodando", total_ciclos=0, ciclo_atual=0)
            self._safe_configure_buttons("disabled", "disabled")
            timestamp = time.strftime("%H:%M:%S")
            self.log_textbox.insert("end", f"[{timestamp}] Iniciando {msg_filtro} | Backlog: {backlog_val} (Delay: {delay}s)...\n")
            if not self.monitor_thread_started:
                t = threading.Thread(target=utils.monitorar_tecla_escape, args=(self.log_textbox,), daemon=True)
                t.start()
                self.monitor_thread_started = True
            t_core = threading.Thread(target=automacao_core, args=(self.log_textbox, cidade_final, backlog_val, delay, self._safe_configure_buttons, self._safe_update_gui))
            t_core.start()
        except Exception as e:
            timestamp = time.strftime("%H:%M:%S")
            self.log_textbox.insert("end", f"[{timestamp}] ❌ Erro ao iniciar: {e}\n")
            self._safe_update_gui(status="Erro")
            self._safe_configure_buttons("normal", "disabled")

    def continuar_automacao(self):
        if not utils.PARAR_AUTOMACAO: return
        try:
            delay_str = self.delay_combobox.get()
            novo_delay = utils.validar_e_obter_delay(delay_str)
            utils.DELAY_ATUAL = novo_delay 
            utils.PARAR_AUTOMACAO = False
            self._safe_update_gui(status="Rodando")
            self._safe_configure_buttons("disabled", "disabled")
            timestamp = time.strftime("%H:%M:%S")
            self.log_textbox.insert("end", f"[{timestamp}] ▶ Retomando com Delay ajustado para: {novo_delay}s\n")
        except Exception as e:
            self.log_textbox.insert("end", f"❌ Erro ao continuar: {e}\n")

    def add_cidade_ui(self):
        self.focus()
        msg, ok = adicionar_cidade(self.nova_cidade_entry.get())
        if ok: 
            self.nova_cidade_entry.delete(0, "end")
            self.carregar_cidades_db()
        exibir_popup(title="Cadastro", message=msg, icon="check" if ok else "cancel")
        self.nova_cidade_entry.focus_set()
        
    def del_cidade_ui(self):
        self.focus()
        id_str = self.del_id_entry.get().strip()
        if not id_str:
             exibir_popup("Aviso", "Por favor, informe o ID da cidade para excluir.", "warning")
             self.del_id_entry.focus_set()
             return
        nome = buscar_nome_cidade_por_id(id_str)
        if not nome:
             exibir_popup("Erro", f"Não foi encontrada nenhuma cidade com o ID '{id_str}'.", "cancel")
             self.del_id_entry.focus_set()
             return
        if exibir_confirmacao(title="Confirmar Exclusão", message=f"Tem certeza que deseja excluir a cidade abaixo?\n\nID: {id_str}\nNome: {nome}\n\nEssa ação não pode ser desfeita.", icon="question", option_ok="Excluir", option_cancel="Cancelar") == "Excluir":
            msg, ok = excluir_cidade(id_str)
            if ok:
                self.del_id_entry.delete(0, "end")
                self.carregar_cidades_db()
            exibir_popup(title="Exclusão", message=msg, icon="check" if ok else "cancel")
        self.del_id_entry.focus_set()

    def abrir_excel(self):
        ok, msg = utils.abrir_planilha_alvo()
        timestamp = time.strftime("%H:%M:%S")
        self.log_textbox.insert("end", f"[{timestamp}] {msg}\n")