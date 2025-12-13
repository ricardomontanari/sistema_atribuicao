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

# --- CONSTANTES DE CONFIGURAÇÃO E ESTILO ---
CONFIG_LOGIN_FILE = "login_config.json"

# Dimensões
BTN_HEIGHT_DEFAULT = 35
BTN_HEIGHT_MAIN = 40

# Paleta de Cores
COLOR_SUCCESS = "green"           # Botões: Iniciar, Adicionar, Salvar
COLOR_SUCCESS_HOVER = "darkgreen"

COLOR_DANGER = "red"              # Botões: Excluir, Parar, Remover
COLOR_DANGER_HOVER = "darkred"

COLOR_WARNING = "#1565C0"         # Botão: CONTINUAR (Azul Real)
COLOR_WARNING_HOVER = "#0D47A1"

COLOR_NEUTRAL = "#00796B"         # Botão: EXCEL (Teal/Verde-petróleo)
COLOR_NEUTRAL_HOVER = "#004D40"

# NOVA COR: Cinza Claro para Limpar (Lighter Gray for Clear Button)
# ATUALIZADO: Usando um tom de cinza mais claro (Prata) para melhor visibilidade no tema escuro.
COLOR_CLEAR = "#A0A0A0" 
COLOR_CLEAR_HOVER = "#808080"

# ####################################################################
# --- COMPONENTES PERSONALIZADOS ---
# ####################################################################

class ReadOnlyTextbox(ctk.CTkTextbox):
    """
    Caixa de texto somente leitura para logs e listagens.
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
# --- FUNÇÕES AUXILIARES DE POPUP ---
# ####################################################################

def exibir_popup(title, message, icon="info"):
    """Exibe um popup simples de informação/erro."""
    msg_box = CTkMessagebox(title=title, message=message, icon=icon)
    msg_box.get()

def exibir_confirmacao(title, message, icon="question", option_ok="Confirmar", option_cancel="Cancelar"):
    """Exibe um popup de confirmação e retorna a opção escolhida."""
    msg_box = CTkMessagebox(title=title, message=message, icon=icon, 
                            option_1=option_cancel, option_2=option_ok)
    return msg_box.get()

# ####################################################################
# --- CLASSE 1: JANELA DE LOGIN ---
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
            main_app = App() # RENOMEADO: Chama a classe App
            main_app.mainloop()
            self.destroy()
        else:
            exibir_popup(title="Erro de Acesso", message="Usuário ou Senha inválidos.", icon="cancel")

# ####################################################################
# --- CLASSE 2: APLICAÇÃO PRINCIPAL (APP) ---
# ####################################################################

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configuração da Janela
        self.title(f"Atribuidor ({utils.VERSAO_SISTEMA})")
        
        # --- DEFINIÇÃO DE TAMANHOS ---
        self.geo_login = (400, 450) # Minimalista para Login
        self.geo_main = (650, 650)  # Tamanho principal
        
        # Inicia com o tamanho minimalista centralizado
        self.ajustar_geometria(*self.geo_login)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Inicializa o Banco de Dados
        setup_database()

        # Variáveis de Estado da Automação (GUI)
        self.status_text = ctk.StringVar(value="AGUARDANDO")
        self.status_color = ctk.StringVar(value="gray")
        self.total_de_ciclos_var = ctk.IntVar(value=0)
        self.monitor_thread_started = False
        self.todas_cidades = []

        # Inicia pela tela de Login
        self.construir_tela_login()

    def ajustar_geometria(self, width, height):
        """Calcula a posição central e redimensiona a janela."""
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        pos_x = (screen_width // 2) - (width // 2)
        pos_y = (screen_height // 2) - (height // 2)
        
        self.geometry(f"{width}x{height}+{pos_x}+{pos_y}")

    # #################################################################
    # --- TELA DE LOGIN ---
    # #################################################################

    def construir_tela_login(self):
        """Constrói a interface de login."""
        self.login_frame = ctk.CTkFrame(self)
        self.login_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.login_frame.grid_columnconfigure(0, weight=1)
        self.login_frame.grid_rowconfigure(0, weight=1) # Centralizar verticalmente

        inner_frame = ctk.CTkFrame(self.login_frame, fg_color="transparent")
        inner_frame.grid(row=0, column=0)

        ctk.CTkLabel(inner_frame, text="Acesso ao Atribuidor", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(0, 25))

        self.username_entry = ctk.CTkEntry(inner_frame, placeholder_text="Usuário", width=220, height=35)
        self.username_entry.pack(pady=10)

        self.password_entry = ctk.CTkEntry(inner_frame, placeholder_text="Senha", show="*", width=220, height=35)
        self.password_entry.pack(pady=10)

        self.lembrar_var = ctk.BooleanVar(value=False)
        self.lembrar_checkbox = ctk.CTkCheckBox(inner_frame, text="Lembrar usuário", variable=self.lembrar_var)
        self.lembrar_checkbox.pack(pady=10)

        ctk.CTkButton(
            inner_frame, text="ENTRAR", command=self.realizar_login, width=220, height=35,
            fg_color=COLOR_NEUTRAL, hover_color=COLOR_NEUTRAL_HOVER
        ).pack(pady=20)

        self.bind('<Return>', lambda event: self.realizar_login())
        self.carregar_usuario_salvo()
        
        # Foco inicial
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
            except: pass

    def realizar_login(self):
        user = self.username_entry.get()
        pwd = self.password_entry.get()
        
        if verificar_credenciais(user, pwd):
            # Salva preferência
            try:
                data = {"last_user": user if self.lembrar_var.get() else ""}
                with open(CONFIG_LOGIN_FILE, 'w') as f: json.dump(data, f)
            except: pass
            
            # Remove tela de login e ajusta tamanho
            self.login_frame.destroy()
            self.unbind('<Return>')
            self.ajustar_geometria(*self.geo_main)
            self.construir_tela_principal()
        else:
            exibir_popup("Acesso Negado", "Usuário ou senha incorretos.", "cancel")

    # #################################################################
    # --- TELA PRINCIPAL (APÓS LOGIN) ---
    # #################################################################

    def construir_tela_principal(self):
        """Constrói a interface principal com abas."""
        # Configurações de Layout
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        self.tabview.add("Automação")
        self.tabview.add("Cadastro")
        self.tabview.add("Usuários")

        self.setup_tab_automacao()
        self.setup_tab_cadastro()
        self.setup_tab_usuarios()
        
        # Carrega dados iniciais
        self.carregar_cidades_db()

    def limpar_combobox_cidade(self, combobox_instance):
        """Define o valor do combobox para 'NENHUM FILTRO' e atualiza exclusivas."""
        combobox_instance.set("NENHUM FILTRO")
        self.atualizar_listas_exclusivas()

    # --- ABA: AUTOMAÇÃO ---
    def setup_tab_automacao(self):
        tab = self.tabview.tab("Automação")
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(5, weight=1) # Log expande (row 4 mudou para 5)

        # Painel de Controles
        ctrl_frame = ctk.CTkFrame(tab)
        ctrl_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        # Colunas: 0 (Label), 1 (ComboBox), 2 (Botão Limpar), 3 (Label), 4 (Valor)
        ctrl_frame.columnconfigure((1, 4), weight=1)

        # Função auxiliar para criar Combo + Botão Limpar
        def create_city_input(frame, row, label_text):
            ctk.CTkLabel(frame, text=label_text).grid(row=row, column=0, padx=5, pady=5, sticky="e")
            combo = ctk.CTkComboBox(frame, width=180, command=self.atualizar_listas_exclusivas)
            combo.grid(row=row, column=1, padx=(5, 0), pady=5, sticky="ew")
            
            # Botão Limpar
            btn_clear = ctk.CTkButton(
                frame, text="X", width=25, height=25,
                fg_color=COLOR_CLEAR, hover_color=COLOR_CLEAR_HOVER,
                command=lambda cb=combo: self.limpar_combobox_cidade(cb)
            )
            btn_clear.grid(row=row, column=2, padx=(5, 10), pady=5, sticky="w")
            
            return combo

        # Cidades (Linhas 0 a 4)
        self.cidades_combobox_1 = create_city_input(ctrl_frame, 0, "Cidade Principal:")
        self.cidades_combobox_2 = create_city_input(ctrl_frame, 1, "Cidade 2 (Opc):")
        self.cidades_combobox_3 = create_city_input(ctrl_frame, 2, "Cidade 3 (Opc):")
        self.cidades_combobox_4 = create_city_input(ctrl_frame, 3, "Cidade 4 (Opc):")
        self.cidades_combobox_5 = create_city_input(ctrl_frame, 4, "Cidade 5 (Opc):")

        # Controles (Colunas 3 e 4)

        # Delay
        ctk.CTkLabel(ctrl_frame, text="Delay (s):").grid(row=0, column=3, padx=5, pady=5, sticky="e")
        self.delay_combobox = ctk.CTkComboBox(ctrl_frame, values=["0.05", "0.1", "0.2", "0.5", "1.0", "1.5", "2.0"], width=80)
        self.delay_combobox.set("0.2")
        self.delay_combobox.grid(row=0, column=4, padx=10, pady=5, sticky="ew")

        # Status
        ctk.CTkLabel(ctrl_frame, text="Status:").grid(row=1, column=3, padx=5, pady=5, sticky="e")
        self.status_lbl = ctk.CTkLabel(ctrl_frame, textvariable=self.status_text, text_color="white", corner_radius=5, width=100)
        self.status_lbl.grid(row=1, column=4, padx=10, pady=5, sticky="ew")
        self.status_color.trace_add("write", lambda *args: self.status_lbl.configure(fg_color=self.status_color.get()))
        self.status_color.set("gray")

        # Backlog
        ctk.CTkLabel(ctrl_frame, text="Backlog:").grid(row=2, column=3, padx=5, pady=5, sticky="e")
        self.backlog_combobox = ctk.CTkComboBox(ctrl_frame, values=["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "12", "14", "18", "26"], width=80)
        self.backlog_combobox.set("1")
        self.backlog_combobox.grid(row=2, column=4, padx=10, pady=5, sticky="ew")

        # Linhas vazias para manter o espaçamento uniforme
        # Removido (Não necessário com o grid unificado)

        # Botões de Ação
        btn_frame = ctk.CTkFrame(tab, fg_color="transparent")
        btn_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        
        self.iniciar_btn = ctk.CTkButton(
            btn_frame, text="INICIAR", command=self.iniciar_automacao, 
            height=BTN_HEIGHT_MAIN, fg_color=COLOR_SUCCESS, hover_color=COLOR_SUCCESS_HOVER
        )
        self.iniciar_btn.pack(side="left", fill="x", expand=True, padx=5)
        
        self.continuar_btn = ctk.CTkButton(
            btn_frame, text="CONTINUAR", command=self.continuar_automacao, state="disabled", 
            height=BTN_HEIGHT_MAIN, fg_color=COLOR_WARNING, hover_color=COLOR_WARNING_HOVER
        )
        self.continuar_btn.pack(side="left", fill="x", expand=True, padx=5)

        self.abrir_btn = ctk.CTkButton(
            btn_frame, text="ABRIR EXCEL", command=self.abrir_excel, width=100, 
            height=BTN_HEIGHT_MAIN, fg_color=COLOR_NEUTRAL, hover_color=COLOR_NEUTRAL_HOVER
        )
        self.abrir_btn.pack(side="right", padx=5)

        # Barra de Progresso
        prog_frame = ctk.CTkFrame(tab, fg_color="transparent")
        prog_frame.grid(row=2, column=0, padx=10, pady=5, sticky="ew")
        
        self.progresso_bar = ctk.CTkProgressBar(prog_frame)
        self.progresso_bar.pack(side="left", fill="x", expand=True, padx=(5, 10))
        self.progresso_bar.set(0)
        
        self.progresso_contador = ctk.CTkLabel(prog_frame, text="0/0", font=("Arial", 12, "bold"))
        self.progresso_contador.pack(side="right", padx=5)

        # Log
        ctk.CTkLabel(tab, text="Log de Execução:", anchor="w").grid(row=3, column=0, padx=10, pady=(5, 0), sticky="nw")
        self.log_textbox = ReadOnlyTextbox(tab)
        self.log_textbox.grid(row=4, column=0, padx=10, pady=(0, 10), sticky="nsew") # Linha 4 na aba automação

    # --- ABA: CADASTRO ---
    def setup_tab_cadastro(self):
        tab = self.tabview.tab("Cadastro")
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)

        # Adicionar Cidade
        add_frame = ctk.CTkFrame(tab)
        add_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        
        self.nova_cidade_entry = ctk.CTkEntry(add_frame, placeholder_text="Nome da Cidade", height=BTN_HEIGHT_DEFAULT)
        self.nova_cidade_entry.pack(side="left", fill="x", expand=True, padx=10, pady=10)
        
        ctk.CTkButton(
            add_frame, text="Adicionar", width=100, height=BTN_HEIGHT_DEFAULT,
            fg_color=COLOR_SUCCESS, hover_color=COLOR_SUCCESS_HOVER,
            command=self.add_cidade_ui
        ).pack(side="right", padx=10)

        # Lista de Cidades
        self.lista_cidades_txt = ReadOnlyTextbox(tab)
        self.lista_cidades_txt.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")

        # Excluir Cidade
        del_frame = ctk.CTkFrame(tab)
        del_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        
        self.del_id_entry = ctk.CTkEntry(del_frame, placeholder_text="ID para Excluir", height=BTN_HEIGHT_DEFAULT)
        self.del_id_entry.pack(side="left", fill="x", expand=True, padx=10, pady=10)
        
        ctk.CTkButton(
            del_frame, text="Excluir", width=100, height=BTN_HEIGHT_DEFAULT,
            fg_color=COLOR_DANGER, hover_color=COLOR_DANGER_HOVER,
            command=self.del_cidade_ui
        ).pack(side="right", padx=10)

    # --- ABA: USUÁRIOS ---
    def setup_tab_usuarios(self):
        tab = self.tabview.tab("Usuários")
        tab.columnconfigure(0, weight=1)
        tab.columnconfigure(1, weight=1)
        tab.rowconfigure(0, weight=1)
        
        # Painel Esquerdo: Cadastro
        left_frame = ctk.CTkFrame(tab)
        left_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        
        ctk.CTkLabel(left_frame, text="Novo Operador", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(20, 10))
        
        self.new_username_entry = ctk.CTkEntry(left_frame, placeholder_text="Usuário", height=BTN_HEIGHT_DEFAULT)
        self.new_username_entry.pack(pady=10, padx=20, fill="x")

        self.new_password_entry = ctk.CTkEntry(left_frame, show="*", placeholder_text="Senha", height=BTN_HEIGHT_DEFAULT)
        self.new_password_entry.pack(pady=10, padx=20, fill="x")

        ctk.CTkButton(
            left_frame, text="CADASTRAR", height=BTN_HEIGHT_DEFAULT,
            fg_color=COLOR_SUCCESS, hover_color=COLOR_SUCCESS_HOVER,
            command=self.add_usuario_ui
        ).pack(pady=20, padx=20, fill="x")
        
        # Painel Direito: Listagem
        right_frame = ctk.CTkFrame(tab)
        right_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        right_frame.rowconfigure(1, weight=1) 
        right_frame.columnconfigure(0, weight=1)

        ctk.CTkLabel(right_frame, text="Usuários Ativos", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, pady=(20, 10))

        self.lista_usuarios_txt = ReadOnlyTextbox(right_frame)
        self.lista_usuarios_txt.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")
        
        # Exclusão de Usuário
        del_user_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        del_user_frame.grid(row=2, column=0, pady=10, padx=10, sticky="ew")
        
        self.del_user_entry = ctk.CTkEntry(del_user_frame, placeholder_text="Nome para remover", height=BTN_HEIGHT_DEFAULT)
        self.del_user_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        ctk.CTkButton(
            del_user_frame, text="REMOVER", width=80, height=BTN_HEIGHT_DEFAULT,
            fg_color=COLOR_DANGER, hover_color=COLOR_DANGER_HOVER,
            command=self.del_usuario_ui
        ).pack(side="right")
        
        self.carregar_lista_usuarios()

    # #################################################################
    # --- MÉTODOS DE LÓGICA DA INTERFACE ---
    # #################################################################

    def _safe_update_gui(self, status=None, total_ciclos=None, ciclo_atual=None):
        """Atualiza a GUI de forma segura a partir de threads."""
        def update():
            if status:
                self.status_text.set(status.upper())
                # Adicionado "Parado" com cor vermelha
                colors = {"Rodando": "blue", "Pausado": "orange", "Erro": "red", "Finalizado": "green", "Parado": "red"}
                self.status_color.set(colors.get(status, "gray"))
            
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
        # LÓGICA DO BOTÃO PARAR:
        # Se o botão Continuar estiver ativo ('normal'), significa que estamos em PAUSA.
        if c_state == "normal":
            self.iniciar_btn.configure(text="PARAR", fg_color=COLOR_DANGER, hover_color=COLOR_DANGER_HOVER, 
                                       command=self.parar_automacao, state="normal")
        else:
            # Caso contrário (Rodando ou Parado), volta a ser 'INICIAR'
            self.iniciar_btn.configure(text="INICIAR", fg_color=COLOR_SUCCESS, hover_color=COLOR_SUCCESS_HOVER,
                                       command=self.iniciar_automacao, state=i_state)
        
        self.continuar_btn.configure(state=c_state)

    # --- LÓGICA DE AÇÃO ---

    def parar_automacao(self):
        """Função chamada pelo botão PARAR (Vermelho)."""
        self.log_textbox.insert("end", f"[{time.strftime('%H:%M:%S')}] 🛑 Solicitando parada forçada...\n")
        utils.CANCELAR_AUTOMACAO = True
        utils.PARAR_AUTOMACAO = False # Destrava o loop de pausa para permitir o cancelamento
        self.iniciar_btn.configure(state="disabled") # Evita clique duplo enquanto processa o cancelamento

    def iniciar_automacao(self):
        try:
            # Verifica Planilha Aberta
            nome_busca = utils.NOME_ARQUIVO_ALVO.split('.')[0] 
            janelas = pyautogui.getWindowsWithTitle(nome_busca)
            if not any(nome_busca.lower() in str(j.title).lower() for j in janelas):
                 exibir_popup("Erro", f"A planilha '{utils.NOME_ARQUIVO_ALVO}' precisa estar aberta.", "cancel")
                 return

            # Coleta Cidades
            cidades = []
            for cb in [self.cidades_combobox_1, self.cidades_combobox_2, self.cidades_combobox_3,
                       self.cidades_combobox_4, self.cidades_combobox_5]:
                val = cb.get()
                if val and val not in ["NENHUM FILTRO", "SELECIONE", ""]:
                    cidades.append(val)
            
            if not cidades:
                exibir_popup("Aviso", "Selecione pelo menos a Cidade Principal.", "warning")
                return 

            cidade_final = ",".join(cidades)
            backlog = self.backlog_combobox.get()
            delay = utils.validar_e_obter_delay(self.delay_combobox.get())
            
            # Configura Variáveis Globais
            utils.DELAY_ATUAL = delay 
            utils.INDICE_ATUAL_DO_CICLO = 0
            utils.PARAR_AUTOMACAO = False
            utils.CANCELAR_AUTOMACAO = False # Reseta flag de cancelamento
            
            # Reseta UI
            self.log_textbox.delete("1.0", "end")
            self._safe_update_gui(status="Rodando", total_ciclos=0, ciclo_atual=0)
            self._safe_configure_buttons("disabled", "disabled")
            
            self.log_textbox.insert("end", f"[{time.strftime('%H:%M:%S')}] Iniciando para: {cidade_final}\n")
            
            # Inicia Monitoramento ESC
            if not self.monitor_thread_started:
                t = threading.Thread(target=utils.monitorar_tecla_escape, args=(self.log_textbox,), daemon=True)
                t.start()
                self.monitor_thread_started = True
            
            # Inicia Thread Principal
            t_core = threading.Thread(
                target=automacao_core, 
                args=(self.log_textbox, cidade_final, backlog, delay, self._safe_configure_buttons, self._safe_update_gui)
            )
            t_core.start()
            
        except Exception as e:
            self.log_textbox.insert("end", f"❌ Erro ao iniciar: {e}\n")
            self._safe_update_gui(status="Erro")
            self._safe_configure_buttons("normal", "disabled")

    def continuar_automacao(self):
        if not utils.PARAR_AUTOMACAO: return
        try:
            novo_delay = utils.validar_e_obter_delay(self.delay_combobox.get())
            utils.DELAY_ATUAL = novo_delay 
            utils.PARAR_AUTOMACAO = False
            
            self._safe_update_gui(status="Rodando")
            self._safe_configure_buttons("disabled", "disabled")
            self.log_textbox.insert("end", f"[{time.strftime('%H:%M:%S')}] ▶ Retomando...\n")
        except Exception as e:
            self.log_textbox.insert("end", f"❌ Erro ao continuar: {e}\n")

    # --- CRUD e Outros ---

    def carregar_cidades_db(self):
        self.todas_cidades = buscar_nomes_cidades()
        opcoes = ["NENHUM FILTRO"] + self.todas_cidades
        
        # Atualiza todos os comboboxes
        for cb in [self.cidades_combobox_1, self.cidades_combobox_2, self.cidades_combobox_3, 
                   self.cidades_combobox_4, self.cidades_combobox_5]:
            cb.configure(values=opcoes)
            if cb.get() not in opcoes: cb.set("NENHUM FILTRO")
            
        # Atualiza listagem
        lista = listar_cidades()
        self.lista_cidades_txt.delete("0.0", "end")
        if not lista:
            self.lista_cidades_txt.insert("end", "Nenhuma cidade cadastrada.\n")
        else:
            for id_c, nome in lista:
                self.lista_cidades_txt.insert("end", f"[{id_c}] {nome}\n")

    def carregar_lista_usuarios(self):
        users = listar_usuarios()
        self.lista_usuarios_txt.delete("0.0", "end")
        if not users:
            self.lista_usuarios_txt.insert("end", "Nenhum usuário.\n")
        else:
            for u in users:
                icon = "👑 " if u == "admin" else "👤 "
                self.lista_usuarios_txt.insert("end", f"{icon}{u}\n")

    def atualizar_listas_exclusivas(self, choice=None):
        """Impede selecionar a mesma cidade em múltiplos campos."""
        combos = [self.cidades_combobox_1, self.cidades_combobox_2, self.cidades_combobox_3,
                  self.cidades_combobox_4, self.cidades_combobox_5]
        selecoes = [cb.get() for cb in combos]
        ignorar = ["NENHUM FILTRO", "SELECIONE", "", None]
        
        for i, cb in enumerate(combos):
            atual = cb.get()
            outras_sel = [s for j, s in enumerate(selecoes) if j != i and s not in ignorar]
            opcoes_disp = ["NENHUM FILTRO"] + [c for c in self.todas_cidades if c not in outras_sel]
            cb.configure(values=opcoes_disp)
            # Restaura a seleção se ela ainda for válida, senão reseta
            if atual in opcoes_disp: cb.set(atual)
            else: cb.set("NENHUM FILTRO")

    def add_cidade_ui(self):
        msg, ok = adicionar_cidade(self.nova_cidade_entry.get())
        if ok: 
            self.nova_cidade_entry.delete(0, "end")
            self.carregar_cidades_db()
        exibir_popup("Cadastro", msg, "check" if ok else "cancel")

    def del_cidade_ui(self):
        id_str = self.del_id_entry.get().strip()
        if not id_str: return
        
        nome = buscar_nome_cidade_por_id(id_str)
        if not nome:
             exibir_popup("Erro", "ID não encontrado.", "cancel")
             return
             
        if exibir_confirmacao("Excluir", f"Excluir cidade '{nome}' (ID {id_str})?", "warning") == "Confirmar":
            msg, ok = excluir_cidade(id_str)
            if ok:
                self.del_id_entry.delete(0, "end")
                self.carregar_cidades_db()
            exibir_popup("Sucesso", msg, "check")

    def add_usuario_ui(self):
        msg, ok = adicionar_usuario(self.new_username_entry.get(), self.new_password_entry.get())
        if ok:
            self.new_username_entry.delete(0, "end")
            self.new_password_entry.delete(0, "end")
            self.carregar_lista_usuarios()
        exibir_popup("Cadastro", msg, "check" if ok else "cancel")

    def del_usuario_ui(self):
        user = self.del_user_entry.get().strip()
        if not user: return
        
        if exibir_confirmacao("Excluir", f"Remover usuário '{user}'?", "warning") == "Confirmar":
            msg, ok = excluir_usuario(user)
            if ok:
                self.del_user_entry.delete(0, "end")
                self.carregar_lista_usuarios()
            exibir_popup("Resultado", msg, "check" if ok else "cancel")

    def abrir_excel(self):
        ok, msg = utils.abrir_planilha_alvo()
        self.log_textbox.insert("end", f"[{time.strftime('%H:%M:%S')}] {msg}\n")