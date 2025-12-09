import customtkinter as ctk
import threading 
import utils
from CTkMessagebox import CTkMessagebox
from automation_logic import automacao_core
from db_manager import (
    setup_database, 
    adicionar_cidade, 
    buscar_nomes_cidades,
    listar_cidades,
    excluir_cidade
)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Configuração da Janela Principal ---
        self.title(f"Sistema de Automação ({utils.VERSAO_SISTEMA})")
        self.geometry("650x750")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- Variáveis de Estado da GUI ---
        self.status_text = ctk.StringVar(value="PRONTO")
        self.status_color = ctk.StringVar(value="green")
        self.progresso_atual = ctk.DoubleVar(value=0)
        self.total_de_ciclos_var = ctk.IntVar(value=0)
        
        # Controle da Thread de Monitoramento (ESC)
        self.monitor_thread_started = False

        # --- Inicialização do Banco de Dados ---
        setup_database() 

        # --- Configuração das Abas (Tabview) ---
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.tabview.add("Automação")
        self.tabview.add("Cadastro")
        
        self.setup_automacao_tab()
        self.setup_cadastro_tab()
        
        self.carregar_cidades_db()

    # ####################################################################
    # --- MÉTODOS THREAD-SAFE ---
    # ####################################################################

    def _safe_update_gui(self, status=None, total_ciclos=None, ciclo_atual=None):
        def update():
            if status:
                self.status_text.set(status.upper())
                colors = {
                    "Rodando": "blue", 
                    "Pausado": "orange", 
                    "Erro": "red", 
                    "Finalizado": "green",
                    "PRONTO": "gray"
                }
                self.status_color.set(colors.get(status, self.status_color.get()))

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

    # ####################################################################
    # --- CONFIGURAÇÃO DA INTERFACE ---
    # ####################################################################

    def setup_automacao_tab(self):
        tab = self.tabview.tab("Automação")
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(3, weight=1) 

        # 1. Frame de Controles
        ctrl_frame = ctk.CTkFrame(tab)
        ctrl_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        ctrl_frame.columnconfigure((1, 3), weight=1) 
        
        # Filtro de Cidade
        ctk.CTkLabel(ctrl_frame, text="Cidade:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.cidades_combobox = ctk.CTkComboBox(ctrl_frame, width=200)
        self.cidades_combobox.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        # Seleção de Delay
        ctk.CTkLabel(ctrl_frame, text="Delay(s):").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        # Adicionado "0.05" na lista de opções
        self.delay_combobox = ctk.CTkComboBox(ctrl_frame, values=["0.05", "0.1", "0.2", "0.5", "1.0", "1.5", "2.0"], width=80)
        self.delay_combobox.set("0.5")
        self.delay_combobox.grid(row=0, column=3, padx=5, pady=5, sticky="ew")

        # Status
        ctk.CTkLabel(ctrl_frame, text="Status:").grid(row=0, column=4, padx=5, pady=5, sticky="e")
        self.status_lbl = ctk.CTkLabel(ctrl_frame, textvariable=self.status_text, text_color="white", corner_radius=5, width=100)
        self.status_lbl.grid(row=0, column=5, padx=5, pady=5, sticky="ew")
        self.status_color.trace_add("write", lambda *args: self.status_lbl.configure(fg_color=self.status_color.get()))
        self.status_color.set("gray") 

        # 2. Botões
        btn_frame = ctk.CTkFrame(tab, fg_color="transparent")
        btn_frame.grid(row=1, column=0, padx=10, pady=0, sticky="ew")
        
        self.iniciar_btn = ctk.CTkButton(btn_frame, text="▶ INICIAR", command=self.iniciar_automacao, fg_color="green", hover_color="darkgreen")
        self.iniciar_btn.pack(side="left", fill="x", expand=True, padx=5)
        
        self.continuar_btn = ctk.CTkButton(btn_frame, text="⏯ CONTINUAR", command=self.continuar_automacao, state="disabled", fg_color="orange", text_color="black", hover_color="#FFB74D")
        self.continuar_btn.pack(side="left", fill="x", expand=True, padx=5)

        self.abrir_btn = ctk.CTkButton(btn_frame, text="📂 Excel", command=self.abrir_excel, width=80, fg_color="#3B8ED0", hover_color="#36719F")
        self.abrir_btn.pack(side="right", padx=5)

        # 3. Progresso
        prog_frame = ctk.CTkFrame(tab, fg_color="transparent")
        prog_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        self.progresso_bar = ctk.CTkProgressBar(prog_frame)
        self.progresso_bar.pack(side="left", fill="x", expand=True, padx=(5, 10))
        self.progresso_bar.set(0)
        self.progresso_contador = ctk.CTkLabel(prog_frame, text="0/0", font=("Arial", 12, "bold"))
        self.progresso_contador.pack(side="right", padx=5)

        # 4. Log
        ctk.CTkLabel(tab, text="Log de Execução:", anchor="w").grid(row=3, column=0, padx=10, pady=(10, 0), sticky="nw")
        self.log_textbox = ctk.CTkTextbox(tab)
        self.log_textbox.grid(row=4, column=0, padx=10, pady=(0, 10), sticky="nsew")

    def setup_cadastro_tab(self):
        tab = self.tabview.tab("Cadastro")
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1) 

        # Add
        add_frame = ctk.CTkFrame(tab)
        add_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        self.nova_cidade_entry = ctk.CTkEntry(add_frame, placeholder_text="Nome da Cidade")
        self.nova_cidade_entry.pack(side="left", fill="x", expand=True, padx=10, pady=10)
        ctk.CTkButton(add_frame, text="➕ Adicionar", width=100, command=self.add_cidade_ui).pack(side="right", padx=10, pady=10)

        # List
        self.lista_cidades_txt = ctk.CTkTextbox(tab)
        self.lista_cidades_txt.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")

        # Del
        del_frame = ctk.CTkFrame(tab)
        del_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        self.del_id_entry = ctk.CTkEntry(del_frame, placeholder_text="ID para Excluir")
        self.del_id_entry.pack(side="left", fill="x", expand=True, padx=10, pady=10)
        ctk.CTkButton(del_frame, text="🗑 Excluir", width=100, fg_color="red", hover_color="darkred", command=self.del_cidade_ui).pack(side="right", padx=10, pady=10)

    # ####################################################################
    # --- HANDLERS ---
    # ####################################################################

    def carregar_cidades_db(self):
        nomes = buscar_nomes_cidades()
        self.cidades_combobox.configure(values=["NENHUM FILTRO"] + nomes)
        self.cidades_combobox.set("NENHUM FILTRO")
        
        lista = listar_cidades()
        self.lista_cidades_txt.delete("0.0", "end")
        if not lista:
            self.lista_cidades_txt.insert("end", "Nenhuma cidade cadastrada.\n")
        for id_c, nome in lista:
            self.lista_cidades_txt.insert("end", f"[{id_c}] {nome}\n")

    def iniciar_automacao(self):
        try:
            cidade = self.cidades_combobox.get()
            delay_str = self.delay_combobox.get()
            
            # Valida e atualiza a variável global de delay
            delay = utils.validar_e_obter_delay(delay_str)
            utils.DELAY_ATUAL = delay # <--- ATUALIZA GLOBAL

            utils.INDICE_ATUAL_DO_CICLO = 0
            utils.PARAR_AUTOMACAO = False
            
            self.log_textbox.delete("1.0", "end")
            self._safe_update_gui(status="Rodando", total_ciclos=0, ciclo_atual=0)
            self._safe_configure_buttons("disabled", "disabled")
            
            self.log_textbox.insert("end", f"Iniciando com Delay: {delay}s...\n")

            if not self.monitor_thread_started:
                t = threading.Thread(target=utils.monitorar_tecla_escape, args=(self.log_textbox,), daemon=True)
                t.start()
                self.monitor_thread_started = True

            t_core = threading.Thread(target=automacao_core, 
                                      args=(self.log_textbox, cidade, delay, self._safe_configure_buttons, self._safe_update_gui))
            t_core.start()

        except Exception as e:
            self.log_textbox.insert("end", f"❌ Erro ao iniciar: {e}\n")
            self._safe_update_gui(status="Erro")
            self._safe_configure_buttons("normal", "disabled")

    def continuar_automacao(self):
        """Libera a thread de automação da pausa, com NOVO delay."""
        if not utils.PARAR_AUTOMACAO: return
        
        try:
            # Pega o valor ATUAL do combobox, caso o usuário tenha mudado
            delay_str = self.delay_combobox.get()
            novo_delay = utils.validar_e_obter_delay(delay_str)
            
            # Atualiza a global para que a thread perceba a mudança
            utils.DELAY_ATUAL = novo_delay # <--- ATUALIZA GLOBAL DINAMICAMENTE
            
            utils.PARAR_AUTOMACAO = False
            self._safe_update_gui(status="Rodando")
            self._safe_configure_buttons("disabled", "disabled")
            self.log_textbox.insert("end", f"▶ Retomando com Delay ajustado para: {novo_delay}s\n")
            
        except Exception as e:
            self.log_textbox.insert("end", f"❌ Erro ao continuar: {e}\n")

    def add_cidade_ui(self):
        msg, ok = adicionar_cidade(self.nova_cidade_entry.get())
        if ok: 
            self.nova_cidade_entry.delete(0, "end")
            self.carregar_cidades_db()
        CTkMessagebox(title="Cadastro", message=msg, icon="check" if ok else "cancel")

    def del_cidade_ui(self):
        msg, ok = excluir_cidade(self.del_id_entry.get())
        if ok:
            self.del_id_entry.delete(0, "end")
            self.carregar_cidades_db()
        CTkMessagebox(title="Exclusão", message=msg, icon="check" if ok else "cancel")

    def abrir_excel(self):
        ok, msg = utils.abrir_planilha_alvo()
        self.log_textbox.insert("end", msg + "\n")