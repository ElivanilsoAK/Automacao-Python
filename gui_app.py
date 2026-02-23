import customtkinter as ctk
import sys
import threading
from config_manager import load_config, save_config
from automation_worker import AutomationWorker

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("BotEliva - Automação")
        self.geometry("850x650")
        
        # Icon
        import os
        icon_path = "app_icon.ico"
        if getattr(sys, 'frozen', False):
            icon_path = os.path.join(sys._MEIPASS, "app_icon.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)
        
        # Configuração Inicial
        self.config = load_config()
        self.worker = None

        # Layout Main
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        self.tab_monitor = self.tabview.add("Monitor")
        self.tab_config = self.tabview.add("Configuração")
        
        self.setup_monitor_tab()
        self.setup_config_tab()

        # Protocolo de Fechamento
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_monitor_tab(self):
        # Header
        self.header_label = ctk.CTkLabel(self.tab_monitor, text="Painel de Controle", font=("Segoe UI", 24, "bold"), text_color="#198754")
        self.header_label.pack(pady=(10, 5))

        # Status Bar
        self.status_label = ctk.CTkLabel(self.tab_monitor, text="Status: Parado", font=("Segoe UI", 16, "bold"))
        self.status_label.pack(pady=5)

        # Contadores (Stats)
        self.stats_frame = ctk.CTkFrame(self.tab_monitor, fg_color="transparent")
        self.stats_frame.pack(pady=5, fill="x", padx=10)
        self.stats_frame.grid_columnconfigure((0,1,2), weight=1)
        
        # Coluna 1
        self.stat_next_run_lbl = ctk.CTkLabel(self.stats_frame, text="Próxima Exec.", font=("Segoe UI", 12))
        self.stat_next_run_lbl.grid(row=0, column=0, pady=(0, 2))
        self.stat_next_run_val = ctk.CTkLabel(self.stats_frame, text="--:--", font=("Segoe UI", 16, "bold"), text_color="#0dcaf0")
        self.stat_next_run_val.grid(row=1, column=0)
        
        # Coluna 2
        self.stat_success_lbl = ctk.CTkLabel(self.stats_frame, text="Sucessos (Hoje)", font=("Segoe UI", 12))
        self.stat_success_lbl.grid(row=0, column=1, pady=(0, 2))
        self.stat_success_val = ctk.CTkLabel(self.stats_frame, text="0", font=("Segoe UI", 16, "bold"), text_color="#198754")
        self.stat_success_val.grid(row=1, column=1)

        # Coluna 3
        self.stat_fails_lbl = ctk.CTkLabel(self.stats_frame, text="Falhas/Pulos", font=("Segoe UI", 12))
        self.stat_fails_lbl.grid(row=0, column=2, pady=(0, 2))
        self.stat_fails_val = ctk.CTkLabel(self.stats_frame, text="0", font=("Segoe UI", 16, "bold"), text_color="#dc3545")
        self.stat_fails_val.grid(row=1, column=2)

        # Log Text Box
        self.log_box = ctk.CTkTextbox(self.tab_monitor, width=780, height=400, font=("Consolas", 12), text_color="#d1d1d1", fg_color="#1e1e1e", border_color="#333333", border_width=1)
        self.log_box.pack(pady=10, padx=10, expand=True, fill="both")
        self.log_box.insert("0.0", "--- Bem-vindo ao BotEliva ---\nConfigure suas credenciais na aba Configuração e clique em Iniciar.\n\n")

        # Botões
        self.btn_frame = ctk.CTkFrame(self.tab_monitor, fg_color="transparent")
        self.btn_frame.pack(pady=10)

        self.start_btn = ctk.CTkButton(self.btn_frame, text="Iniciar Serviço", command=self.start_service, fg_color="green")
        self.start_btn.pack(side="left", padx=10)

        self.stop_btn = ctk.CTkButton(self.btn_frame, text="Parar Serviço", command=self.stop_service, fg_color="red", state="disabled")
        self.stop_btn.pack(side="left", padx=10)

    def setup_config_tab(self):
        self.tab_config.grid_columnconfigure(1, weight=1)
        
        fields = [
            ("URL InControl", "incontrol_url"),
            ("Usuário InControl", "incontrol_user"),
            ("Senha InControl", "incontrol_password", True),
            ("Servidor SMTP", "smtp_server"),
            ("Porta SMTP", "smtp_port"),
            ("Email Remetente", "smtp_user"),
            ("Senha Email", "smtp_password", True),
            ("Destinatários (separar por ,)", "email_recipients"),
            ("Intervalo (minutos)", "interval_minutes")
        ]

        self.entries = {}
        row = 0
        
        # Campo Especial para Pasta de Download
        ctk.CTkLabel(self.tab_config, text="Pasta de Download").grid(row=row, column=0, padx=10, pady=5, sticky="e")
        path_frame = ctk.CTkFrame(self.tab_config, fg_color="transparent")
        path_frame.grid(row=row, column=1, padx=10, pady=5, sticky="w")
        
        self.path_entry = ctk.CTkEntry(path_frame, width=300)
        self.path_entry.insert(0, str(self.config.get("download_path", "")))
        self.path_entry.pack(side="left")
        
        self.browse_btn = ctk.CTkButton(path_frame, text="Selecionar", width=80, command=self.browse_folder)
        self.browse_btn.pack(side="left", padx=5)
        self.entries["download_path"] = self.path_entry # Link para salvar
        
        row += 1

        ctk.CTkLabel(self.tab_config, text="Tipo de Frequência").grid(row=row, column=0, padx=10, pady=5, sticky="e")
        self.type_var = ctk.StringVar(value=str(self.config.get("schedule_type", "Diário")))
        type_menu = ctk.CTkOptionMenu(self.tab_config, values=["Diário", "Semanal", "Mensal"], variable=self.type_var, width=400, command=self.update_dynamic_fields)
        type_menu.grid(row=row, column=1, padx=10, pady=5, sticky="w")
        row += 1

        fields = [
            ("URL InControl", "incontrol_url"),
            ("Usuário InControl", "incontrol_user"),
            ("Senha InControl", "incontrol_password", True),
            ("Servidor SMTP", "smtp_server"),
            ("Porta SMTP", "smtp_port"),
            ("Email Remetente", "smtp_user"),
            ("Senha Email", "smtp_password", True),
            ("Destinatários (separar por ,)", "email_recipients"),
            ("Horários (ex: 08:00, 15:30)", "schedule_times")
        ]

        for label_text, key, *args in fields:
            is_pass = args[0] if args else False
            
            ctk.CTkLabel(self.tab_config, text=label_text).grid(row=row, column=0, padx=10, pady=5, sticky="e")
            entry = ctk.CTkEntry(self.tab_config, width=400, show="*" if is_pass else "")
            entry.insert(0, str(self.config.get(key, "")))
            entry.grid(row=row, column=1, padx=10, pady=5, sticky="w")
            self.entries[key] = entry
            row += 1

        # --- CAMPOS DINÂMICOS DE DIAS ---
        self.dynamic_frame = ctk.CTkFrame(self.tab_config, fg_color="transparent")
        self.dynamic_frame.grid(row=row, column=0, columnspan=2, sticky="ew", pady=5)
        
        # Variáveis locais para armazenar estado dinâmico
        self.monthly_entry = None
        self.weekly_vars = {}
        
        row += 1
        
        # Iniciar campos dinâmicos
        self.update_dynamic_fields(self.type_var.get())
        
        # ------------------------------

        save_btn = ctk.CTkButton(self.tab_config, text="Salvar Configurações", command=self.save_settings)
        save_btn.grid(row=row, column=0, columnspan=2, pady=20)
        
        # Headless Checkbox
        self.headless_var = ctk.BooleanVar(value=self.config.get("headless", False))
        self.headless_check = ctk.CTkCheckBox(self.tab_config, text="Modo Silencioso (Sem Janela)", variable=self.headless_var)
        self.headless_check.grid(row=row+1, column=0, columnspan=2, pady=10)
        
    def update_dynamic_fields(self, selected_type):
        # Limpar o frame
        for widget in self.dynamic_frame.winfo_children():
            widget.destroy()
        
        self.monthly_entry = None
        self.weekly_vars = {}
        
        if selected_type == "Semanal":
            lbl = ctk.CTkLabel(self.dynamic_frame, text="Dias da Semana")
            lbl.grid(row=0, column=0, padx=10, pady=5, sticky="e")
            
            checkbox_frame = ctk.CTkFrame(self.dynamic_frame, fg_color="transparent")
            checkbox_frame.grid(row=0, column=1, padx=10, pady=5, sticky="w")
            
            dias = [("Seg", "seg"), ("Ter", "ter"), ("Qua", "qua"), 
                    ("Qui", "qui"), ("Sex", "sex"), ("Sáb", "sab"), ("Dom", "dom")]
            
            saved_days = str(self.config.get("schedule_days", "")).lower()
            
            for i, (label, key) in enumerate(dias):
                var = ctk.BooleanVar(value=(key in saved_days))
                cb = ctk.CTkCheckBox(checkbox_frame, text=label, variable=var, width=60)
                cb.grid(row=0, column=i, padx=5, pady=5)
                self.weekly_vars[key] = var
                
        elif selected_type == "Mensal":
            lbl = ctk.CTkLabel(self.dynamic_frame, text="Dias do Mês (1-31, separe por vírgula)")
            lbl.grid(row=0, column=0, padx=10, pady=5, sticky="e")
            
            self.monthly_entry = ctk.CTkEntry(self.dynamic_frame, width=400)
            
            # Se a última config era mensal, pega os dias, senão deixa vazio/limpo
            saved_days = str(self.config.get("schedule_days", ""))
            # Verifica se os dias salvos são números (ex: "1,15") para não preencher "seg,qua" num form mensal
            import re
            if re.match(r'^[\d,\s]+$', saved_days):
                self.monthly_entry.insert(0, saved_days)
            else:
                self.monthly_entry.insert(0, "")
                
            self.monthly_entry.grid(row=0, column=1, padx=10, pady=5, sticky="w")
        
        # Se for "Diário", o frame fica vazio (label e entry invisíveis)

    def browse_folder(self):
        folder = ctk.filedialog.askdirectory()
        if folder:
            self.path_entry.delete(0, "end")
            self.path_entry.insert(0, folder)

    def log(self, message):
        def _update():
            # Formata log
            from datetime import datetime
            time_str = datetime.now().strftime("%H:%M:%S")
            self.log_box.insert("end", f"[{time_str}] {message}\n")
            
            # Limita a 500 linhas para economizar memória
            current_lines = float(self.log_box.index("end-1c"))
            if current_lines > 500:
                self.log_box.delete("1.0", f"{current_lines - 500}.0")
            
            self.log_box.see("end")
        self.after(0, _update)

    def update_status(self, message):
        def _update():
            self.status_label.configure(text=f"Status: {message}")
        self.after(0, _update)

    def clear_log(self):
        def _clear():
            self.log_box.delete("1.0", "end")
        self.after(0, _clear)
        
    def update_stats(self, successes, fails, next_run):
        def _update():
            self.stat_success_val.configure(text=str(successes))
            self.stat_fails_val.configure(text=str(fails))
            self.stat_next_run_val.configure(text=str(next_run))
        self.after(0, _update)

    def start_service(self):
        self.log("--- Inicializando ---")
        self.worker = AutomationWorker(self.config, self.log, self.update_status, self.clear_log, self.update_stats)
        self.worker.start()
        
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.tabview.set("Monitor")

    def stop_service(self):
        if self.worker:
            self.worker.stop()
        
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")

    def save_settings(self):
        new_config = {}
        for key, entry in self.entries.items():
            val = entry.get()
            if key == "smtp_port":
                try:
                    val = int(val)
                except:
                    val = 0
            new_config[key] = val
        
        new_config["schedule_type"] = self.type_var.get()
        new_config["headless"] = self.headless_var.get()
        
        # Recuperando e salvando o schedule_days dinâmico
        sched_type = self.type_var.get()
        if sched_type == "Semanal":
            selected_days = [k for k, v in self.weekly_vars.items() if v.get()]
            new_config["schedule_days"] = ",".join(selected_days)
        elif sched_type == "Mensal" and self.monthly_entry:
            new_config["schedule_days"] = self.monthly_entry.get().strip()
        else:
            new_config["schedule_days"] = ""
        
        if save_config(new_config):
            self.config = new_config
            self.log("Configurações salvas com sucesso!")
            if self.worker:
                self.log("⚠️ Reinicie o serviço para aplicar novas configurações.")
        else:
            self.log("Erro ao salvar configurações.")

    def on_close(self):
        self.stop_service()
        self.destroy()
        sys.exit()

if __name__ == "__main__":
    app = App()
    app.mainloop()
