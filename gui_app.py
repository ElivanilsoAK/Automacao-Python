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
        self.title("BotEliva — Automação de Relatórios")
        self.geometry("900x680")
        self.minsize(800, 600)

        # Icon
        import os
        icon_path = "app_icon.ico"
        if getattr(sys, 'frozen', False):
            icon_path = os.path.join(sys._MEIPASS, "app_icon.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)

        self.config = load_config()
        self.worker = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

        self.tab_monitor = self.tabview.add("📊 Monitor")
        self.tab_config = self.tabview.add("⚙️ Configuração")

        self.setup_monitor_tab()
        self.setup_config_tab()

        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_monitor_tab(self):
        # Header
        header_frame = ctk.CTkFrame(self.tab_monitor, fg_color="transparent")
        header_frame.pack(fill="x", padx=5, pady=(10, 5))

        self.header_label = ctk.CTkLabel(
            header_frame, text="BotEliva — Painel de Controle",
            font=("Segoe UI", 22, "bold"), text_color="#198754"
        )
        self.header_label.pack(side="left")

        # Status Bar com cor dinâmica
        self.status_frame = ctk.CTkFrame(self.tab_monitor, corner_radius=10, fg_color="#1a1a2e")
        self.status_frame.pack(fill="x", padx=5, pady=5)
        self.status_label = ctk.CTkLabel(
            self.status_frame, text="● Status: Parado",
            font=("Segoe UI", 14, "bold"), text_color="#dc3545"
        )
        self.status_label.pack(pady=8)

        # Stats Cards
        self.stats_frame = ctk.CTkFrame(self.tab_monitor, fg_color="transparent")
        self.stats_frame.pack(pady=5, fill="x", padx=5)
        self.stats_frame.grid_columnconfigure((0, 1, 2), weight=1)

        # Card: Próxima Execução
        _card_next = ctk.CTkFrame(self.stats_frame, corner_radius=10, fg_color="#1a1a2e")
        _card_next.grid(row=0, column=0, padx=5, pady=3, sticky="ew")
        ctk.CTkLabel(_card_next, text="Próxima Exec.", font=("Segoe UI", 11), text_color="#888").pack(pady=(8, 0))
        self.stat_next_run_val = ctk.CTkLabel(_card_next, text="--:--", font=("Segoe UI", 18, "bold"), text_color="#0dcaf0")
        self.stat_next_run_val.pack(pady=(0, 8))

        # Card: Sucessos
        _card_suc = ctk.CTkFrame(self.stats_frame, corner_radius=10, fg_color="#1a1a2e")
        _card_suc.grid(row=0, column=1, padx=5, pady=3, sticky="ew")
        ctk.CTkLabel(_card_suc, text="Sucessos (Sessão)", font=("Segoe UI", 11), text_color="#888").pack(pady=(8, 0))
        self.stat_success_val = ctk.CTkLabel(_card_suc, text="0", font=("Segoe UI", 18, "bold"), text_color="#198754")
        self.stat_success_val.pack(pady=(0, 8))

        # Card: Falhas
        _card_fail = ctk.CTkFrame(self.stats_frame, corner_radius=10, fg_color="#1a1a2e")
        _card_fail.grid(row=0, column=2, padx=5, pady=3, sticky="ew")
        ctk.CTkLabel(_card_fail, text="Falhas/Pulos", font=("Segoe UI", 11), text_color="#888").pack(pady=(8, 0))
        self.stat_fails_val = ctk.CTkLabel(_card_fail, text="0", font=("Segoe UI", 18, "bold"), text_color="#dc3545")
        self.stat_fails_val.pack(pady=(0, 8))

        # Log Box
        self.log_box = ctk.CTkTextbox(
            self.tab_monitor, font=("Consolas", 12),
            text_color="#d1d1d1", fg_color="#0d0d0d",
            border_color="#2a2a2a", border_width=1
        )
        self.log_box.pack(pady=(8, 5), padx=5, expand=True, fill="both")
        self.log_box.insert("0.0", "--- Bem-vindo ao BotEliva ---\nConfigure suas credenciais na aba ⚙️ Configuração e clique em Iniciar.\n\n")

        # Botões de Controle
        self.btn_frame = ctk.CTkFrame(self.tab_monitor, fg_color="transparent")
        self.btn_frame.pack(pady=(5, 10))

        self.start_btn = ctk.CTkButton(
            self.btn_frame, text="▶  Iniciar Serviço",
            command=self.start_service, fg_color="#198754", hover_color="#145c38",
            font=("Segoe UI", 13, "bold"), width=160, height=38
        )
        self.start_btn.pack(side="left", padx=8)

        self.stop_btn = ctk.CTkButton(
            self.btn_frame, text="■  Parar Serviço",
            command=self.stop_service, fg_color="#c0392b", hover_color="#96281b",
            font=("Segoe UI", 13, "bold"), width=160, height=38, state="disabled"
        )
        self.stop_btn.pack(side="left", padx=8)

        # NOVO: Botão "Executar Agora" para testar sem esperar agendamento
        self.run_now_btn = ctk.CTkButton(
            self.btn_frame, text="⚡  Executar Agora",
            command=self.run_now, fg_color="#e67e22", hover_color="#ca6f1e",
            font=("Segoe UI", 13, "bold"), width=160, height=38, state="disabled"
        )
        self.run_now_btn.pack(side="left", padx=8)

    def setup_config_tab(self):
        # Usar um scrollable frame para a aba de configuração
        scroll = ctk.CTkScrollableFrame(self.tab_config, label_text="Configurações do Sistema", label_font=("Segoe UI", 14, "bold"))
        scroll.pack(fill="both", expand=True, padx=10, pady=10)
        scroll.grid_columnconfigure(1, weight=1)

        self.entries = {}
        row = 0

        # Pasta de Download
        ctk.CTkLabel(scroll, text="Pasta de Download", anchor="e").grid(row=row, column=0, padx=10, pady=6, sticky="e")
        path_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        path_frame.grid(row=row, column=1, padx=10, pady=6, sticky="ew")
        path_frame.grid_columnconfigure(0, weight=1)

        self.path_entry = ctk.CTkEntry(path_frame)
        self.path_entry.insert(0, str(self.config.get("download_path", "")))
        self.path_entry.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        ctk.CTkButton(path_frame, text="Selecionar", width=90, command=self.browse_folder).grid(row=0, column=1)
        self.entries["download_path"] = self.path_entry
        row += 1

        # Tipo de Frequência
        ctk.CTkLabel(scroll, text="Tipo de Frequência", anchor="e").grid(row=row, column=0, padx=10, pady=6, sticky="e")
        self.type_var = ctk.StringVar(value=str(self.config.get("schedule_type", "Diário")))
        ctk.CTkOptionMenu(
            scroll, values=["Diário", "Semanal", "Mensal"],
            variable=self.type_var, command=self.update_dynamic_fields
        ).grid(row=row, column=1, padx=10, pady=6, sticky="w")
        row += 1

        # Campos principais
        fields = [
            ("URL InControl", "incontrol_url"),
            ("Usuário InControl", "incontrol_user"),
            ("Senha InControl", "incontrol_password", True),
            ("Servidor SMTP", "smtp_server"),
            ("Porta SMTP", "smtp_port"),
            ("Email Remetente", "smtp_user"),
            ("Senha Email", "smtp_password", True),
            ("Destinatários (separar por vírgula)", "email_recipients"),
            ("Horários (ex: 08:00, 15:30)", "schedule_times"),
        ]

        for label_text, key, *args in fields:
            is_pass = args[0] if args else False
            ctk.CTkLabel(scroll, text=label_text, anchor="e").grid(row=row, column=0, padx=10, pady=6, sticky="e")
            entry = ctk.CTkEntry(scroll, show="*" if is_pass else "")
            entry.insert(0, str(self.config.get(key, "")))
            entry.grid(row=row, column=1, padx=10, pady=6, sticky="ew")
            self.entries[key] = entry
            row += 1

        # Campos dinâmicos de dias
        self.dynamic_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        self.dynamic_frame.grid(row=row, column=0, columnspan=2, sticky="ew", pady=5)
        self.monthly_entry = None
        self.weekly_vars = {}
        row += 1

        self.update_dynamic_fields(self.type_var.get())

        # Separador visual
        ctk.CTkLabel(scroll, text="", height=5).grid(row=row, column=0, columnspan=2)
        row += 1

        # Checkbox Headless
        self.headless_var = ctk.BooleanVar(value=self.config.get("headless", False))
        ctk.CTkCheckBox(
            scroll, text="Modo Silencioso (Sem Janela do Navegador)",
            variable=self.headless_var, font=("Segoe UI", 12)
        ).grid(row=row, column=0, columnspan=2, pady=8)
        row += 1

        # Botão Salvar
        ctk.CTkButton(
            scroll, text="💾  Salvar Configurações",
            command=self.save_settings, fg_color="#198754", hover_color="#145c38",
            font=("Segoe UI", 13, "bold"), height=40
        ).grid(row=row, column=0, columnspan=2, pady=15, sticky="ew", padx=10)

    def update_dynamic_fields(self, selected_type):
        for widget in self.dynamic_frame.winfo_children():
            widget.destroy()

        self.monthly_entry = None
        self.weekly_vars = {}

        if selected_type == "Semanal":
            ctk.CTkLabel(self.dynamic_frame, text="Dias da Semana", anchor="e").grid(row=0, column=0, padx=10, pady=5, sticky="e")
            checkbox_frame = ctk.CTkFrame(self.dynamic_frame, fg_color="transparent")
            checkbox_frame.grid(row=0, column=1, padx=10, pady=5, sticky="w")

            dias = [("Seg", "seg"), ("Ter", "ter"), ("Qua", "qua"),
                    ("Qui", "qui"), ("Sex", "sex"), ("Sáb", "sab"), ("Dom", "dom")]
            saved_days = str(self.config.get("schedule_days", "")).lower()

            for i, (label, key) in enumerate(dias):
                var = ctk.BooleanVar(value=(key in saved_days))
                ctk.CTkCheckBox(checkbox_frame, text=label, variable=var, width=60).grid(row=0, column=i, padx=5)
                self.weekly_vars[key] = var

        elif selected_type == "Mensal":
            ctk.CTkLabel(self.dynamic_frame, text="Dias do Mês (1-31)", anchor="e").grid(row=0, column=0, padx=10, pady=5, sticky="e")
            self.monthly_entry = ctk.CTkEntry(self.dynamic_frame, placeholder_text="ex: 1, 15, 30")
            saved_days = str(self.config.get("schedule_days", ""))
            import re
            if re.match(r'^[\d,\s]+$', saved_days):
                self.monthly_entry.insert(0, saved_days)
            self.monthly_entry.grid(row=0, column=1, padx=10, pady=5, sticky="w")

    def browse_folder(self):
        folder = ctk.filedialog.askdirectory()
        if folder:
            self.path_entry.delete(0, "end")
            self.path_entry.insert(0, folder)

    def log(self, message):
        def _update():
            from datetime import datetime
            time_str = datetime.now().strftime("%H:%M:%S")

            # Coloração de texto por tipo de mensagem
            tag = "normal"
            if any(w in message for w in ["❌", "ERRO", "Erro", "Falha", "falha"]):
                tag = "error"
            elif any(w in message for w in ["✅", "sucesso", "Sucesso", "enviado", "concluído"]):
                tag = "success"
            elif any(w in message for w in ["⚠️", "AVISO", "Aviso"]):
                tag = "warn"

            self.log_box.tag_config("error", foreground="#ff6b6b")
            self.log_box.tag_config("success", foreground="#6bcb77")
            self.log_box.tag_config("warn", foreground="#ffd166")
            self.log_box.tag_config("normal", foreground="#d1d1d1")

            self.log_box.insert("end", f"[{time_str}] {message}\n", tag)

            # Limitar a 600 linhas
            lines = int(self.log_box.index("end-1c").split(".")[0])
            if lines > 600:
                self.log_box.delete("1.0", f"{lines - 600}.0")

            self.log_box.see("end")
        self.after(0, _update)

    def update_status(self, message):
        def _update():
            # Cor dinâmica baseada no status
            if any(w in message for w in ["Parado", "⛔"]):
                color = "#dc3545"
                dot = "●"
            elif any(w in message for w in ["Aguardando", "⏳"]):
                color = "#ffc107"
                dot = "◑"
            elif any(w in message for w in ["Executando", "Extraindo", "Analisando", "Enviando", "⚡"]):
                color = "#0dcaf0"
                dot = "◉"
            else:
                color = "#198754"
                dot = "●"
            self.status_label.configure(text=f"{dot} Status: {message}", text_color=color)
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
        self.log("--- ▶ Iniciando Serviço ---")
        self.worker = AutomationWorker(
            self.config, self.log, self.update_status, self.clear_log, self.update_stats
        )
        self.worker.start()
        self.update_status("Aguardando...")

        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.run_now_btn.configure(state="normal")
        self.tabview.set("📊 Monitor")

    def stop_service(self):
        if self.worker:
            self.worker.stop()
        self.update_status("Parado")
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.run_now_btn.configure(state="disabled")

    def run_now(self):
        """NOVO: Executa um ciclo imediatamente sem esperar o agendamento."""
        if self.worker and not self.worker.stop_flag:
            self.log("⚡ Execução manual solicitada...")
            self.run_now_btn.configure(state="disabled", text="Executando...")
            def _run_and_re_enable():
                self.worker.run_task()
                self.after(0, lambda: self.run_now_btn.configure(state="normal", text="⚡  Executar Agora"))
            t = threading.Thread(target=_run_and_re_enable, daemon=True)
            t.start()
        else:
            self.log("⚠️ Inicie o serviço antes de usar 'Executar Agora'.")

    def save_settings(self):
        new_config = {}
        for key, entry in self.entries.items():
            val = entry.get()
            if key == "smtp_port":
                try:
                    val = int(val)
                except:
                    val = 587
            new_config[key] = val

        new_config["schedule_type"] = self.type_var.get()
        new_config["headless"] = self.headless_var.get()

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
            self.log("✅ Configurações salvas com sucesso!")
            if self.worker:
                self.log("⚠️ Reinicie o serviço para aplicar as novas configurações.")
        else:
            self.log("❌ Erro ao salvar configurações.")

    def on_close(self):
        self.stop_service()
        self.destroy()
        sys.exit()


if __name__ == "__main__":
    app = App()
    app.mainloop()
