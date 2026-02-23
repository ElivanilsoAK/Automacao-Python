import time
import os
import signal
import sys
from automation_worker import AutomationWorker
from config_manager import load_config
from dotenv import load_dotenv

# Carregar variáveis de ambiente se existirem, mas priorizar config.json
load_dotenv()

def main():
    print("--- BotEliva CLI ---")
    
    # 1. Carregar Configuração
    config = load_config()
    
    # Merge com variáveis de ambiente se config.json estiver vazio em certos campos
    if not config.get("incontrol_user"):
        config["incontrol_user"] = os.getenv("INCONTROL_USER", "")
    if not config.get("incontrol_password"):
        config["incontrol_password"] = os.getenv("INCONTROL_PASSWORD", "")
    if not config.get("smtp_password"):
        config["smtp_password"] = os.getenv("SMTP_PASSWORD", "")

    # Callback simples para print
    def log_printer(msg):
        print(f"[Worker] {msg}")

    def status_printer(msg):
        print(f"[Status] {msg}")

    # 2. Instanciar Worker
    worker = AutomationWorker(config, logger_callback=log_printer, status_callback=status_printer)
    
    # Handler para CTRL+C
    def signal_handler(sig, frame):
        print("\nCTRL+C detectado. Parando...")
        worker.stop()
        sys.exit(0)
    
    signal.signal(signal.SIGINT, signal_handler)

    # 3. Rodar (thread principal vai ficar fazendo join ou loop)
    # Como o worker herda de Thread, start() inicia em background.
    # Mas no modo CLI queremos que o programa principal espere.
    
    worker.start()
    
    try:
        while worker.is_alive():
            time.sleep(1)
    except KeyboardInterrupt:
        worker.stop()

if __name__ == "__main__":
    main()


