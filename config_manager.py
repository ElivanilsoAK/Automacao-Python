import json
import os

CONFIG_FILE = "config.json"

DEFAULT_CONFIG = {
    "incontrol_url": "https://<IP_DO_INCONTROL>:<PORTA>/#/home/eventos-usuario",
    "incontrol_user": "admin",
    "incontrol_password": "",
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587,
    "smtp_user": "",
    "smtp_password": "",
    "email_recipients": "",
    "coluna_empresa": "Departamento",
    "schedule_type": "Diário",
    "schedule_times": "08:00",
    "schedule_days": "",
    "download_path": os.getcwd() # Default: pasta atual
}

def load_config():
    if not os.path.exists(CONFIG_FILE):
        save_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG
    
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return DEFAULT_CONFIG

def save_config(config_data):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config_data, f, indent=4)
        return True
    except Exception as e:
        print(f"Erro ao salvar config: {e}")
        return False
