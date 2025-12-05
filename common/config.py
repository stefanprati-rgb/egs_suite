from pathlib import Path
import os

# Base paths
BASE_DIR = Path(__file__).parent.parent
LOG_DIR = BASE_DIR / "logs"
DATA_DIR = BASE_DIR / "data"
APPS_DIR = BASE_DIR / "apps"

# Ensure directories exist
LOG_DIR.mkdir(exist_ok=True)
DATA_DIR.mkdir(exist_ok=True)

# Common constants
COMPANY_NAME = "EGS Energia"
EMAIL_SENDER = "atendimento@egsenergia.com.br"

# App specific paths can be defined relative to APPS_DIR
BUSCADOR_DIR = APPS_DIR / "buscador_boletos"
ENVIADOR_DIR = APPS_DIR / "enviador_emails"
UNIFICADOR_DIR = APPS_DIR / "unificador_pdf"
