import os
import logging
from logging.handlers import RotatingFileHandler

# --- Configurações ---
# MODIFICADO: Usa o caminho exato fornecido pelo usuário para salvar os logs.
USER_PROJECT_PATH = r"C:\Users\Stefan_Pratti\OneDrive - GRUPO GERA\Área de Trabalho\Projeto EGS"

# Verifica se o caminho do usuário existe, senão usa o local do script como fallback.
if os.path.isdir(USER_PROJECT_PATH):
    LOG_BASE_PATH = USER_PROJECT_PATH
else:
    LOG_BASE_PATH = os.getcwd() # getcwd() pega o diretório atual do script

LOG_FOLDER = os.path.join(LOG_BASE_PATH, "logs_enviador")
LOG_FILE_PATH = os.path.join(LOG_FOLDER, "log_enviador_emails.log")
EMAIL_REMETENTE = "atendimento@egsenergia.com.br"

# Logging configuration is now handled by EGS_Suite.common.logging
# keeping constants if needed by other modules, but setup_logging is deprecated.

def setup_logging():
    # Deprecated: Now handled by common logger
    pass

