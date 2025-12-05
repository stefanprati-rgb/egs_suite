import sys
from pathlib import Path
import tkinter as tk
import logging
from tkinter import messagebox

# Add EGS_Suite root to path to allow absolute imports
current_dir = Path(__file__).parent
suite_root = current_dir.parent.parent.parent
sys.path.append(str(suite_root))

from EGS_Suite.common.logging import setup_logger
from EGS_Suite.apps.enviador_emails.modules.gui import EmailSenderApp

if __name__ == "__main__":
    try:
        # Configura o logging
        setup_logger('enviador_emails', queue_callback=None)
        
        root = tk.Tk()
        app = EmailSenderApp(root)
        logging.info("Interface gráfica inicializada com sucesso.")
        root.mainloop()
    except Exception as e:
        # Tenta logar o erro crítico
        try:
            logging.critical(f"Erro fatal na inicialização da aplicação: {e}", exc_info=True)
        except:
            print(f"Erro fatal (logging falhou): {e}")
            
        messagebox.showerror("ERRO FATAL DE INICIALIZAÇÃO", f"Ocorreu um erro crítico ao iniciar: {e}\n\nConsulte o log para mais detalhes.")
    finally:
        logging.info("Sessão finalizada.")
