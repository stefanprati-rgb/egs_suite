import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import sys
import os
from pathlib import Path

# Setup path to ensure imports work if run as script
current_dir = Path(__file__).parent
if str(current_dir) not in sys.path:
    sys.path.append(str(current_dir))

try:
    from common.logging import setup_logger, get_logger
    from common.config import COMPANY_NAME, LOG_DIR
except ImportError:
    # Fallback if run from outside without module context
    sys.path.append(str(current_dir.parent))
    from EGS_Suite.common.logging import setup_logger, get_logger
    from EGS_Suite.common.config import COMPANY_NAME, LOG_DIR

class EGSLauncher:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{COMPANY_NAME} - Suite Integrada")
        self.root.geometry("600x450")
        self.root.configure(bg="#f8f9fa")
        
        self.logger = setup_logger('launcher', None)
        
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("TButton", font=("Segoe UI", 11), padding=10)
        self.style.configure("TLabel", background="#f8f9fa", font=("Segoe UI", 10))
        self.style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"), foreground="#003366")
        
        self._create_ui()

    def _create_ui(self):
        # Header
        header = ttk.Frame(self.root, padding="20")
        header.pack(fill=tk.X)
        
        lbl_title = ttk.Label(header, text="EGS Suite", style="Title.TLabel")
        lbl_title.pack(anchor=tk.CENTER)
        
        lbl_subtitle = ttk.Label(header, text="Selecione o módulo para iniciar", font=("Segoe UI", 12))
        lbl_subtitle.pack(anchor=tk.CENTER, pady=(5, 0))
        
        # Buttons Frame
        btn_frame = ttk.Frame(self.root, padding="30")
        btn_frame.pack(fill=tk.BOTH, expand=True)
        
        # Grid layout for buttons
        btn_frame.columnconfigure(0, weight=1)
        
        # Button 1: Buscador
        btn_buscador = ttk.Button(
            btn_frame, 
            text="🔍 Buscar Boletos (Outlook)", 
            command=self.launch_buscador
        )
        btn_buscador.pack(fill=tk.X, pady=10)
        
        # Button 2: Unificador
        btn_unificador = ttk.Button(
            btn_frame, 
            text="📄 Unificar PDFs (Faturas + Boletos)", 
            command=self.launch_unificador
        )
        btn_unificador.pack(fill=tk.X, pady=10)
        
        # Button 3: Enviador
        btn_enviador = ttk.Button(
            btn_frame, 
            text="📧 Enviar E-mails", 
            command=self.launch_enviador
        )
        btn_enviador.pack(fill=tk.X, pady=10)
        
        # Footer
        footer = ttk.Frame(self.root, padding="10")
        footer.pack(fill=tk.X, side=tk.BOTTOM)
        
        lbl_log = ttk.Label(footer, text=f"Logs: {LOG_DIR}", font=("Segoe UI", 8), foreground="#666666")
        lbl_log.pack(side=tk.LEFT)
        
        btn_open_logs = ttk.Button(footer, text="Abrir Pasta", command=self.open_logs, width=12)
        btn_open_logs.pack(side=tk.RIGHT)

    def _launch_module(self, module_name, friendly_name):
        try:
            self.logger.info(f"Launching {friendly_name}...")
            # Run as a separate process using python -m
            # Assuming we are in EGS_Suite root or parent is in path
            # We use the full package path: EGS_Suite.apps.MODULE.main
            
            cmd = [sys.executable, "-m", f"EGS_Suite.apps.{module_name}.main"]
            
            # Use Popen to not block the launcher
            subprocess.Popen(cmd, cwd=current_dir.parent)
            
        except Exception as e:
            msg = f"Erro ao iniciar {friendly_name}: {e}"
            self.logger.error(msg)
            messagebox.showerror("Erro", msg)

    def launch_buscador(self):
        self._launch_module("buscador_boletos", "Buscador de Boletos")

    def launch_unificador(self):
        self._launch_module("unificador_pdf", "Unificador de PDFs")

    def launch_enviador(self):
        self._launch_module("enviador_emails", "Enviador de E-mails")

    def open_logs(self):
        try:
            os.startfile(LOG_DIR)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir a pasta: {e}")

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = EGSLauncher(root)
        root.mainloop()
    except Exception as e:
        # Fallback error reporting
        import traceback
        traceback.print_exc()
        try:
            messagebox.showerror("Fatal Error", f"Launcher crashed: {e}")
        except:
            print(e)
