import sys
from pathlib import Path
import tkinter as tk

# Add EGS_Suite root to path to allow absolute imports
current_dir = Path(__file__).parent
suite_root = current_dir.parent.parent.parent
sys.path.append(str(suite_root))

from EGS_Suite.apps.buscador_boletos.modules.gui import App
from EGS_Suite.common.logging import setup_logger

if __name__ == "__main__":
    # Setup logger with UI callback support (if needed, or just basic)
    # The App class will likely hook into the logger later, for now just setup file logging
    setup_logger('buscador_boletos', queue_callback=None)
    
    root = tk.Tk()
    app = App(root)
    root.mainloop()
