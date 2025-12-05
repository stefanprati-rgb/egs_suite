#!/usr/bin/env python3
"""
Unificador de Faturas e Boletos - Ponto de Entrada

Combina PDFs de faturas e boletos por UC (Unidade Consumidora) e valor.

Uso:
    python -m unificador
    
Ou:
    python main.py
"""

import sys
from pathlib import Path
import tkinter as tk

# Add EGS_Suite root to path to allow absolute imports
current_dir = Path(__file__).parent
suite_root = current_dir.parent.parent.parent
sys.path.append(str(suite_root))

from EGS_Suite.apps.unificador_pdf.ui import AppUnificador

def main():
    """Função principal da aplicação."""
    root = tk.Tk()
    app = AppUnificador(root)
    root.mainloop()


if __name__ == "__main__":
    main()