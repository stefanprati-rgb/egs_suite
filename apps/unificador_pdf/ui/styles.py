"""
Estilos e temas para a interface gráfica.
"""

from tkinter import ttk
from ..config import COLORS, FONTS


def configurar_estilos() -> dict:
    """
    Configura os estilos ttk e retorna o dicionário de cores.
    
    Returns:
        Dicionário com as cores configuradas
    """
    style = ttk.Style()
    
    style.configure(
        'Title.TLabel',
        font=FONTS['title'],
        foreground=COLORS['text']
    )
    
    style.configure(
        'Subtitle.TLabel',
        font=FONTS['subtitle'],
        foreground=COLORS['text_light']
    )
    
    style.configure(
        'Card.TFrame',
        background=COLORS['card'],
        relief='flat'
    )
    
    style.configure(
        'Counter.TLabel',
        font=FONTS['small'],
        foreground=COLORS['text_light']
    )
    
    style.configure(
        'Primary.TButton',
        font=FONTS['body_bold']
    )
    
    return COLORS
