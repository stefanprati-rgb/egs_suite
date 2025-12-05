"""
Componentes reutilizáveis da interface gráfica.
"""

import tkinter as tk
from tkinter import ttk, scrolledtext
from typing import Callable

from ..config import COLORS, FONTS


def criar_card(parent: tk.Widget, bg: str = COLORS['card']) -> tk.Frame:
    """
    Cria um card com bordas arredondadas e sombra.
    
    Args:
        parent: Widget pai
        bg: Cor de fundo
        
    Returns:
        Frame configurado como card
    """
    card = tk.Frame(
        parent,
        bg=bg,
        relief='flat',
        borderwidth=1,
        highlightbackground=COLORS['card_border'],
        highlightthickness=1
    )
    return card


def criar_seletor_pasta(
    parent: tk.Widget,
    label: str,
    var_caminho: tk.StringVar,
    var_contador: tk.StringVar,
    comando: Callable,
    emoji: str,
    cores: dict
) -> tk.Frame:
    """
    Cria um seletor de pasta com contador de arquivos.
    
    Args:
        parent: Widget pai
        label: Texto do label
        var_caminho: StringVar para o caminho
        var_contador: StringVar para o contador
        comando: Função ao clicar no botão
        emoji: Emoji para o label
        cores: Dicionário de cores
        
    Returns:
        Frame com o seletor completo
    """
    frame = tk.Frame(parent, bg=cores['card'])
    
    # Header
    header = tk.Frame(frame, bg=cores['card'])
    header.pack(fill=tk.X, pady=(0, 5))
    
    tk.Label(
        header,
        text=f"{emoji} {label}",
        font=FONTS['body_bold'],
        bg=cores['card'],
        fg=cores['text']
    ).pack(side=tk.LEFT)
    
    ttk.Label(
        header,
        textvariable=var_contador,
        style='Counter.TLabel'
    ).pack(side=tk.RIGHT)
    
    # Input
    input_frame = tk.Frame(frame, bg=cores['card'])
    input_frame.pack(fill=tk.X)
    
    entry = tk.Entry(
        input_frame,
        textvariable=var_caminho,
        state='readonly',
        font=FONTS['small'],
        relief='solid',
        borderwidth=1,
        bg=cores['input_bg'],
        fg=cores['text']
    )
    entry.pack(side=tk.LEFT, expand=True, fill=tk.X, ipady=8)
    
    btn = tk.Button(
        input_frame,
        text="Escolher",
        command=comando,
        font=FONTS['small_bold'],
        bg=cores['primary'],
        fg='white',
        relief='flat',
        cursor='hand2',
        padx=20,
        pady=8
    )
    btn.pack(side=tk.LEFT, padx=(10, 0))
    
    # Hover effects
    btn.bind('<Enter>', lambda e: btn.config(bg=cores['primary_hover']))
    btn.bind('<Leave>', lambda e: btn.config(bg=cores['primary']))
    
    return frame


def criar_botao_acao(
    parent: tk.Widget,
    texto: str,
    comando: Callable,
    cor: str,
    cor_hover: str,
    **kwargs
) -> tk.Button:
    """
    Cria um botão de ação com hover effect.
    
    Args:
        parent: Widget pai
        texto: Texto do botão
        comando: Função ao clicar
        cor: Cor de fundo
        cor_hover: Cor no hover
        **kwargs: Argumentos adicionais para o Button
        
    Returns:
        Button configurado
    """
    btn = tk.Button(
        parent,
        text=texto,
        command=comando,
        font=FONTS['heading'],
        bg=cor,
        fg='white',
        relief='flat',
        cursor='hand2',
        pady=12,
        **kwargs
    )
    
    btn.bind('<Enter>', lambda e: btn.config(bg=cor_hover) if btn['state'] == 'normal' else None)
    btn.bind('<Leave>', lambda e: btn.config(bg=cor) if btn['state'] == 'normal' else None)
    
    return btn


def criar_area_relatorio(
    parent: tk.Widget,
    cores: dict
) -> scrolledtext.ScrolledText:
    """
    Cria área de texto para relatório.
    
    Args:
        parent: Widget pai
        cores: Dicionário de cores
        
    Returns:
        ScrolledText configurado
    """
    return scrolledtext.ScrolledText(
        parent,
        height=12,
        state='disabled',
        wrap=tk.WORD,
        font=FONTS['mono'],
        bg=cores['bg'],
        fg=cores['text'],
        relief='flat',
        borderwidth=0
    )
