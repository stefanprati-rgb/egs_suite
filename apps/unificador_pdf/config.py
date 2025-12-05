"""
Configurações e constantes do Unificador.
"""

# Clientes conhecidos para nomes amigáveis
CLIENTES_CONHECIDOS = {
    'igreja_do_evangelho_quadrangular': 'IGREJA_EVANGELHO_QUADRANGULAR',
}

# Configurações da janela
WINDOW_CONFIG = {
    'title': "Unificador de Faturas e Boletos",
    'geometry': "900x720",
    'min_width': 800,
    'min_height': 650,
}

# Paleta de cores
COLORS = {
    'primary': '#2563eb',
    'primary_hover': '#1d4ed8',
    'success': '#10b981',
    'success_hover': '#059669',
    'warning': '#f59e0b',
    'error': '#ef4444',
    'error_hover': '#dc2626',
    'bg': '#f8fafc',
    'card': '#ffffff',
    'card_border': '#e2e8f0',
    'input_bg': '#f1f5f9',
    'text': '#1e293b',
    'text_light': '#64748b',
    'disabled': '#9ca3af',
}

# Fontes
FONTS = {
    'title': ('Segoe UI', 18, 'bold'),
    'subtitle': ('Segoe UI', 10),
    'heading': ('Segoe UI', 11, 'bold'),
    'body': ('Segoe UI', 10),
    'body_bold': ('Segoe UI', 10, 'bold'),
    'small': ('Segoe UI', 9),
    'small_bold': ('Segoe UI', 9, 'bold'),
    'mono': ('Consolas', 9),
    'icon': ('Segoe UI', 32),
}

# Padrões regex para extração de UC
UC_PATTERNS = [
    r'Unidade\s+Consumidora\D*(\d+[/\-]\d+[\-\d]*)',
    r'Codigo\s+Instala\D*(\d+[/\-]\d+[\-\d]*)',
    r'Instala\D*(\d+[/\-]\d+[\-\d]*)',
    r'(\d{2,}/\d{4,}-\d+)',  # formato explícito: XX/XXXX-X
]

# Padrões regex para extração de valores
VALUE_PATTERNS = {
    'fatura_total': r'Total\s+a\s+pagar.*?R\$\s*([\d\.,]+)',
    'fatura_isolado': r'R\$\s*([\d\.,]{4,})',
    'boleto_documento': r'=\)\s*Valor\s+do\s+Documento\D*([\d\.,]+)',
    'boleto_generico': r'Valor\s+do\s+Documento\D*([\d\.,]+)',
}
