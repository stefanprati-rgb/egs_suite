"""
Configurações e constantes do Unificador.
Define os padrões regex para extração de dados.
"""

# Configurações da janela (Visual)
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

# Clientes conhecidos
CLIENTES_CONHECIDOS = {
    'igreja_do_evangelho_quadrangular': 'IGREJA_EVANGELHO_QUADRANGULAR',
}

# ==========================================
# PADRÕES DE EXTRAÇÃO (REGEX)
# ==========================================

# Padrões para Unidade Consumidora (UC) ou Instalação
UC_PATTERNS = [
    # Rótulos explícitos
    r'(?:unidade\s+consumidora|c[óo]digo\s+de\s+instala[çc][ãa]o|c[óo]digo\s+instala[çc][ãa]o)\s*[:\.]?\s*([\d\./-]+)',
    r'(?:instala[çc][ãa]o)\s*[:\.]?\s*(\d+)',
    # Formatos específicos (Ex: 10/10232-7)
    r'(\d{2}/\d{4,}-\d)',
    r'(\d{10})', # Sequência de 10 dígitos (comum em UCs)
]

# Padrões para Valores Monetários
VALUE_PATTERNS = {
    # Fatura: Busca por "Total a pagar" ou similares
    'fatura_total': r'(?:Total\s+a\s+pagar|Valor\s+a\s+Pagar|Valor\s+Total|Total\s+da\s+Conta)[\s\S]{0,50}?R\$\s*([\d\.,]+)',
    
    # Boleto: Busca por campo "Valor do Documento"
    'boleto_documento': r'(?:=\s*)?Valor\s*do\s*Documento[\s\S]{0,20}?R\$\s*([\d\.,]+)',
    
    # Boleto: Linha Digitável (Captura sequências longas de números separadas por espaço ou ponto)
    # Ex: 52990.00108 90001.415851 ...
    'boleto_linha_digitavel': r'(\d{5}\.?\d{5}\s+\d{5}\.?\d{6}\s+\d{5}\.?\d{6}\s+\d\s+\d{14})',
}