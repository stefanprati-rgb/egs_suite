"""
Módulo de extratores para o Unificador de PDF.
"""

from .uc_extractor import (
    extrai_uc, 
    extrai_uc_do_texto, 
    extrai_referencia,
    analisar_texto_pdf,
    normalizar_uc
)

from .value_extractor import (
    extrair_valor_fatura, 
    extrair_valor_boleto,
    analisar_valores_pdf
)