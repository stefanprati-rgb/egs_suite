"""
Módulo de manipulação de PDFs.
"""

from .reader import validar_pdf_cabecalho, abrir_pdf_seguro, extrair_texto_pdf
from .writer import unir_pdfs, criar_nome_arquivo

__all__ = [
    'validar_pdf_cabecalho',
    'abrir_pdf_seguro',
    'extrair_texto_pdf',
    'unir_pdfs',
    'criar_nome_arquivo',
]
