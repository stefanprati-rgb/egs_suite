"""
Escrita e união de arquivos PDF.
"""

import io
from pathlib import Path
from pypdf import PdfWriter

from ..config import CLIENTES_CONHECIDOS
from .reader import abrir_pdf_seguro


def criar_nome_arquivo(
    uc: str,
    nome_fatura: str,
    nome_boleto: str,
    mes: str = "Mes",
    ano: str = "2025"
) -> str:
    """
    Cria nome padronizado para o arquivo unificado.
    
    Args:
        uc: Número da UC
        nome_fatura: Nome original do arquivo da fatura
        nome_boleto: Nome original do arquivo do boleto
        mes: Mês de referência
        ano: Ano de referência
        
    Returns:
        Nome formatado: UC_{uc}_{cliente}_{mes}_{ano}.pdf
    """
    cliente = "CLIENTE"
    
    for nome_original, nome_formatado in CLIENTES_CONHECIDOS.items():
        if nome_original in nome_fatura.lower():
            cliente = nome_formatado
            break
    
    periodo = mes[:3]
    return f"UC_{uc}_{cliente}_{periodo}_{ano}.pdf"


def unir_pdfs(
    caminhos: list[str | Path],
    ordem_boleto_primeiro: bool = False
) -> bytes:
    """
    Une múltiplos PDFs em um único arquivo.
    
    Args:
        caminhos: Lista de caminhos para os PDFs (fatura, boleto)
        ordem_boleto_primeiro: Se True, coloca boleto antes da fatura
        
    Returns:
        Bytes do PDF unificado
        
    Raises:
        RuntimeError: Se algum PDF não puder ser lido
    """
    if ordem_boleto_primeiro:
        caminhos = list(reversed(caminhos))
    
    writer = PdfWriter()
    
    for caminho in caminhos:
        reader = abrir_pdf_seguro(Path(caminho))
        for page in reader.pages:
            writer.add_page(page)
    
    # Escreve para buffer em memória
    output = io.BytesIO()
    writer.write(output)
    output.seek(0)
    
    return output.read()
