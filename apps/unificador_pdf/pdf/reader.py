"""
Leitura e validação de arquivos PDF.
"""

from pathlib import Path
from pypdf import PdfReader, errors as pdf_errors
import pdfplumber



def validar_pdf_cabecalho(caminho_arquivo: Path) -> bool:
    """
    Valida se o arquivo é um PDF válido verificando o cabeçalho.
    
    Args:
        caminho_arquivo: Caminho para o arquivo PDF
        
    Returns:
        True se o arquivo começa com '%PDF-', False caso contrário
    """
    try:
        with open(caminho_arquivo, 'rb') as f:
            return f.read(5) == b'%PDF-'
    except Exception:
        return False


def abrir_pdf_seguro(caminho_arquivo: Path) -> PdfReader:
    """
    Abre um PDF de forma segura, tratando criptografia e erros.
    
    Args:
        caminho_arquivo: Caminho para o arquivo PDF
        
    Returns:
        PdfReader configurado para leitura
        
    Raises:
        RuntimeError: Se o PDF não puder ser lido
    """
    try:
        reader = PdfReader(str(caminho_arquivo), strict=False)
        
        # Tenta descriptografar se necessário
        if getattr(reader, "is_encrypted", False):
            try:
                reader.decrypt("")
            except Exception:
                pass
            
            if reader.is_encrypted:
                raise pdf_errors.FileNotDecryptedError(
                    "PDF criptografado e sem senha."
                )
        
        # Valida que tem páginas
        if len(reader.pages) == 0:
            raise ValueError("PDF sem páginas válidas")
        
        return reader
        
    except Exception as e:
        raise RuntimeError(
            f"Falha ao ler PDF: {type(e).__name__} - {str(e)}"
        )


def extrair_texto_pdf(caminho_arquivo: Path) -> str:
    """
    Extrai todo o texto de um PDF usando pdfplumber.
    
    Args:
        caminho_arquivo: Caminho para o arquivo PDF
        
    Returns:
        String com todo o texto concatenado das páginas
    """
    textos = []
    try:
        with pdfplumber.open(caminho_arquivo) as pdf:
            for page in pdf.pages:
                texto = page.extract_text() or ""
                if texto:
                    textos.append(texto)
        return "\n".join(textos)
    except Exception as e:
        raise RuntimeError(f"Erro ao extrair texto com pdfplumber: {e}")

