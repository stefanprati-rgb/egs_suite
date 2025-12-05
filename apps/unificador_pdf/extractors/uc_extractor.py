"""
Extração de UC (Unidade Consumidora) de arquivos e textos PDF.
"""

import re
from typing import Optional
from ..config import UC_PATTERNS
from ..logging_utils import get_logger

def normalizar_uc(uc_raw: str) -> str:
    """Remove caracteres não numéricos para comparação."""
    return re.sub(r'[^\d]', '', uc_raw)

def extrai_uc(nome_arquivo: str) -> Optional[str]:
    """Extrai UC do nome do arquivo."""
    log = get_logger()
    nome_limpo = normalizar_uc(nome_arquivo)
    
    # Procura sequências de 6 a 12 dígitos no nome
    match = re.search(r'\d{6,12}', nome_limpo)
    if match:
        uc = match.group(0)
        log.debug(f"[UC_NOME] Encontrada no nome: {uc}")
        return uc
    return None

def extrai_uc_do_texto(texto: str, nome_arquivo: str = "") -> Optional[str]:
    """Extrai UC do conteúdo do PDF."""
    log = get_logger()
    
    if not texto:
        return None
        
    for padrao in UC_PATTERNS:
        matches = re.finditer(padrao, texto, re.IGNORECASE)
        for match in matches:
            uc_encontrada = match.group(1)
            # Limpa para verificar se é válida
            uc_limpa = normalizar_uc(uc_encontrada)
            
            # Validação básica: UCs geralmente têm entre 6 e 12 dígitos
            if 6 <= len(uc_limpa) <= 12:
                log.debug(f"[UC_TEXTO] Encontrada no texto: {uc_limpa} (Raw: {uc_encontrada})")
                return uc_limpa
                
    log.warning(f"[UC_TEXTO] Nenhuma UC encontrada em {nome_arquivo}")
    return None

def extrai_referencia(texto: str, nome_arquivo: str = "") -> Optional[str]:
    """Extrai mês/ano de referência (ex: nov/2025)."""
    # Regex para meses abreviados ou completos e ano
    match = re.search(r'(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[a-z]*\W+(\d{4})', texto, re.IGNORECASE)
    if match:
        return f"{match.group(1).lower()}/{match.group(2)}"
    return None