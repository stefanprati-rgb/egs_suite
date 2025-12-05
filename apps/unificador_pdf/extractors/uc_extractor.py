"""
Extração de UC (Unidade Consumidora) de arquivos e textos PDF.
"""

import re
from typing import Optional
from ..config import UC_PATTERNS
from ..logging_utils import get_logger

def normalizar_uc(uc_raw: str) -> str:
    """Remove caracteres não numéricos para comparação."""
    if not uc_raw:
        return ""
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
            uc_limpa = normalizar_uc(uc_encontrada)
            
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

def analisar_texto_pdf(texto: str, nome_arquivo: str = "") -> dict:
    """
    Analisa o texto do PDF e retorna informações de diagnóstico.
    Utilizado pela interface para mostrar status do arquivo.
    """
    log = get_logger()
    
    # Busca rápida por padrões para o diagnóstico
    possiveis_ucs = []
    for p in UC_PATTERNS:
        found = re.findall(p, texto, re.IGNORECASE)
        possiveis_ucs.extend(found[:2]) # Pega os 2 primeiros de cada padrão

    analise = {
        'arquivo': nome_arquivo,
        'tamanho_texto': len(texto),
        'linhas': texto.count('\n'),
        'tem_unidade_consumidora': 'unidade' in texto.lower() and 'consumidora' in texto.lower(),
        'tem_instalacao': 'instala' in texto.lower(),
        'tem_valor_rs': 'r$' in texto.lower(),
        'possiveis_ucs': [normalizar_uc(u) for u in possiveis_ucs if u],
    }
    
    log.info(f"[ANÁLISE] {nome_arquivo}: {analise['tamanho_texto']} chars, UCs={analise['possiveis_ucs']}")
    return analise