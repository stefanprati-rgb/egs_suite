"""
Extração de UC (Unidade Consumidora) de arquivos e textos PDF.
Inclui logging detalhado para debugging.
"""

import re
from typing import Optional
from ..config import UC_PATTERNS
from ..logging_utils import get_logger


def extrai_uc(nome_arquivo: str) -> Optional[str]:
    """
    Extrai UC do nome do arquivo (6+ dígitos contínuos).
    
    Args:
        nome_arquivo: Nome do arquivo PDF
        
    Returns:
        String com os dígitos da UC ou None se não encontrado
    """
    log = get_logger()
    log.debug(f"[UC_NOME] Tentando extrair UC do nome: '{nome_arquivo}'")
    
    nome_normalizado = re.sub(r'[_\-]', '', nome_arquivo)
    log.debug(f"[UC_NOME] Nome normalizado: '{nome_normalizado}'")
    
    match = re.search(r'\d{6,}', nome_normalizado)
    
    if match:
        uc = match.group(0)
        log.debug(f"[UC_NOME] ✓ UC encontrada no nome: {uc}")
        return uc
    else:
        log.debug(f"[UC_NOME] ✗ Nenhum padrão de 6+ dígitos encontrado no nome")
        return None


def extrai_uc_do_texto(texto: str, nome_arquivo: str = "") -> Optional[str]:
    """
    Extrai UC de formatos como '10/5202-7' mesmo com OCR ruim ou caracteres especiais.
    
    Args:
        texto: Texto extraído do PDF
        nome_arquivo: Nome do arquivo (para logging)
        
    Returns:
        String com os dígitos da UC ou None se não encontrado
    """
    log = get_logger()
    
    # Log dos primeiros caracteres do texto
    texto_preview = texto[:500].replace('\n', ' ').replace('\r', '')
    log.debug(f"[UC_TEXTO] Arquivo: {nome_arquivo}")
    log.debug(f"[UC_TEXTO] Preview do texto (500 chars): {texto_preview}")
    log.debug(f"[UC_TEXTO] Tamanho total do texto: {len(texto)} caracteres")
    
    if not texto or len(texto.strip()) < 10:
        log.warning(f"[UC_TEXTO] Texto muito curto ou vazio para {nome_arquivo}")
        return None
    
    # Normaliza dois-pontos (fullwidth colon → regular colon)
    texto_normalizado = re.sub(r'[：:]', ':', texto)
    
    log.debug(f"[UC_TEXTO] Testando {len(UC_PATTERNS)} padrões de UC...")
    
    for i, padrao in enumerate(UC_PATTERNS, 1):
        log.debug(f"[UC_TEXTO] Padrão {i}/{len(UC_PATTERNS)}: {padrao}")
        
        match = re.search(padrao, texto_normalizado, re.IGNORECASE | re.DOTALL)
        
        if match:
            uc_bruta = match.group(1)
            uc_limpa = re.sub(r'[^\d]', '', uc_bruta)
            
            log.debug(f"[UC_TEXTO]   → Match encontrado: '{uc_bruta}' → limpa: '{uc_limpa}'")
            
            if len(uc_limpa) >= 6:
                log.debug(f"[UC_TEXTO] ✓ UC válida extraída do texto: {uc_limpa}")
                return uc_limpa
            else:
                log.debug(f"[UC_TEXTO]   → UC muito curta ({len(uc_limpa)} dígitos), continuando...")
        else:
            log.debug(f"[UC_TEXTO]   → Sem match")
    
    # Log de palavras-chave encontradas para ajudar no debug
    keywords = ['unidade', 'consumidora', 'instalação', 'instalacao', 'codigo', 'uc:', 'uc ']
    found_keywords = [kw for kw in keywords if kw.lower() in texto.lower()]
    log.debug(f"[UC_TEXTO] Palavras-chave encontradas: {found_keywords}")
    
    # Tentar encontrar qualquer sequência de dígitos com / ou -
    possiveis_ucs = re.findall(r'\d{2,}/\d{3,}-?\d*', texto)
    if possiveis_ucs:
        log.debug(f"[UC_TEXTO] Possíveis UCs não capturadas pelos padrões: {possiveis_ucs[:5]}")
    
    log.warning(f"[UC_TEXTO] ✗ Nenhuma UC encontrada no texto de {nome_arquivo}")
    return None


def analisar_texto_pdf(texto: str, nome_arquivo: str = "") -> dict:
    """
    Analisa o texto do PDF e retorna informações de diagnóstico.
    
    Args:
        texto: Texto extraído do PDF
        nome_arquivo: Nome do arquivo
        
    Returns:
        Dicionário com informações de análise
    """
    log = get_logger()
    
    analise = {
        'arquivo': nome_arquivo,
        'tamanho_texto': len(texto),
        'linhas': texto.count('\n'),
        'palavras': len(texto.split()),
        'tem_unidade_consumidora': 'unidade' in texto.lower() and 'consumidora' in texto.lower(),
        'tem_instalacao': 'instala' in texto.lower(),
        'tem_valor_rs': 'r$' in texto.lower(),
        'tem_total_pagar': 'total' in texto.lower() and 'pagar' in texto.lower(),
        'possiveis_ucs': re.findall(r'\d{2,}/\d{3,}-?\d*', texto)[:5],
        'possiveis_valores': re.findall(r'R\$\s*[\d\.,]+', texto)[:5],
    }
    
    log.info(f"[ANÁLISE] {nome_arquivo}:")
    log.info(f"  - Tamanho: {analise['tamanho_texto']} chars, {analise['linhas']} linhas")
    log.info(f"  - Tem 'Unidade Consumidora': {analise['tem_unidade_consumidora']}")
    log.info(f"  - Tem 'Instalação': {analise['tem_instalacao']}")
    log.info(f"  - Tem valores R$: {analise['tem_valor_rs']}")
    log.info(f"  - Possíveis UCs: {analise['possiveis_ucs']}")
    log.info(f"  - Possíveis valores: {analise['possiveis_valores']}")
    
    return analise
