"""
Extração de valores monetários de faturas e boletos.
Inclui logging detalhado para debugging.
"""

import re
from typing import Optional
from ..config import VALUE_PATTERNS
from ..logging_utils import get_logger


def _parse_valor_brasileiro(valor_str: str) -> Optional[float]:
    """
    Converte string de valor brasileiro (ex: '1.997,44') para float.
    
    Args:
        valor_str: String com valor no formato brasileiro
        
    Returns:
        Float com o valor ou None se conversão falhar
    """
    log = get_logger()
    
    try:
        # Remove pontos de milhar e substitui vírgula por ponto
        valor_limpo = valor_str.replace('.', '').replace(',', '.')
        valor = float(valor_limpo)
        log.debug(f"[VALOR] Conversão: '{valor_str}' → '{valor_limpo}' → {valor}")
        return valor
    except ValueError as e:
        log.debug(f"[VALOR] Erro ao converter '{valor_str}': {e}")
        return None


def extrair_valor_fatura(texto: str, nome_arquivo: str = "") -> Optional[float]:
    """
    Extrai valor da fatura do texto do PDF.
    
    Procura por:
    1. Bloco "Total a pagar ... R$ X.XXX,XX"
    2. Valor isolado "R$ X.XXX,XX"
    
    Args:
        texto: Texto extraído do PDF da fatura
        nome_arquivo: Nome do arquivo (para logging)
        
    Returns:
        Float com o valor ou None se não encontrado
    """
    log = get_logger()
    
    log.debug(f"[VALOR_FATURA] Extraindo valor de: {nome_arquivo}")
    log.debug(f"[VALOR_FATURA] Tamanho do texto: {len(texto)} chars")
    
    # Log de todos os valores R$ encontrados
    todos_valores = re.findall(r'R\$\s*([\d\.,]+)', texto)
    log.debug(f"[VALOR_FATURA] Todos os valores R$ encontrados: {todos_valores[:10]}")
    
    # Primeiro tenta o bloco "Total a pagar"
    log.debug(f"[VALOR_FATURA] Tentando padrão 'Total a pagar'...")
    match = re.search(
        VALUE_PATTERNS['fatura_total'],
        texto,
        re.IGNORECASE | re.DOTALL
    )
    
    if match:
        valor_str = match.group(1)
        log.debug(f"[VALOR_FATURA] Match 'Total a pagar': '{valor_str}'")
        
        valor = _parse_valor_brasileiro(valor_str)
        if valor is not None:
            log.debug(f"[VALOR_FATURA] ✓ Valor extraído (Total a pagar): R$ {valor:.2f}")
            return valor
        else:
            log.debug(f"[VALOR_FATURA] Falha ao converter valor do 'Total a pagar'")
    else:
        log.debug(f"[VALOR_FATURA] Padrão 'Total a pagar' não encontrado")
        
        # Mostrar contexto onde poderia estar
        if 'total' in texto.lower():
            idx = texto.lower().find('total')
            contexto = texto[max(0, idx-20):idx+100]
            log.debug(f"[VALOR_FATURA] Contexto de 'total': ...{contexto}...")
    
    # Depois tenta o valor isolado perto do topo
    log.debug(f"[VALOR_FATURA] Tentando padrão 'valor isolado R$'...")
    match = re.search(VALUE_PATTERNS['fatura_isolado'], texto)
    
    if match:
        valor_str = match.group(1)
        log.debug(f"[VALOR_FATURA] Match 'valor isolado': '{valor_str}'")
        
        valor = _parse_valor_brasileiro(valor_str)
        if valor is not None:
            log.debug(f"[VALOR_FATURA] ✓ Valor extraído (isolado): R$ {valor:.2f}")
            return valor
        else:
            log.debug(f"[VALOR_FATURA] Falha ao converter valor isolado")
    else:
        log.debug(f"[VALOR_FATURA] Padrão 'valor isolado' não encontrado")
    
    log.warning(f"[VALOR_FATURA] ✗ Nenhum valor encontrado em {nome_arquivo}")
    return None


def extrair_valor_boleto(texto: str, nome_arquivo: str = "") -> Optional[float]:
    """
    Extrai valor do documento do boleto.
    
    Procura por:
    1. "(=) Valor do Documento: X.XXX,XX"
    2. "Valor do Documento: X.XXX,XX" (fallback)
    
    Args:
        texto: Texto extraído do PDF do boleto
        nome_arquivo: Nome do arquivo (para logging)
        
    Returns:
        Float com o valor ou None se não encontrado
    """
    log = get_logger()
    
    log.debug(f"[VALOR_BOLETO] Extraindo valor de: {nome_arquivo}")
    log.debug(f"[VALOR_BOLETO] Tamanho do texto: {len(texto)} chars")
    
    # Log de todos os valores encontrados
    todos_valores = re.findall(r'[\d\.,]{4,}', texto)
    log.debug(f"[VALOR_BOLETO] Todos os números encontrados: {todos_valores[:10]}")
    
    # Primeiro tenta o formato com (=)
    log.debug(f"[VALOR_BOLETO] Tentando padrão '(=) Valor do Documento'...")
    match = re.search(
        VALUE_PATTERNS['boleto_documento'],
        texto,
        re.IGNORECASE | re.DOTALL
    )
    
    if match:
        valor_str = match.group(1)
        log.debug(f"[VALOR_BOLETO] Match '(=) Valor do Documento': '{valor_str}'")
        
        valor = _parse_valor_brasileiro(valor_str)
        if valor is not None:
            log.debug(f"[VALOR_BOLETO] ✓ Valor extraído: R$ {valor:.2f}")
            return valor
        else:
            log.debug(f"[VALOR_BOLETO] Falha ao converter valor")
    else:
        log.debug(f"[VALOR_BOLETO] Padrão '(=) Valor do Documento' não encontrado")
    
    # Fallback mais genérico
    log.debug(f"[VALOR_BOLETO] Tentando padrão genérico 'Valor do Documento'...")
    match = re.search(
        VALUE_PATTERNS['boleto_generico'],
        texto,
        re.IGNORECASE
    )
    
    if match:
        valor_str = match.group(1)
        log.debug(f"[VALOR_BOLETO] Match 'Valor do Documento': '{valor_str}'")
        
        valor = _parse_valor_brasileiro(valor_str)
        if valor is not None:
            log.debug(f"[VALOR_BOLETO] ✓ Valor extraído (genérico): R$ {valor:.2f}")
            return valor
        else:
            log.debug(f"[VALOR_BOLETO] Falha ao converter valor genérico")
    else:
        log.debug(f"[VALOR_BOLETO] Padrão genérico não encontrado")
        
        # Mostrar contexto onde poderia estar
        if 'valor' in texto.lower():
            idx = texto.lower().find('valor')
            contexto = texto[max(0, idx-20):idx+100]
            log.debug(f"[VALOR_BOLETO] Contexto de 'valor': ...{contexto}...")
    
    # Tentar padrões alternativos de boleto
    padroes_alternativos = [
        r'Valor\s+Cobrado\D*([\d\.,]+)',
        r'Valor\s+Total\D*([\d\.,]+)',
        r'VALOR\D*([\d\.,]+)',
    ]
    
    for padrao in padroes_alternativos:
        log.debug(f"[VALOR_BOLETO] Tentando padrão alternativo: {padrao}")
        match = re.search(padrao, texto, re.IGNORECASE)
        if match:
            valor_str = match.group(1)
            valor = _parse_valor_brasileiro(valor_str)
            if valor is not None and valor > 0:
                log.debug(f"[VALOR_BOLETO] ✓ Valor extraído (alternativo): R$ {valor:.2f}")
                return valor
    
    log.warning(f"[VALOR_BOLETO] ✗ Nenhum valor encontrado em {nome_arquivo}")
    return None


def analisar_valores_pdf(texto: str, nome_arquivo: str = "", tipo: str = "fatura") -> dict:
    """
    Analisa todos os valores monetários encontrados no PDF.
    
    Args:
        texto: Texto extraído do PDF
        nome_arquivo: Nome do arquivo
        tipo: 'fatura' ou 'boleto'
        
    Returns:
        Dicionário com análise de valores
    """
    log = get_logger()
    
    # Encontrar todos os valores R$
    valores_rs = re.findall(r'R\$\s*([\d\.,]+)', texto)
    
    # Encontrar números que parecem valores
    numeros = re.findall(r'\b(\d{1,3}(?:\.\d{3})*,\d{2})\b', texto)
    
    analise = {
        'arquivo': nome_arquivo,
        'tipo': tipo,
        'valores_rs': valores_rs[:10],
        'numeros_formatados': numeros[:10],
        'tem_total_pagar': bool(re.search(r'total\s+a\s+pagar', texto, re.IGNORECASE)),
        'tem_valor_documento': bool(re.search(r'valor\s+do\s+documento', texto, re.IGNORECASE)),
    }
    
    log.info(f"[ANÁLISE_VALORES] {nome_arquivo} ({tipo}):")
    log.info(f"  - Valores R$: {analise['valores_rs']}")
    log.info(f"  - Números formatados: {analise['numeros_formatados']}")
    log.info(f"  - Tem 'Total a pagar': {analise['tem_total_pagar']}")
    log.info(f"  - Tem 'Valor do Documento': {analise['tem_valor_documento']}")
    
    return analise
