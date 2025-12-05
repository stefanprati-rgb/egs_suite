"""
Extração de valores monetários. Foco em robustez usando Linha Digitável.
"""

import re
from typing import Optional
from ..config import VALUE_PATTERNS
from ..logging_utils import get_logger

def _str_to_float(valor_str: str) -> Optional[float]:
    """Converte '1.234,56' para 1234.56"""
    try:
        limpo = valor_str.replace('.', '').replace(',', '.')
        return float(limpo)
    except:
        return None

def extrair_valor_fatura(texto: str, nome_arquivo: str = "") -> Optional[float]:
    """Tenta extrair o 'Total a Pagar' da fatura."""
    log = get_logger()
    
    match = re.search(VALUE_PATTERNS['fatura_total'], texto, re.IGNORECASE)
    if match:
        valor = _str_to_float(match.group(1))
        log.debug(f"[VALOR_FATURA] Encontrado: {valor}")
        return valor
        
    return None

def extrair_valor_boleto(texto: str, nome_arquivo: str = "") -> Optional[float]:
    """
    Extrai valor do boleto.
    Estratégia:
    1. Tenta campo explícito 'Valor do Documento'.
    2. Se falhar, tenta decodificar a Linha Digitável (Código de Barras).
    """
    log = get_logger()
    
    # 1. Tentativa via Rótulo
    match_doc = re.search(VALUE_PATTERNS['boleto_documento'], texto, re.IGNORECASE)
    if match_doc:
        valor = _str_to_float(match_doc.group(1))
        if valor:
            log.debug(f"[VALOR_BOLETO] Via rótulo: {valor}")
            return valor

    # 2. Tentativa via Linha Digitável (Mais robusta para layouts ruins)
    # Remove quebras de linha para unir números que podem estar quebrados visualmente
    texto_continuo = texto.replace('\n', ' ').replace('\r', '')
    
    # Procura por sequências longas de números que parecem linha digitável
    # Removemos pontos e espaços para facilitar a busca genérica
    numeros_apenas = re.sub(r'[^\d]', '', texto)
    
    # Linha digitável padrão tem 47 ou 48 dígitos (com ou sem os dígitos verificadores expandidos)
    # O valor está sempre nos últimos 10 dígitos da sequência de 44 caracteres do código de barras real,
    # ou no final da linha digitável.
    
    # Padrão simples: Procurar sequência de 44 a 48 dígitos
    match_barras = re.search(r'(\d{44,48})', numeros_apenas)
    
    if match_barras:
        sequencia = match_barras.group(1)
        # O valor são os últimos 10 dígitos (sem pontuação). Os 2 últimos são centavos.
        valor_str = sequencia[-10:]
        try:
            valor = float(valor_str) / 100.0
            if valor > 0:
                log.debug(f"[VALOR_BOLETO] Via linha digitável: {valor} (Seq: {valor_str})")
                return valor
        except:
            pass
            
    log.warning(f"[VALOR_BOLETO] Falha ao extrair valor em {nome_arquivo}")
    return None