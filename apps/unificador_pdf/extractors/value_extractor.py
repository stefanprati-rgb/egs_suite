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
    
    # Tenta padrão principal
    match = re.search(VALUE_PATTERNS['fatura_total'], texto, re.IGNORECASE)
    if match:
        valor = _str_to_float(match.group(1))
        log.debug(f"[VALOR_FATURA] Encontrado (padrão principal): {valor}")
        return valor
    
    # Tenta padrão alternativo
    match = re.search(VALUE_PATTERNS['fatura_total_alt'], texto, re.IGNORECASE)
    if match:
        valor = _str_to_float(match.group(1))
        log.debug(f"[VALOR_FATURA] Encontrado (padrão alternativo): {valor}")
        return valor
        
    log.warning(f"[VALOR_FATURA] Nenhum valor encontrado em {nome_arquivo}")
    return None

def extrair_valor_boleto(texto: str, nome_arquivo: str = "") -> Optional[float]:
    """
    Extrai valor do boleto.
    Estratégia:
    1. Tenta campo explícito 'Valor do Documento'.
    2. Se falhar, tenta decodificar a Linha Digitável (Código de Barras).
    """
    log = get_logger()
    
    # 1. Tentativa via Rótulo (padrão principal)
    match_doc = re.search(VALUE_PATTERNS['boleto_documento'], texto, re.IGNORECASE)
    if match_doc:
        valor = _str_to_float(match_doc.group(1))
        if valor:
            log.debug(f"[VALOR_BOLETO] Via rótulo (principal): {valor}")
            return valor
    
    # 1b. Tentativa via Rótulo (padrão alternativo)
    match_doc_alt = re.search(VALUE_PATTERNS['boleto_documento_alt'], texto, re.IGNORECASE)
    if match_doc_alt:
        valor = _str_to_float(match_doc_alt.group(1))
        if valor:
            log.debug(f"[VALOR_BOLETO] Via rótulo (alternativo): {valor}")
            return valor

    # 2. Tentativa via Linha Digitável (Mais robusta para layouts ruins)
    numeros_apenas = re.sub(r'[^\d]', '', texto)
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

def comparar_valores(valor1: Optional[float], valor2: Optional[float]) -> bool:
    """Compara dois valores com match exato (sem tolerância)."""
    if valor1 is None or valor2 is None:
        return False
    return abs(valor1 - valor2) < 0.01  # Tolerância mínima para erros de arredondamento

def analisar_valores_pdf(texto: str, nome_arquivo: str = "", tipo: str = "fatura") -> dict:
    """
    Analisa os valores encontrados para diagnóstico visual na interface.
    """
    # Encontrar todos os valores R$ para mostrar possíveis candidatos
    valores_rs = re.findall(r'R\$\s*([\d\.,]+)', texto)
    
    analise = {
        'arquivo': nome_arquivo,
        'tipo': tipo,
        'valores_rs': valores_rs[:5], # Top 5 valores encontrados
        'tem_total_pagar': bool(re.search(VALUE_PATTERNS['fatura_total'], texto, re.IGNORECASE)),
        'tem_valor_documento': bool(re.search(VALUE_PATTERNS['boleto_documento'], texto, re.IGNORECASE)),
        'tem_linha_digitavel': bool(re.search(r'\d{40,}', re.sub(r'[^\d]', '', texto)))
    }
    
    return analise