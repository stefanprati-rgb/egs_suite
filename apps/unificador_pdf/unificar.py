"""
Unificador de PDFs - Versão Simples
Junta fatura + boleto com mesma UC em um único PDF de 2 páginas.
"""

import os
import re
from pathlib import Path
from PyPDF2 import PdfMerger, PdfReader
from logging_utils import get_logger

def extrair_uc_do_nome(nome_arquivo):
    """Extrai UC do nome do arquivo (primeiros dígitos antes do _)."""
    partes = nome_arquivo.split('_')
    if partes and partes[0]:
        # Pega apenas dígitos
        uc = re.sub(r'[^\d]', '', partes[0])
        if len(uc) >= 6:
            return uc[:7]  # Limita a 7 dígitos
    return None

def unificar_pdfs(pasta_faturas, pasta_boletos, pasta_saida):
    """
    Unifica faturas e boletos com mesma UC.
    
    Args:
        pasta_faturas: Caminho para pasta com faturas
        pasta_boletos: Caminho para pasta com boletos
        pasta_saida: Caminho para salvar PDFs unificados
    """
    log = get_logger()
    
    # Cria pasta de saída se não existir
    Path(pasta_saida).mkdir(parents=True, exist_ok=True)
    
    # Lista todos os PDFs
    faturas = list(Path(pasta_faturas).glob('*.pdf'))
    boletos = list(Path(pasta_boletos).glob('*.pdf'))
    
    log.info(f"Encontradas {len(faturas)} faturas e {len(boletos)} boletos")
    
    # Indexa boletos por UC
    boletos_por_uc = {}
    for boleto_path in boletos:
        uc = extrair_uc_do_nome(boleto_path.name)
        if uc:
            boletos_por_uc[uc] = boleto_path
            log.info(f"Boleto indexado: UC {uc} -> {boleto_path.name}")
        else:
            log.warning(f"Boleto sem UC válida: {boleto_path.name}")
    
    # Processa faturas
    unificados = 0
    nao_encontrados = []
    
    for fatura_path in faturas:
        uc = extrair_uc_do_nome(fatura_path.name)
        
        if not uc:
            log.warning(f"Fatura sem UC válida: {fatura_path.name}")
            nao_encontrados.append(fatura_path.name)
            continue
        
        # Procura boleto correspondente
        if uc in boletos_por_uc:
            boleto_path = boletos_por_uc[uc]
            
            # Nome do arquivo unificado
            nome_saida = f"{uc}_unificado.pdf"
            caminho_saida = Path(pasta_saida) / nome_saida
            
            try:
                # Junta os PDFs
                merger = PdfMerger()
                
                # Adiciona fatura (página 1)
                merger.append(str(fatura_path))
                
                # Adiciona boleto (página 2)
                merger.append(str(boleto_path))
                
                # Salva
                merger.write(str(caminho_saida))
                merger.close()
                
                log.print_success(f"✓ Unificado: UC {uc} -> {nome_saida}")
                log.info(f"  Fatura: {fatura_path.name}")
                log.info(f"  Boleto: {boleto_path.name}")
                
                unificados += 1
                
            except Exception as e:
                log.print_error(f"✗ Erro ao unificar UC {uc}: {e}")
                
        else:
            log.warning(f"✗ Boleto não encontrado para UC {uc} (fatura: {fatura_path.name})")
            nao_encontrados.append(fatura_path.name)
    
    # Resumo
    log.print_section("RESUMO")
    log.print_success(f"✓ {unificados} PDFs unificados com sucesso!")
    
    if nao_encontrados:
        log.warning(f"⚠ {len(nao_encontrados)} faturas sem boleto correspondente:")
        for nome in nao_encontrados[:10]:  # Mostra no máximo 10
            log.warning(f"  - {nome}")
        if len(nao_encontrados) > 10:
            log.warning(f"  ... e mais {len(nao_encontrados) - 10}")
    
    log.info(f"\nPDFs salvos em: {pasta_saida}")
    
    return unificados

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) != 4:
        print("Uso: python unificar.py <pasta_faturas> <pasta_boletos> <pasta_saida>")
        print("\nExemplo:")
        print('  python unificar.py "C:\\Faturas" "C:\\Boletos" "C:\\Unificados"')
        sys.exit(1)
    
    pasta_faturas = sys.argv[1]
    pasta_boletos = sys.argv[2]
    pasta_saida = sys.argv[3]
    
    # Valida pastas
    if not Path(pasta_faturas).exists():
        print(f"✗ Pasta de faturas não encontrada: {pasta_faturas}")
        sys.exit(1)
    
    if not Path(pasta_boletos).exists():
        print(f"✗ Pasta de boletos não encontrada: {pasta_boletos}")
        sys.exit(1)
    
    # Executa
    unificar_pdfs(pasta_faturas, pasta_boletos, pasta_saida)
