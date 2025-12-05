"""
Script de diagnóstico para identificar problemas no pareamento de PDFs.
Execute este script na pasta onde estão seus PDFs para ver detalhes do processamento.
"""

import sys
import os
import re
from pathlib import Path

# Adiciona o caminho do módulo ao sys.path
current_dir = Path(__file__).parent
if str(current_dir) not in sys.path:
    sys.path.insert(0, str(current_dir))

# Importa diretamente os módulos necessários
try:
    import pdfplumber
except ImportError:
    print("❌ Erro: pdfplumber não está instalado!")
    print("Execute: pip install pdfplumber")
    sys.exit(1)

# Importa configurações e funções localmente
import config
from logging_utils import get_logger

def ler_pdf(caminho):
    """Lê o texto de um PDF usando pdfplumber (com fallback para PyPDF2)."""
    # Verifica se o arquivo existe
    if not Path(caminho).exists():
        print(f"❌ Arquivo não encontrado: {caminho}")
        return None
    
    print(f"📂 Lendo: {Path(caminho).name}")
    print(f"📏 Tamanho: {Path(caminho).stat().st_size / 1024:.2f} KB")
    
    # Tentativa 1: pdfplumber (melhor para PDFs estruturados)
    try:
        with pdfplumber.open(caminho) as pdf:
            print(f"📄 Páginas: {len(pdf.pages)}")
            texto = ""
            for i, pagina in enumerate(pdf.pages, 1):
                texto_pagina = pagina.extract_text() or ""
                print(f"   Página {i}: {len(texto_pagina)} caracteres")
                texto += texto_pagina
            
            if texto.strip():  # Se conseguiu extrair texto
                print(f"✓ Total extraído com pdfplumber: {len(texto)} caracteres")
                return texto
            else:
                print("⚠️  pdfplumber não extraiu nenhum texto")
    except Exception as e:
        print(f"⚠️  pdfplumber falhou: {type(e).__name__}: {e}")
    
    # Tentativa 2: PyPDF2 (fallback)
    try:
        import PyPDF2
        print("🔄 Tentando com PyPDF2...")
        with open(caminho, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            texto = ""
            for pagina in reader.pages:
                texto += pagina.extract_text() or ""
            if texto.strip():
                print(f"✓ Texto extraído com PyPDF2: {len(texto)} caracteres")
                return texto
            else:
                print("⚠️  PyPDF2 não extraiu nenhum texto")
    except ImportError:
        print("❌ PyPDF2 não está instalado. Execute: pip install PyPDF2")
    except Exception as e:
        print(f"❌ PyPDF2 também falhou: {type(e).__name__}: {e}")
    
    print("❌ Nenhum método conseguiu extrair texto do PDF")
    return None

def normalizar_uc(uc_raw: str) -> str:
    """Remove caracteres não numéricos para comparação."""
    if not uc_raw:
        return ""
    return re.sub(r'[^\d]', '', uc_raw)

def extrai_uc(nome_arquivo: str):
    """Extrai UC do nome do arquivo no formato {UC}_{NOME}_{DATA}.pdf"""
    # Tenta extrair UC antes do primeiro underscore
    partes = nome_arquivo.split('_')
    if partes and partes[0]:
        primeira_parte = partes[0]
        uc_candidata = re.sub(r'[^\d]', '', primeira_parte)
        
        if 6 <= len(uc_candidata) <= 12:
            return uc_candidata
    
    # Fallback: Busca padrão antigo
    nome_limpo = normalizar_uc(nome_arquivo)
    match = re.search(r'\d{6,12}', nome_limpo)
    if match:
        return match.group(0)
    
    return None

def extrai_uc_do_texto(texto: str, nome_arquivo: str = ""):
    """Extrai UC do conteúdo do PDF."""
    if not texto:
        return None
        
    for padrao in config.UC_PATTERNS:
        matches = re.finditer(padrao, texto, re.IGNORECASE)
        for match in matches:
            uc_encontrada = match.group(1)
            uc_limpa = normalizar_uc(uc_encontrada)
            
            if 6 <= len(uc_limpa) <= 12:
                return uc_limpa
                
    return None

def extrai_referencia(texto: str, nome_arquivo: str = ""):
    """Extrai mês/ano de referência (ex: nov/2025)."""
    match = re.search(r'(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[a-z]*\W+(\d{4})', texto, re.IGNORECASE)
    if match:
        return f"{match.group(1).lower()}/{match.group(2)}"
    return None

def _str_to_float(valor_str: str):
    """Converte '1.234,56' para 1234.56"""
    try:
        limpo = valor_str.replace('.', '').replace(',', '.')
        return float(limpo)
    except:
        return None

def extrair_valor_fatura(texto: str, nome_arquivo: str = ""):
    """Tenta extrair o 'Total a Pagar' da fatura."""
    # Tenta padrão principal
    match = re.search(config.VALUE_PATTERNS['fatura_total'], texto, re.IGNORECASE)
    if match:
        return _str_to_float(match.group(1))
    
    # Tenta padrão alternativo
    match = re.search(config.VALUE_PATTERNS['fatura_total_alt'], texto, re.IGNORECASE)
    if match:
        return _str_to_float(match.group(1))
        
    return None

def extrair_valor_boleto(texto: str, nome_arquivo: str = ""):
    """Extrai valor do boleto."""
    # Tentativa via Rótulo (padrão principal)
    match_doc = re.search(config.VALUE_PATTERNS['boleto_documento'], texto, re.IGNORECASE)
    if match_doc:
        valor = _str_to_float(match_doc.group(1))
        if valor:
            return valor
    
    # Tentativa via Rótulo (padrão alternativo)
    match_doc_alt = re.search(config.VALUE_PATTERNS['boleto_documento_alt'], texto, re.IGNORECASE)
    if match_doc_alt:
        valor = _str_to_float(match_doc_alt.group(1))
        if valor:
            return valor

    # Tentativa via Linha Digitável
    numeros_apenas = re.sub(r'[^\d]', '', texto)
    match_barras = re.search(r'(\d{44,48})', numeros_apenas)
    
    if match_barras:
        sequencia = match_barras.group(1)
        valor_str = sequencia[-10:]
        try:
            valor = float(valor_str) / 100.0
            if valor > 0:
                return valor
        except:
            pass
            
    return None

def diagnosticar_arquivo(caminho_pdf: str, tipo: str):
    """Diagnostica um único arquivo PDF."""
    log = get_logger()
    
    print(f"\n{'='*80}")
    print(f"📄 DIAGNÓSTICO: {Path(caminho_pdf).name}")
    print(f"{'='*80}\n")
    
    # 1. Extração de texto
    log.print_section("1️⃣  EXTRAÇÃO DE TEXTO")
    texto = ler_pdf(caminho_pdf)
    
    if not texto:
        log.print_error("❌ Falha ao extrair texto do PDF!")
        return None
    
    print(f"✓ Texto extraído: {len(texto)} caracteres")
    print(f"✓ Linhas: {texto.count(chr(10))}")
    print(f"\n📝 Primeiros 500 caracteres:")
    print(f"{'-'*80}")
    print(texto[:500])
    print(f"{'-'*80}\n")
    
    # 2. Extração de UC
    log.print_section("2️⃣  EXTRAÇÃO DE UC")
    
    nome_arquivo = Path(caminho_pdf).name
    uc_nome = extrai_uc(nome_arquivo)
    uc_texto = extrai_uc_do_texto(texto, nome_arquivo)
    
    print(f"📛 Nome do arquivo: {nome_arquivo}")
    print(f"🔢 UC do nome: {uc_nome if uc_nome else '❌ NÃO ENCONTRADA'}")
    print(f"🔢 UC do texto: {uc_texto if uc_texto else '❌ NÃO ENCONTRADA'}")
    
    if uc_nome and uc_texto:
        uc_nome_norm = normalizar_uc(uc_nome)
        uc_texto_norm = normalizar_uc(uc_texto)
        
        if uc_nome_norm == uc_texto_norm:
            print(f"✅ UC normalizada coincide: {uc_nome_norm}")
        else:
            print(f"⚠️  UC divergente!")
            print(f"   Nome normalizado: {uc_nome_norm}")
            print(f"   Texto normalizado: {uc_texto_norm}")
    
    # 3. Extração de Valor
    log.print_section("3️⃣  EXTRAÇÃO DE VALOR")
    
    if tipo == 'fatura':
        valor = extrair_valor_fatura(texto, nome_arquivo)
    else:
        valor = extrair_valor_boleto(texto, nome_arquivo)
    
    if valor:
        print(f"💰 Valor encontrado: R$ {valor:.2f}")
    else:
        print(f"❌ Valor NÃO encontrado")
        print(f"\n🔍 Buscando 'R$' no texto...")
        import re
        valores_rs = re.findall(r'R\$\s*([\d\.,]+)', texto)
        if valores_rs:
            print(f"   Valores R$ encontrados no texto: {valores_rs[:5]}")
        else:
            print(f"   Nenhum 'R$' encontrado no texto!")
    
    # 4. Extração de Referência
    log.print_section("4️⃣  EXTRAÇÃO DE PERÍODO")
    
    referencia = extrai_referencia(texto, nome_arquivo)
    
    if referencia:
        print(f"📅 Período encontrado: {referencia}")
    else:
        print(f"❌ Período NÃO encontrado")
    
    # 5. Resumo
    log.print_section("📊 RESUMO")
    
    resultado = {
        'arquivo': nome_arquivo,
        'tipo': tipo,
        'uc_nome': uc_nome,
        'uc_texto': uc_texto,
        'uc_normalizada': normalizar_uc(uc_nome) if uc_nome else normalizar_uc(uc_texto) if uc_texto else None,
        'valor': valor,
        'referencia': referencia,
        'texto_length': len(texto)
    }
    
    print(f"Tipo: {tipo.upper()}")
    print(f"UC: {resultado['uc_normalizada'] if resultado['uc_normalizada'] else '❌ AUSENTE'}")
    print(f"Valor: R$ {valor:.2f}" if valor else "❌ AUSENTE")
    print(f"Período: {referencia if referencia else '❌ AUSENTE'}")
    
    return resultado

def diagnosticar_pareamento(fatura_path: str, boleto_path: str):
    """Diagnostica o pareamento entre uma fatura e um boleto."""
    log = get_logger()
    
    log.print_section("🔗 DIAGNÓSTICO DE PAREAMENTO")
    
    print("Analisando FATURA...")
    fatura = diagnosticar_arquivo(fatura_path, 'fatura')
    
    print("\n" + "="*80 + "\n")
    
    print("Analisando BOLETO...")
    boleto = diagnosticar_arquivo(boleto_path, 'boleto')
    
    if not fatura or not boleto:
        log.print_error("Falha ao processar um ou ambos os arquivos!")
        return
    
    # Comparação
    log.print_section("🔍 COMPARAÇÃO")
    
    print(f"{'Critério':<20} {'Fatura':<30} {'Boleto':<30} {'Match?':<10}")
    print(f"{'-'*90}")
    
    # UC
    uc_match = fatura['uc_normalizada'] == boleto['uc_normalizada'] if fatura['uc_normalizada'] and boleto['uc_normalizada'] else False
    print(f"{'UC Normalizada':<20} {str(fatura['uc_normalizada']):<30} {str(boleto['uc_normalizada']):<30} {'✅ SIM' if uc_match else '❌ NÃO':<10}")
    
    # Valor
    valor_match = False
    if fatura['valor'] and boleto['valor']:
        diff = abs(fatura['valor'] - boleto['valor'])
        valor_match = diff < 0.01
        fatura_val = f"R$ {fatura['valor']:.2f}"
        boleto_val = f"R$ {boleto['valor']:.2f}"
        match_msg = '✅ SIM' if valor_match else f'❌ NÃO (diff: R$ {diff:.2f})'
        print(f"{'Valor':<20} {fatura_val:<30} {boleto_val:<30} {match_msg:<10}")
    else:
        print(f"{'Valor':<20} {str(fatura['valor']):<30} {str(boleto['valor']):<30} {'❌ AUSENTE':<10}")
    
    # Período
    periodo_match = fatura['referencia'] == boleto['referencia'] if fatura['referencia'] and boleto['referencia'] else False
    print(f"{'Período':<20} {str(fatura['referencia']):<30} {str(boleto['referencia']):<30} {'✅ SIM' if periodo_match else '❌ NÃO':<10}")
    
    # Resultado final
    print(f"\n{'-'*90}\n")
    
    if uc_match and valor_match and periodo_match:
        log.print_success("✅ PAREAMENTO VÁLIDO - Todos os critérios coincidem!")
    elif uc_match:
        print("⚠️  PAREAMENTO PARCIAL - UC coincide, mas:")
        if not valor_match:
            print("   ❌ Valores divergem")
        if not periodo_match:
            print("   ❌ Períodos divergem")
    else:
        log.print_error("❌ PAREAMENTO INVÁLIDO - UC não coincide!")

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) == 2:
        # Modo: diagnosticar um único arquivo
        arquivo = sys.argv[1]
        tipo = input("Tipo do arquivo (fatura/boleto): ").strip().lower()
        diagnosticar_arquivo(arquivo, tipo)
    
    elif len(sys.argv) == 3:
        # Modo: diagnosticar pareamento
        fatura = sys.argv[1]
        boleto = sys.argv[2]
        diagnosticar_pareamento(fatura, boleto)
    
    else:
        print("Uso:")
        print("  python diagnostico.py <arquivo.pdf> - Diagnostica um arquivo")
        print("  python diagnostico.py <fatura.pdf> <boleto.pdf> - Diagnostica pareamento")
