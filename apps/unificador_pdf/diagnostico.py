"""
Script de diagnóstico para identificar problemas no pareamento de PDFs.
Execute este script na pasta onde estão seus PDFs para ver detalhes do processamento.
"""

import sys
from pathlib import Path

# Adiciona o caminho do módulo
sys.path.insert(0, str(Path(__file__).parent))

from extractors.uc_extractor import extrai_uc, extrai_uc_do_texto, normalizar_uc, extrai_referencia
from extractors.value_extractor import extrair_valor_fatura, extrair_valor_boleto
from pdf.pdf_reader import ler_pdf
from logging_utils import get_logger

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
        print(f"{'Valor':<20} {f'R$ {fatura['valor']:.2f}':<30} {f'R$ {boleto['valor']:.2f}':<30} {'✅ SIM' if valor_match else f'❌ NÃO (diff: R$ {diff:.2f})':<10}")
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
