"""
Demo das melhorias visuais do sistema de logging com colorama.
Execute este arquivo para ver as cores e formatações no console.
"""

from logging_utils import get_logger

def demo():
    log = get_logger()
    
    # Seção 1: Extração de PDFs
    log.print_section("📄 EXTRAÇÃO DE PDFS")
    
    log.info("Iniciando extração de documentos...")
    
    # Simulação de progresso
    total_files = 10
    for i in range(1, total_files + 1):
        log.print_progress(i, total_files, f"arquivo_{i}.pdf")
        import time
        time.sleep(0.2)
    
    print()  # Nova linha após progresso
    log.print_success(f"{total_files} arquivos extraídos com sucesso!")
    
    # Seção 2: Validação
    log.print_section("🔍 VALIDAÇÃO DE DADOS")
    
    log.info("Validando UC do nome vs UC do texto...")
    log.success("UC coincide: 1052027")
    
    log.info("Comparando valores...")
    log.success("Valores coincidem: R$ 1.997,44")
    
    log.warning("Período não encontrado em 2 documentos")
    
    # Seção 3: Pareamento
    log.print_section("🔗 PAREAMENTO DE DOCUMENTOS")
    
    log.info("Processando pareamento...")
    log.success("5 pares formados com 100% de confiança")
    log.warning("2 documentos não pareados")
    log.error("1 documento com valores divergentes")
    
    # Resultado final
    log.print_section("📊 RESULTADO FINAL")
    log.print_success("Processamento concluído com sucesso!")
    
    print(f"\n{'-'*70}\n")

if __name__ == "__main__":
    demo()
