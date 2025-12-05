import os
import logging

def buscar_pdf_uc(pasta_pdfs, uc_normalizada, mes_ref):
    """Busca o PDF correspondente a uma UC (aceita novos e antigos formatos)."""
    logging.info(f"Buscando por PDF da UC: {uc_normalizada} (Mês: {mes_ref})")
    
    # 1. Tentar encontrar pelo formato NOVO (padrão Unificador.py atualizado)
    # Padrão: UC_{uc}_{cliente}_{mes}_{ano}.pdf
    for raiz, _, arquivos in os.walk(pasta_pdfs):
        for nome in arquivos:
            if nome.lower().endswith('.pdf'):
                # Verificar se é o padrão novo (começa com UC_ e contém a UC)
                if nome.lower().startswith(f"uc_{uc_normalizada.lower()}_") and len(nome) > 10:
                    caminho = os.path.join(raiz, nome)
                    logging.info(f"PDF encontrado (formato novo): {os.path.basename(caminho)}")
                    return caminho
    
    # 2. Tentar encontrar pelo formato ANTIGO (compatibilidade)
    # Padrão antigo: {uc}_{mes_ano}.pdf -> ex: 1052027_092025.pdf
    nome_arquivo_esperado = f"{uc_normalizada}_{mes_ref[4:]}{mes_ref[:4]}.pdf".lower()
    
    for raiz, _, arquivos in os.walk(pasta_pdfs):
        for nome in arquivos:
            if nome.lower() == nome_arquivo_esperado:
                caminho = os.path.join(raiz, nome)
                logging.info(f"PDF encontrado (formato antigo): {os.path.basename(caminho)}")
                return caminho

    # 3. Busca mais flexível: encontrar qualquer arquivo que contenha a UC
    for raiz, _, arquivos in os.walk(pasta_pdfs):
        for nome in arquivos:
            if nome.lower().endswith('.pdf') and uc_normalizada.lower() in nome.lower():
                caminho = os.path.join(raiz, nome)
                logging.info(f"PDF encontrado (busca flexível): {os.path.basename(caminho)}")
                return caminho

    logging.warning(f"UC {uc_normalizada}: nenhum PDF encontrado em '{pasta_pdfs}' (testados formatos novo, antigo e busca flexível).")
    return None
