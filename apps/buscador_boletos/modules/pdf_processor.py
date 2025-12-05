import re
import logging
import PyPDF2
from .config import SENHAS_COMUNS

def extrair_uc_do_pdf(stream_do_pdf):
    try:
        stream_do_pdf.seek(0)
        reader = PyPDF2.PdfReader(stream_do_pdf)
        
        if reader.is_encrypted:
            decrypted = False
            for senha in SENHAS_COMUNS:
                if reader.decrypt(senha) != 0:
                    decrypted = True
                    logging.info(f"PDF descriptografado com a senha: '{senha if senha else 'Vazia'}'.")
                    break
            if not decrypted:
                logging.warning("PDF criptografado, não foi possível abri-lo.")
                return None

        # --- MELHORIA: Regex em duas etapas ---
        padrao_uc_especifico = re.compile(r'(?:UC|Unidade\s+Consumidora)[:\s]*?(10\s*[/\-]\s*\d{3,9}\s*[-\s]*\d)', re.I | re.S)
        padrao_uc_geral = re.compile(r'(10\s*[/\-]\s*\d{3,9}\s*[-\s]*\d)', re.S)
        
        texto_total = "\n".join((page.extract_text() or "") for page in reader.pages)

        if not texto_total.strip():
             logging.warning("PDF parece ser uma imagem (sem texto extraível).")
             return None

        # 1ª Tentativa: Padrão específico (mais seguro)
        m = padrao_uc_especifico.search(texto_total)
        if m:
            uc_limpa = re.sub(r'\D', '', m.group(1))
            logging.info(f"UC '{uc_limpa}' encontrada com padrão específico.")
            return uc_limpa

        # 2ª Tentativa: Padrão geral (fallback)
        m = padrao_uc_geral.search(texto_total)
        if m:
            uc_limpa = re.sub(r'\D', '', m.group(1))
            logging.info(f"UC '{uc_limpa}' encontrada com padrão geral (fallback).")
            return uc_limpa
            
    except Exception as e:
        logging.error(f"Erro ao ler PDF: {e}")
    
    logging.warning("Nenhuma UC encontrada no PDF.")
    return None

def extrair_nome_do_pdf(stream_do_pdf):
    try:
        stream_do_pdf.seek(0)
        reader = PyPDF2.PdfReader(stream_do_pdf)
        
        # Tentar extrair da primeira página (onde geralmente está o Pagador)
        texto_pagina_1 = reader.pages[0].extract_text()
        
        if not texto_pagina_1:
            return "ClienteDesconhecido"

        # Regex baseado no exemplo do usuário: "Pagador: NOME CNPJ/CPF:"
        # Usa re.DOTALL (re.S) para pegar multilinhas se necessário, mas foca em travar no CNPJ
        padrao_nome = re.compile(r'Pagador:\s*(.*?)\s*CNPJ\/CPF', re.IGNORECASE | re.DOTALL)
        
        m = padrao_nome.search(texto_pagina_1)
        if m:
            nome_sujo = m.group(1).replace('\n', ' ').strip()
            # Limpeza extra para evitar nomes muito longos ou caracteres estranhos
            nome_limpo = re.sub(r'\s+', ' ', nome_sujo) # Remove espaços duplos
            logging.info(f"Nome extraído: '{nome_limpo}'")
            return nome_limpo
            
    except Exception as e:
        logging.error(f"Erro ao extrair nome do PDF: {e}")
    
    return "ClienteDesconhecido"
