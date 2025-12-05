import os
import shutil
from io import BytesIO
from tkinter import messagebox

from EGS_Suite.common.logging import get_logger
from .config import PASTA_SAIDA_BASE, PASTA_LOGS, PASTA_SAIDA_BOLETOS, PASTA_SAIDA_FALHAS
from .utils import hash_bytes
from .pdf_processor import extrair_uc_do_pdf

logger = get_logger('buscador_boletos')

def carregar_hashes_existentes(pasta):
    hs = set()
    if not os.path.exists(pasta): return hs
    for nome in os.listdir(pasta):
        if nome.lower().endswith(".pdf"):
            try:
                with open(os.path.join(pasta, nome), "rb") as f:
                    hs.add(hash_bytes(f.read()))
            except Exception as e:
                logger.warning(f"Falha ao carregar hash de '{nome}': {e}")
    logger.info(f"{len(hs)} hashes de PDFs existentes foram carregados.")
    return hs

def salvar_bytes(caminho, dados):
    if not os.path.exists(caminho):
        with open(caminho, "wb") as f: f.write(dados)
        return True
    return False

def corrigir_pdfs_antigos(pasta_origem=PASTA_SAIDA_FALHAS, pasta_destino=PASTA_SAIDA_BOLETOS):
    os.makedirs(pasta_destino, exist_ok=True)
    total, ok, fail, ja_existia = 0, 0, 0, 0

    if not os.path.exists(pasta_origem):
        return "Pasta de origem 'boletos_sem_uc' não encontrada."

    for nome in os.listdir(pasta_origem):
        if not nome.lower().endswith(".pdf"):
            continue
        total += 1
        caminho = os.path.join(pasta_origem, nome)
        try:
            with open(caminho, "rb") as f:
                dados = f.read()
            uc = extrair_uc_do_pdf(BytesIO(dados))
            if uc:
                novo_nome = f"{uc}.pdf"
                destino = os.path.join(pasta_destino, novo_nome)
                if not os.path.exists(destino):
                    shutil.move(caminho, destino)
                    ok += 1
                    logger.info(f"[corrigir] Movido: {nome} -> {novo_nome}")
                else:
                    ja_existia += 1
                    os.remove(caminho)
                    logger.warning(f"[corrigir] Destino '{novo_nome}' já existia. Duplicata '{nome}' removida.")
            else:
                fail += 1
                logger.warning(f"[corrigir] Sem UC em '{nome}'. Arquivo mantido na pasta de falhas.")
        except Exception as e:
            fail += 1
            logger.error(f"[corrigir] Erro em '{nome}': {e}")

    resumo = (f"Total de arquivos verificados: {total}\n"
              f"✅ Renomeados com sucesso: {ok}\n"
              f"⚠️ Duplicatas encontradas (removidas): {ja_existia}\n"
              f"❌ Falhas na extração (permanecem na pasta): {fail}")
    return resumo
