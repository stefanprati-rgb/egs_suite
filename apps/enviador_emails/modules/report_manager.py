import os
import pandas as pd
import logging
from typing import List, Dict, Any, Optional
from datetime import datetime
from .config import LOG_FOLDER

def gerar_relatorio_falhas(lista_falhas: List[Dict[str, Any]]) -> Optional[str]:
    """Gera um arquivo CSV com o relatório de falhas."""
    if not lista_falhas:
        logging.info("Nenhuma falha registrada. Relatório de falhas não gerado.")
        return None
    
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f"Relatorio_Falhas_Envio_{timestamp}.csv"
        caminho_relatorio = os.path.join(LOG_FOLDER, nome_arquivo)

        df_relatorio = pd.DataFrame(lista_falhas)
        colunas_ordem = ['UC', 'Nome Cliente', 'E-mail', 'Motivo']
        df_relatorio = df_relatorio.reindex(columns=[c for c in colunas_ordem if c in df_relatorio.columns])
        
        df_relatorio.to_csv(caminho_relatorio, index=False, sep=';', encoding='utf-8-sig')
        
        logging.info(f"Relatório de falhas gerado com sucesso em: {caminho_relatorio}")
        return caminho_relatorio
    except Exception as e:
        logging.error(f"Não foi possível gerar o relatório de falhas. Erro: {e}", exc_info=True)
        return None

def gerar_relatorio_sucessos(lista_sucessos: List[Dict[str, Any]]) -> Optional[str]:
    """Gera um arquivo CSV com o relatório de envios bem-sucedidos."""
    if not lista_sucessos:
        logging.info("Nenhum sucesso registrado. Relatório de sucessos não gerado.")
        return None
    
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f"Relatorio_Sucessos_Envio_{timestamp}.csv"
        caminho_relatorio = os.path.join(LOG_FOLDER, nome_arquivo)

        df_relatorio = pd.DataFrame(lista_sucessos)
        colunas_ordem = ['UC', 'Nome Cliente', 'E-mail', 'Status']
        df_relatorio = df_relatorio.reindex(columns=[c for c in colunas_ordem if c in df_relatorio.columns])
        
        df_relatorio.to_csv(caminho_relatorio, index=False, sep=';', encoding='utf-8-sig')
        
        logging.info(f"Relatório de sucessos gerado com sucesso em: {caminho_relatorio}")
        return caminho_relatorio
    except Exception as e:
        logging.error(f"Não foi possível gerar o relatório de sucessos. Erro: {e}", exc_info=True)
        return None
