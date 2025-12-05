import os
import pandas as pd
import logging
from typing import Dict, Any, Tuple, List
from .utils import normalizar_uc

def carregar_mapa_emails(caminho_processado: str, update_status) -> Tuple[Dict[str, Dict[str, str]], List[Dict[str, Any]]]:
    """
    Carrega e cruza as informações, priorizando a leitura de um arquivo CSV por ser mais rápido.
    """
    # ATUALIZADO: Caminho da base de clientes corrigido.
    caminho_relatorio_fixo_xlsx = r"C:\Users\Stefan_Pratti\GRUPO GERA\Gestão GDC - Documentos\EGS\4 - Base de Clientes\BASE DE CLIENTES - EGS.xlsx"
    # ATUALIZADO: Nome da aba principal da nova planilha.
    nome_aba_relatorio = "base"
    
    # OTIMIZAÇÃO: Define um caminho para a versão CSV do arquivo.
    # O nome do CSV deve ser "BASE DE CLIENTES - EGS.csv" na mesma pasta.
    caminho_relatorio_fixo_csv = os.path.splitext(caminho_relatorio_fixo_xlsx)[0] + ".csv"

    # Colunas identificadas no novo arquivo "base.csv"
    COLUNA_UC = 'INSTALAÇÃO'
    COLUNA_EMAIL = 'E-MAIL'
    COLUNA_NOME = 'NOME COMPLETO OU RAZÃO SOCIAL'

    logging.info(f"Carregando UCs de '{os.path.basename(caminho_processado)}'")
    falhas_carregamento = []

    # 1. Carregar a lista de UCs que FORAM processadas e têm boletos
    try:
        update_status("   - Lendo a planilha de boletos processados...")
        df_processado = pd.read_excel(caminho_processado) if caminho_processado.endswith(('.xlsx', '.xls')) else pd.read_csv(caminho_processado)
        if 'instalacao' not in df_processado.columns:
            raise ValueError("A planilha processada precisa ter uma coluna chamada 'instalacao'.")
        df_processado['UC_NORMALIZADA'] = df_processado['instalacao'].apply(normalizar_uc)
        ucs_a_enviar = set(df_processado['UC_NORMALIZADA'].dropna().unique())
        logging.info(f"Encontradas {len(ucs_a_enviar)} UCs únicas na planilha processada.")
    except Exception as e:
        logging.error(f"Erro ao ler a planilha processada: {e}", exc_info=True)
        raise RuntimeError(f"Não foi possível ler a planilha processada. Verifique o arquivo.\nErro: {e}")

    # 2. Carregar o relatório completo, dando prioridade ao CSV.
    df_infos = None
    try:
        if os.path.exists(caminho_relatorio_fixo_csv):
            update_status("   - Versão CSV encontrada! Lendo arquivo (muito mais rápido)...")
            logging.info(f"Versão CSV da base de clientes encontrada em: {caminho_relatorio_fixo_csv}")
            # Tenta ler com diferentes delimitadores, comum em CSVs no Brasil
            try:
                df_infos = pd.read_csv(caminho_relatorio_fixo_csv, sep=';', encoding='latin-1', low_memory=False)
            except Exception:
                 df_infos = pd.read_csv(caminho_relatorio_fixo_csv, sep=',', low_memory=False)
        else:
            update_status("   - Lendo a base de clientes (.xlsx). Isso pode demorar.")
            update_status(f"   - DICA: Salve a aba '{nome_aba_relatorio}' como CSV na mesma pasta para acelerar.")
            logging.warning(f"Versão CSV não encontrada. Lendo o arquivo .xlsx (lento): {caminho_relatorio_fixo_xlsx}")
            df_infos = pd.read_excel(caminho_relatorio_fixo_xlsx, sheet_name=nome_aba_relatorio)
    except FileNotFoundError:
        logging.error(f"ERRO CRÍTICO: Nenhum arquivo da base de clientes foi encontrado.")
        raise RuntimeError(f"Arquivo da base de clientes não encontrado. Verifique o caminho no código.")
    except Exception as e:
        logging.error(f"Erro ao ler a base de clientes: {e}", exc_info=True)
        raise RuntimeError(f"Não foi possível ler o relatório da base de clientes.\nErro: {e}")

    update_status("   - Cruzando dados e validando e-mails...")
    df_infos.columns = [str(c).strip() for c in df_infos.columns]
    colunas_necessarias = [COLUNA_UC, COLUNA_EMAIL, COLUNA_NOME]
    if not all(col in df_infos.columns for col in colunas_necessarias):
        raise ValueError(f"O relatório precisa das colunas: {', '.join(colunas_necessarias)}. Colunas encontradas: {list(df_infos.columns)}")

    df_infos['UC_NORMALIZADA'] = df_infos[COLUNA_UC].apply(normalizar_uc)

    # 3. Verificar quais UCs da planilha processada não estão no relatório geral
    ucs_no_relatorio = set(df_infos['UC_NORMALIZADA'].dropna().unique())
    ucs_nao_encontradas = ucs_a_enviar - ucs_no_relatorio
    if ucs_nao_encontradas:
        logging.warning(f"AVISO: {len(ucs_nao_encontradas)} UCs da planilha processada não foram encontradas na Base de Clientes.")
        for uc in ucs_nao_encontradas:
            falhas_carregamento.append({'UC': uc, 'Motivo': 'UC da planilha processada não encontrada na Base de Clientes'})

    # 4. Filtrar o relatório para conter apenas UCs que devem ser enviadas e estão presentes
    df_final = df_infos[df_infos['UC_NORMALIZADA'].isin(ucs_a_enviar)].copy()
    
    # 5. Checar por dados faltantes ou inválidos na lista já filtrada
    for _, row in df_final.iterrows():
        uc_norm = row['UC_NORMALIZADA']
        email = row[COLUNA_EMAIL]
        nome = row[COLUNA_NOME]
        motivo = None

        if pd.isna(email) or not isinstance(email, str) or '@' not in email:
            motivo = 'E-mail ausente ou inválido no relatório'
        elif pd.isna(nome):
            motivo = 'Nome/Razão Social ausente no relatório'
        
        if motivo:
            logging.warning(f"UC {uc_norm} ignorada: {motivo}.")
            falhas_carregamento.append({'UC': uc_norm, 'Nome Cliente': nome, 'E-mail': email, 'Motivo': motivo})

    # Renomeia as colunas para o padrão que o resto do script espera, garantindo compatibilidade
    df_final.rename(columns={
        COLUNA_EMAIL: 'E-mail',
        COLUNA_NOME: 'Nome/Razão Social'
    }, inplace=True)

    # 6. Remover duplicatas e linhas com dados faltantes para criar o mapa final
    df_final.dropna(subset=['E-mail', 'Nome/Razão Social'], inplace=True)
    df_final = df_final[df_final['E-mail'].str.contains('@', na=False)]
    df_final.drop_duplicates(subset=['UC_NORMALIZADA'], keep='first', inplace=True)

    email_map = df_final.set_index('UC_NORMALIZADA')[['E-mail', 'Nome/Razão Social']].to_dict(orient='index')
    
    logging.info(f"Mapa de e-mails carregado. {len(email_map)} UCs válidas para envio. {len(falhas_carregamento)} falhas na preparação.")
    return email_map, falhas_carregamento
