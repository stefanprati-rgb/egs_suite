import logging
import time
import threading
from .data_manager import carregar_mapa_emails
from .pdf_finder import buscar_pdf_uc
from .email_sender import enviar_email_outlook
from .report_manager import gerar_relatorio_falhas, gerar_relatorio_sucessos

import pythoncom
import win32com.client

def enviar_emails_worker(
    pasta_pdfs, caminho_processado, mes_ref, assunto, corpo, modo_envio,
    update_status, update_progress, on_completion, progress_bar
):
    """Função principal que executa o envio em uma thread separada."""
    pythoncom.CoInitialize() # Inicializa COM para esta thread
    outlook_app = None
    try:
        update_status("--- Iniciando processo de envio em massa (MODULARIZADO E OTIMIZADO) ---")

        # Cria a instância do Outlook UMA VEZ para toda a sessão
        try:
            outlook_app = win32com.client.Dispatch("Outlook.Application")
        except Exception as e:
            raise RuntimeError(f"Não foi possível iniciar o Outlook: {e}")
        
        try:
            update_status("1. Carregando e cruzando informações...")
            email_map, falhas_preparacao = carregar_mapa_emails(caminho_processado, update_status)
        except Exception as e:
            # A função de conclusão (on_completion) agora para a barra de progresso
            update_status(f"ERRO FATAL: {e}. Consulte o log.")
            logging.critical(f"ERRO FATAL ao carregar arquivos: {e}", exc_info=True)
            on_completion(0, 0, 0, None, None)
            return

        # Para a animação e volta ao modo normal (determinado)
        progress_bar.stop()
        progress_bar.config(mode='determinate')
        update_progress(0, 1) # Reseta para 0%

        total_ucs = len(email_map)
        update_status(f"2. {total_ucs} UCs válidas para envio. Iniciando busca de PDFs e disparos...")
        
        sucessos = 0
        falhas_processamento = []
        sucessos_lista = []
        salvar_somente = (modo_envio == "Salvar (Rascunhos)")
        mes_ref_formatado = mes_ref.replace('-', '') # AAAA-MM -> AAAAMM

        for i, (uc_normalizada, dados_uc) in enumerate(email_map.items()):
            progresso_atual = i + 1
            nome_cliente = dados_uc.get('Nome/Razão Social', 'N/A')
            email_destino = dados_uc.get('E-mail', 'N/A')

            update_status(f"Processando {progresso_atual}/{total_ucs}: UC {uc_normalizada} ({nome_cliente})")
            
            caminho_pdf = buscar_pdf_uc(pasta_pdfs, uc_normalizada, mes_ref_formatado)
            
            if not caminho_pdf:
                falhas_processamento.append({
                    'UC': uc_normalizada, 'Nome Cliente': nome_cliente, 'E-mail': email_destino,
                    'Motivo': 'PDF não encontrado na pasta (verificar nome do arquivo)'
                })
                update_progress(progresso_atual, total_ucs)
                continue

            try:
                corpo_final = corpo.replace('{CLIENTE_NOME}', nome_cliente).replace('{UC_CODIGO}', uc_normalizada).replace('{MES_REF}', mes_ref)
                assunto_final = assunto.replace('{MES_REF}', mes_ref)
                
                # Passa a instância do outlook_app
                sucesso, status_msg = enviar_email_outlook(
                    outlook_app, email_destino, assunto_final, corpo_final, caminho_pdf, salvar_somente=salvar_somente
                )
                
                if sucesso:
                    sucessos += 1
                    sucessos_lista.append({
                        'UC': uc_normalizada, 'Nome Cliente': nome_cliente, 'E-mail': email_destino,
                        'Status': status_msg
                    })
                else:
                    falhas_processamento.append({
                        'UC': uc_normalizada, 'Nome Cliente': nome_cliente, 'E-mail': email_destino,
                        'Motivo': f'Falha no envio/salvamento via Outlook: {status_msg}'
                    })
            except Exception as e:
                falhas_processamento.append({
                    'UC': uc_normalizada, 'Nome Cliente': nome_cliente, 'E-mail': email_destino,
                    'Motivo': f'Erro inesperado no script: {e}'
                })
                
            update_progress(progresso_atual, total_ucs)
            time.sleep(0.1)

        update_status("3. Finalizando e gerando relatórios...")
        todas_as_falhas = falhas_preparacao + falhas_processamento
        caminho_relatorio_falhas = gerar_relatorio_falhas(todas_as_falhas)
        caminho_relatorio_sucesso = gerar_relatorio_sucessos(sucessos_lista)
        
        update_status("--- FIM DO PROCESSO ---")
        on_completion(total_ucs, sucessos, len(todas_as_falhas), caminho_relatorio_falhas, caminho_relatorio_sucesso)

    finally:
        pythoncom.CoUninitialize() # Libera recursos COM
