import os
import re
import logging
import zipfile
import win32com.client
import pythoncom
from pywintypes import com_error
from datetime import datetime, timedelta
from io import BytesIO

from .config import (
    NOME_CONTA_OUTLOOK, PASTA_SAIDA_BOLETOS, PASTA_SAIDA_FALHAS,
    DOMINIO_REMETENTE_VALIDO, PASTAS_BANIDAS
)
from .utils import normaliza, _to_bytes, _iso, hash_bytes
from .pdf_processor import extrair_uc_do_pdf, extrair_nome_do_pdf
from .file_manager import salvar_bytes, carregar_hashes_existentes

def get_sender_smtp(message):
    try:
        prop = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
        sender_email = message.PropertyAccessor.GetProperty(prop)
        if sender_email and "@" in sender_email: return sender_email.lower()
        if message.SenderEmailType == "EX":
            try:
                sender = message.Sender.GetExchangeUser()
                if sender and sender.PrimarySmtpAddress: return sender.PrimarySmtpAddress.lower()
            except Exception: pass
        return (message.SenderEmailAddress or "").lower()
    except Exception: return ""

def percorrer_e_processar_pasta(pasta, dt_inicio, dt_fim, status_callback, progress_callback, arquivos_salvos, hashes_salvos):
    sucesso_total, falha_total = 0, 0
    status_callback(f"Analisando pasta: {pasta.FolderPath}")
    logging.info(f"--- ANALISANDO PASTA: {pasta.FolderPath} ---")
    
    motivos = {"nao_mail": 0, "fora_periodo": 0, "remetente": 0, "corpo": 0, "sem_anexo_valido": 0, "duplicata_hash": 0, "falha_uc": 0}

    try:
        items = pasta.Items
        items.Sort("[ReceivedTime]", True)
        
        dt_inicio_sql = (dt_inicio - timedelta(days=1)); dt_fim_sql = (dt_fim + timedelta(days=1))
        
        filtro_data_dasl = (f"@SQL=(\"urn:schemas:httpmail:datereceived\" >= '{_iso(dt_inicio_sql)}' AND "
                          f"\"urn:schemas:httpmail:datereceived\" <= '{_iso(dt_fim_sql)}' AND "
                          f"\"urn:schemas:httpmail:messageclass\" = 'IPM.Note')")
        
        items = items.Restrict(filtro_data_dasl)
        total_filtrado = items.Count
        status_callback(f"   {total_filtrado} e-mails no período (pré-filtro por data). Processando...")
        logging.info(f"{total_filtrado} e-mails pré-filtrados (apenas por data) encontrados na pasta.")

    except Exception as e:
        logging.error(f"Erro ao aplicar filtro de data MAPI em '{pasta.Name}': {e}. Usando fallback (mais lento).")
        items = pasta.Items; total_filtrado = items.Count

    progress_callback(0, total_filtrado, "")
    for i, item in enumerate(items):
        progress_callback(i + 1, total_filtrado, f"Analisando e-mail {i+1} de {total_filtrado}")
        try:
            email_id = f"Assunto: '{getattr(item, 'Subject', 'N/A')}'"
            
            if getattr(item, "Class", 0) != 43:
                motivos["nao_mail"] += 1; continue

            received_time = item.ReceivedTime.replace(tzinfo=None)
            if not (dt_inicio <= received_time <= dt_fim):
                motivos["fora_periodo"] += 1; continue
            
            sender = get_sender_smtp(item)
            if not sender.endswith(DOMINIO_REMETENTE_VALIDO):
                motivos["remetente"] += 1; continue
            
            corpo_raw = (getattr(item, "HTMLBody", "") or "") or (item.Body or "")
            if "solicitacao de pagamento" not in normaliza(corpo_raw):
                motivos["corpo"] += 1; continue

            anexo_alvo_bytes = None
            if item.Attachments.Count > 0:
                for att in item.Attachments:
                    fname = str(att.FileName or "").lower()
                    if (fname.startswith("boleto") and fname.endswith(".pdf")):
                        anexo_alvo_bytes = _to_bytes(att.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37010102"))
                        if anexo_alvo_bytes: break
                    
                    if fname.endswith(".zip"):
                        try:
                            zbytes_data = _to_bytes(att.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37010102"))
                            if zbytes_data:
                                with zipfile.ZipFile(BytesIO(zbytes_data)) as zf:
                                    for n in zf.namelist():
                                        if n.lower().startswith("boleto") and n.lower().endswith(".pdf"):
                                            anexo_alvo_bytes = zf.read(n); break
                        except (zipfile.BadZipFile, RuntimeError) as e:
                            logging.warning(f"ZIP inválido/protegido '{fname}' no e-mail {email_id}: {e}")
                        if anexo_alvo_bytes: break

            if anexo_alvo_bytes:
                h = hash_bytes(anexo_alvo_bytes)
                if h in hashes_salvos:
                    motivos["duplicata_hash"] += 1; continue
                
                logging.info(f"PROCESSANDO ANEXO VÁLIDO de {email_id}")
                uc = extrair_uc_do_pdf(BytesIO(anexo_alvo_bytes))
                
                if uc:
                    # Extrair nome do cliente para o arquivo
                    nome_cliente = extrair_nome_do_pdf(BytesIO(anexo_alvo_bytes))
                    nome_cliente_safe = re.sub(r'[\\/*?:"<>|]', "", nome_cliente)[:50].strip()
                    
                    # Data do recebimento do e-mail para evitar duplicatas de competência
                    data_email_str = item.ReceivedTime.strftime("%Y%m%d")
                    
                    nome_arquivo = f"{uc}_{nome_cliente_safe}_{data_email_str}.pdf"
                    caminho = os.path.join(PASTA_SAIDA_BOLETOS, nome_arquivo)
                    
                    if salvar_bytes(caminho, anexo_alvo_bytes):
                        logging.info(f"-> SUCESSO: Boleto salvo em: {caminho}")
                        arquivos_salvos.add(nome_arquivo); hashes_salvos.add(h); sucesso_total += 1
                    else: logging.warning(f"-> AVISO: O arquivo com nome '{nome_arquivo}' já existe no disco.")
                else:
                    motivos["falha_uc"] += 1; falha_total += 1
                    timestamp = item.ReceivedTime.strftime("%Y%m%d_%H%M%S"); safe_subject = re.sub(r'[\\/*?:"<>|]', "", item.Subject)[:50]
                    nome_arquivo_falha = f"{timestamp}_{safe_subject}.pdf"; caminho_falha = os.path.join(PASTA_SAIDA_FALHAS, nome_arquivo_falha)
                    if salvar_bytes(caminho_falha, anexo_alvo_bytes):
                        logging.warning(f"-> FALHA DE UC: Salvo para análise em: {caminho_falha}")
            else:
                motivos["sem_anexo_valido"] += 1
        except Exception as e_item:
            logging.error(f"Erro inesperado ao processar um item individual: {e_item}")

    logging.info(f"RESUMO DA PASTA '{pasta.Name}': Sucesso={sucesso_total}, Falha={falha_total}, Descartes={motivos}")
    if sum(motivos.values()) > 0: status_callback(f"Descartes em '{pasta.Name}': {sum(motivos.values())} (Ver log)")

    for subpasta in pasta.Folders:
        if str(subpasta.Name).strip().lower() in PASTAS_BANIDAS:
            logging.info(f"Ignorando pasta banida: {subpasta.FolderPath}")
            continue
        s, f = percorrer_e_processar_pasta(subpasta, dt_inicio, dt_fim, status_callback, progress_callback, arquivos_salvos, hashes_salvos)
        sucesso_total += s; falha_total += f
    return sucesso_total, falha_total

def buscar_e_salvar_boletos(data_inicio_str, data_fim_str, status_callback, progress_callback, completion_callback):
    pythoncom.CoInitialize()
    try:
        os.makedirs(PASTA_SAIDA_BOLETOS, exist_ok=True); os.makedirs(PASTA_SAIDA_FALHAS, exist_ok=True)
        dt_inicio = datetime.strptime(data_inicio_str, "%d/%m/%Y")
        dt_fim = datetime.strptime(data_fim_str, "%d/%m/%Y").replace(hour=23, minute=59, second=59)
        logging.info(f"Iniciando busca de {data_inicio_str} a {data_fim_str}.")
        status_callback("Conectando ao Outlook...")
        outlook_ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        caixa_entrada = None
        # Tentar obter caixa de entrada padrão primeiro (mais robusto)
        try:
            default_inbox = outlook_ns.GetDefaultFolder(6) # 6 = Inbox
            # Verificar se a conta padrão é a que queremos
            if NOME_CONTA_OUTLOOK.lower() in default_inbox.Parent.Name.lower():
                caixa_entrada = default_inbox
                conta_alvo = default_inbox.Parent
                logging.info(f"Usando Caixa de Entrada padrão da conta '{conta_alvo.Name}'.")
        except com_error:
            logging.warning("Não foi possível acessar a Caixa de Entrada padrão diretamente.")

        # Fallback: Procurar conta por nome se não for a padrão
        if not caixa_entrada:
            conta_alvo = next((c for c in outlook_ns.Folders if str(c.Name).strip().lower() == NOME_CONTA_OUTLOOK.lower()), None)
            if not conta_alvo: raise Exception(f"Conta '{NOME_CONTA_OUTLOOK}' não encontrada.")
            
            caixa_entrada = next((f for f in conta_alvo.Folders if str(f.Name).strip().lower() in ("caixa de entrada", "inbox")), None)
            if not caixa_entrada: raise Exception("Não foi possível localizar a 'Caixa de Entrada' pelo nome.")
            
        logging.info(f"Conectado à conta '{conta_alvo.Name}'.")
        status_callback("Verificando boletos já salvos para evitar duplicatas...")
        arquivos_salvos = set(os.listdir(PASTA_SAIDA_BOLETOS))
        hashes_salvos = carregar_hashes_existentes(PASTA_SAIDA_BOLETOS)
        status_callback(f"--- Iniciando verificação em '{conta_alvo.Name}' ---")
        sucesso, falha = percorrer_e_processar_pasta(caixa_entrada, dt_inicio, dt_fim, status_callback, progress_callback, arquivos_salvos, hashes_salvos)
        
        logging.info("--- PROCESSO FINALIZADO ---")
        logging.info(f"Total de boletos com UC identificada: {sucesso}")
        logging.info(f"Total de boletos sem UC (para análise): {falha}")
        logging.info("="*50 + "\n")
        completion_callback(sucesso, falha)
    except com_error as e:
        logging.error(f"ERRO COM ESPECÍFICO no processo principal: {e}", exc_info=True)
        status_callback(f"\nERRO COM (Outlook): {e}\nVeja o log.")
        completion_callback(0, 0, erro=True)
    except Exception as e:
        logging.error(f"ERRO CRÍTICO no processo principal: {e}", exc_info=True)
        status_callback(f"\nERRO CRÍTICO: {e}\nVeja o log.")
        completion_callback(0, 0, erro=True)
    finally:
        pythoncom.CoUninitialize()
