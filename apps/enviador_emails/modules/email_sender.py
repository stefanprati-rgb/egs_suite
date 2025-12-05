import os
import win32com.client
import pythoncom
import logging
from .config import EMAIL_REMETENTE

def enviar_email_outlook(outlook_app, email_destino, assunto, corpo, anexo_path, conta_remetente=EMAIL_REMETENTE, salvar_somente=False):
    """Envia um e-mail através de uma instância existente do Outlook."""
    try:
        mail_item = outlook_app.CreateItem(0) # 0 = olMailItem

        sender_account = next((acc for acc in outlook_app.Session.Accounts if acc.SmtpAddress.lower() == conta_remetente.lower()), None)
        if not sender_account:
            raise ValueError(f"Conta de remetente '{conta_remetente}' não configurada no Outlook.")
        
        mail_item.SendUsingAccount = sender_account
        mail_item.To = email_destino
        mail_item.Subject = assunto
        mail_item.HTMLBody = corpo
        
        if anexo_path and os.path.exists(anexo_path):
            mail_item.Attachments.Add(anexo_path)
            
        # Garante que os emails são válidos
        mail_item.Recipients.ResolveAll()
        
        if salvar_somente:
            mail_item.Save()
            status = "SALVO (Rascunhos)"
        else:
            mail_item.Send()
            status = "ENVIADO"
        
        logging.info(f"E-mail para {email_destino} - {status} com sucesso.")
        return True, status
    
    except Exception as e:
        logging.error(f"Falha ao processar e-mail para {email_destino}. Erro: {e}", exc_info=True)
        return False, f"ERRO ({str(e).strip()})"
