import os

# ---------------- Configurações de Busca ----------------
# ---------------- Configurações de Busca ----------------
NOME_CONTA_OUTLOOK = "atendimento@egsenergia.com.br"

# Define a pasta base relativa ao diretório raiz do projeto (onde está o script principal)
# O arquivo config.py está em /modules/config.py, então subimos um nível para chegar na raiz
PROJETO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PASTA_SAIDA_BASE = os.path.join(PROJETO_ROOT, "Boletos_Salvos")
PASTA_LOGS = os.path.join(PROJETO_ROOT, "logs")
PASTA_SAIDA_BOLETOS = os.path.join(PASTA_SAIDA_BASE, "boletos_baixados")
PASTA_SAIDA_FALHAS = os.path.join(PASTA_SAIDA_BASE, "boletos_sem_uc")
SENHAS_COMUNS = ["", "123456", "000000", "pinbank"] 

# --- CRITÉRIOS DE BUSCA ---
DOMINIO_REMETENTE_VALIDO = "@pinbank.com.br"

# --- MELHORIA: Lista de pastas a serem ignoradas na busca recursiva ---
PASTAS_BANIDAS = {
    "lixo eletrônico", "itens excluídos", "spam", "deleted items", "junk email",
    "rascunhos", "drafts", "junk e-mail", "arquivados", "archive", "arquivos mortos",
    "conversas", "conversations"
}
