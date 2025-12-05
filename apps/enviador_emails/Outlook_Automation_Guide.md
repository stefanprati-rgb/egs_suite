# Guia Exaustivo: Automação do Microsoft Outlook Classic com pywin32

## Índice
1. [Fundamentos e Configuração](#fundamentos-e-configuração)
2. [Manipulação de Itens](#manipulação-de-itens)
3. [Organização e Navegação](#organização-e-navegação)
4. [Funcionalidades Avançadas](#funcionalidades-avançadas)
5. [Melhores Práticas e Desempenho](#melhores-práticas-e-desempenho)
6. [Exemplos Práticos Completos](#exemplos-práticos-completos)
7. [Referência Rápida de Enumerações](#referência-rápida-de-enumerações)

---

## Fundamentos e Configuração

### 1.1 Instalação e Inicialização

#### Requisitos
- **Sistema Operacional**: Windows (Microsoft Outlook Classic/Desktop)
- **Python**: 3.6+
- **Microsoft Outlook**: Instalado e configurado

#### Instalação de Dependências

```bash
pip install pywin32
python Scripts/pywin32_postinstall.py -install
```

O segundo comando registra as bibliotecas COM necessárias no Windows.

#### Importações Essenciais

```python
import win32com.client
import pythoncom
from pywintypes import com_error
from datetime import datetime, timedelta
import os
```

#### Inicialização COM

Para ambientes multi-thread (scripts complexos):

```python
import pythoncom
import threading

def outlook_task():
    # Inicializa COM no thread atual
    pythoncom.CoInitialize()
    
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        # ... seu código aqui ...
    finally:
        pythoncom.CoUninitialize()

# Usar em thread
thread = threading.Thread(target=outlook_task)
thread.start()
```

### 1.2 Conexão Básica com Outlook

```python
import win32com.client
from pywintypes import com_error

def conectar_outlook():
    """
    Estabelece conexão com aplicação do Outlook.
    Retorna: Objeto Application do Outlook
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        return outlook
    except com_error as e:
        print(f"Erro ao conectar com Outlook: {e}")
        return None

def obter_namespace(outlook):
    """
    Obtém o namespace MAPI do Outlook para acesso a pastas.
    """
    try:
        return outlook.GetNamespace("MAPI")
    except com_error as e:
        print(f"Erro ao obter namespace: {e}")
        return None

# Uso
outlook = conectar_outlook()
if outlook:
    ns = obter_namespace(outlook)
```

### 1.3 Enumerações de Tipo de Item

O Outlook trabalha com diferentes tipos de itens identificados por números:

| ItemType | Código | Descrição |
|----------|--------|-----------|
| **MailItem** | 0 | Mensagens de e-mail |
| **ContactItem** | 2 | Contatos |
| **AppointmentItem** | 1 | Reuniões/Compromissos |
| **TaskItem** | 3 | Tarefas |
| **NoteItem** | 5 | Notas |
| **JournalItem** | 4 | Itens de Diário |

---

## Manipulação de Itens

### 2.1 MailItem (Mensagens de E-mail)

#### 2.1.1 Criar e Enviar E-mail Básico

```python
def criar_e_enviar_email(para, assunto, corpo, html=False):
    """
    Cria e envia um e-mail simples.
    
    Args:
        para (str): Destinatário(s) separado por ';'
        assunto (str): Assunto do e-mail
        corpo (str): Corpo da mensagem
        html (bool): Se o corpo é HTML
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        email = outlook.CreateItem(0)  # 0 = MailItem
        
        email.To = para
        email.Subject = assunto
        
        if html:
            email.HTMLBody = corpo
        else:
            email.Body = corpo
        
        email.Send()
        print("E-mail enviado com sucesso!")
        
    except com_error as e:
        print(f"Erro ao enviar e-mail: {e}")

# Uso
criar_e_enviar_email(
    para="usuario@example.com",
    assunto="Relatório Automático",
    corpo="Olá, segue em anexo o relatório.",
    html=False
)
```

#### 2.1.2 E-mail com Cópia e Cópia Oculta

```python
def criar_e_enviar_email_completo(para, cc, cco, assunto, corpo, html=False):
    """
    Cria e-mail com CC e CCO (Cópia Oculta).
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        email = outlook.CreateItem(0)
        
        email.To = para
        email.CC = cc
        email.BCC = cco
        email.Subject = assunto
        
        if html:
            email.HTMLBody = corpo
        else:
            email.Body = corpo
        
        email.Send()
        
    except com_error as e:
        print(f"Erro: {e}")

# Uso
criar_e_enviar_email_completo(
    para="principal@example.com",
    cc="gerente@example.com;cfo@example.com",
    cco="arquivo@example.com",
    assunto="Informação Importante",
    corpo="Conteúdo da mensagem",
    html=False
)
```

#### 2.1.3 E-mail com Anexos

```python
def criar_email_com_anexos(para, assunto, corpo, caminho_anexos):
    """
    Cria e-mail com um ou múltiplos anexos.
    
    Args:
        para (str): Destinatário
        assunto (str): Assunto
        corpo (str): Corpo
        caminho_anexos (list): Lista de caminhos de arquivos
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        email = outlook.CreateItem(0)
        
        email.To = para
        email.Subject = assunto
        email.Body = corpo
        
        # Adicionar anexos
        for caminho in caminho_anexos:
            if os.path.exists(caminho):
                email.Attachments.Add(os.path.abspath(caminho))
            else:
                print(f"Arquivo não encontrado: {caminho}")
        
        email.Send()
        print(f"E-mail enviado com {len(caminho_anexos)} anexo(s)")
        
    except com_error as e:
        print(f"Erro: {e}")

# Uso
criar_email_com_anexos(
    para="usuario@example.com",
    assunto="Relatório Mensal",
    corpo="Segue os relatórios do mês.",
    caminho_anexos=[
        r"C:\Relatorios\janeiro.xlsx",
        r"C:\Relatorios\fevereiro.xlsx"
    ]
)
```

#### 2.1.4 Responder e Encaminhar E-mails

```python
def responder_email(email_obj, corpo_resposta, html=False):
    """
    Responde um e-mail existente.
    
    Args:
        email_obj: Objeto MailItem do Outlook
        corpo_resposta (str): Texto da resposta
        html (bool): Se usar HTML
    """
    try:
        resposta = email_obj.Reply()
        
        if html:
            resposta.HTMLBody = corpo_resposta + "\n" + resposta.HTMLBody
        else:
            resposta.Body = corpo_resposta + "\n" + resposta.Body
        
        resposta.Send()
        print("Resposta enviada!")
        
    except com_error as e:
        print(f"Erro ao responder: {e}")

def responder_todos_email(email_obj, corpo_resposta, html=False):
    """Responde todos os participantes da conversa."""
    try:
        resposta = email_obj.ReplyAll()
        
        if html:
            resposta.HTMLBody = corpo_resposta + "\n" + resposta.HTMLBody
        else:
            resposta.Body = corpo_resposta + "\n" + resposta.Body
        
        resposta.Send()
        
    except com_error as e:
        print(f"Erro: {e}")

def encaminhar_email(email_obj, para, corpo_adicional=""):
    """
    Encaminha um e-mail para outro(s) destinatário(s).
    """
    try:
        encaminhamento = email_obj.Forward()
        encaminhamento.To = para
        
        if corpo_adicional:
            encaminhamento.Body = corpo_adicional + "\n" + encaminhamento.Body
        
        encaminhamento.Send()
        print("E-mail encaminhado!")
        
    except com_error as e:
        print(f"Erro ao encaminhar: {e}")
```

#### 2.1.5 Salvar Como Rascunho

```python
def criar_email_rascunho(para, assunto, corpo, html=False):
    """
    Cria um e-mail como rascunho (não envia automaticamente).
    Retorna o objeto para possível edição posterior.
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        email = outlook.CreateItem(0)
        
        email.To = para
        email.Subject = assunto
        
        if html:
            email.HTMLBody = corpo
        else:
            email.Body = corpo
        
        # Salvar como rascunho ao invés de enviar
        email.Save()
        print("Rascunho salvo!")
        
        return email
        
    except com_error as e:
        print(f"Erro: {e}")
        return None
```

#### 2.1.6 Propriedades Avançadas de MailItem

```python
def configurar_email_avancado(para, assunto, corpo):
    """
    Demonstra propriedades avançadas de MailItem.
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        email = outlook.CreateItem(0)
        
        # Informações básicas
        email.To = para
        email.Subject = assunto
        email.Body = corpo
        
        # Prioridade (1=Baixa, 2=Normal, 3=Alta)
        email.Importance = 2
        
        # Nível de sensibilidade (0=Normal, 1=Pessoal, 2=Privado, 3=Confidencial)
        email.Sensitivity = 0
        
        # Data de entrega adiada
        email.DeferredDeliveryTime = datetime.now() + timedelta(hours=2)
        
        # Solicitar confirmação de leitura
        email.ReadReceiptRequested = True
        
        # Solicitar confirmação de entrega
        email.DeliveryReceiptRequested = True
        
        # Definir hora de expiração
        email.ExpiryTime = datetime.now() + timedelta(days=7)
        
        # Manter como rascunho ou enviar
        email.Save()  # Rascunho
        # email.Send()  # Enviar
        
        print("E-mail configurado com sucesso!")
        
    except com_error as e:
        print(f"Erro: {e}")
```

### 2.2 AppointmentItem (Calendário)

#### 2.2.1 Criar Compromisso Simples

```python
def criar_compromisso(assunto, data_inicio, duracao_minutos, 
                     local="", corpo=""):
    """
    Cria um compromisso no calendário do Outlook.
    
    Args:
        assunto (str): Título do compromisso
        data_inicio (datetime): Data e hora de início
        duracao_minutos (int): Duração em minutos
        local (str): Local (opcional)
        corpo (str): Descrição (opcional)
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        appointment = outlook.CreateItem(1)  # 1 = AppointmentItem
        
        appointment.Subject = assunto
        appointment.Start = data_inicio
        appointment.Duration = duracao_minutos
        appointment.Location = local
        appointment.Body = corpo
        
        appointment.Save()
        print("Compromisso criado!")
        
        return appointment
        
    except com_error as e:
        print(f"Erro ao criar compromisso: {e}")
        return None

# Uso
agora = datetime.now()
criar_compromisso(
    assunto="Reunião com Cliente",
    data_inicio=agora + timedelta(days=1),
    duracao_minutos=60,
    local="Sala 101",
    corpo="Discussão sobre novo projeto"
)
```

#### 2.2.2 Criar Reunião com Participantes

```python
def criar_reuniao(assunto, data_inicio, duracao_minutos, 
                 participantes, local="", obrigatorio=True):
    """
    Cria uma reunião com lista de participantes.
    
    Args:
        assunto (str): Assunto da reunião
        data_inicio (datetime): Quando inicia
        duracao_minutos (int): Duração
        participantes (list): Lista de e-mails dos participantes
        local (str): Local
        obrigatorio (bool): Participantes são obrigatórios?
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        reuniao = outlook.CreateItem(1)
        
        reuniao.Subject = assunto
        reuniao.Start = data_inicio
        reuniao.Duration = duracao_minutos
        reuniao.Location = local
        
        # Adicionar participantes
        recipients = reuniao.Recipients
        for email in participantes:
            recipient = recipients.Add(email)
            # Type: 1=Obrigatório, 2=Opcional, 3=Recurso
            recipient.Type = 1 if obrigatorio else 2
        
        recipients.ResolveAll()
        
        # Marca como reunião
        reuniao.MeetingStatus = 1  # 1 = olMeeting
        
        reuniao.Send()
        print("Reunião criada e convites enviados!")
        
    except com_error as e:
        print(f"Erro ao criar reunião: {e}")

# Uso
criar_reuniao(
    assunto="Planejamento Q1 2025",
    data_inicio=datetime(2025, 1, 15, 14, 0),
    duracao_minutos=90,
    participantes=["maria@example.com", "joao@example.com"],
    local="Sala de Conferência",
    obrigatorio=True
)
```

#### 2.2.3 Compromissos Recorrentes

```python
def criar_compromisso_recorrente(assunto, data_inicio, duracao_minutos,
                                 padrao_recorrencia, intervalo, 
                                 data_fim=None):
    """
    Cria um compromisso que se repete.
    
    Args:
        assunto (str): Título
        data_inicio (datetime): Primeira ocorrência
        duracao_minutos (int): Duração
        padrao_recorrencia (int): 0=Diário, 1=Semanal, 2=Mensal, 3=Anual
        intervalo (int): A cada N dias/semanas/meses
        data_fim (datetime): Quando termina a recorrência
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        appointment = outlook.CreateItem(1)
        
        appointment.Subject = assunto
        appointment.Start = data_inicio
        appointment.Duration = duracao_minutos
        
        # Configurar recorrência
        recurrence = appointment.GetRecurrencePattern()
        
        # RecurrenceType: 0=Diário, 1=Semanal, 2=Mensal, 3=Anual
        recurrence.RecurrenceType = padrao_recorrencia
        recurrence.Interval = intervalo
        
        if data_fim:
            recurrence.PatternEndDate = data_fim
        
        appointment.Save()
        print("Compromisso recorrente criado!")
        
    except com_error as e:
        print(f"Erro: {e}")

# Uso - Reunião semanal toda terça por 3 meses
criar_compromisso_recorrente(
    assunto="Reunião Semanal",
    data_inicio=datetime(2025, 1, 7, 10, 0),  # Próxima terça
    duracao_minutos=30,
    padrao_recorrencia=1,  # Semanal
    intervalo=1,  # Toda semana
    data_fim=datetime(2025, 4, 7)
)
```

#### 2.2.4 Ler Propriedades de Compromissos

```python
def listar_compromissos_dia(data_alvo):
    """
    Lista todos os compromissos de um determinado dia.
    
    Args:
        data_alvo (datetime): Data para consulta
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        ns = outlook.GetNamespace("MAPI")
        
        # Obter pasta de Calendário
        calendario = ns.GetDefaultFolder(9)  # 9 = Calendário
        
        # Filtrar por data
        data_inicio = data_alvo.replace(hour=0, minute=0, second=0, microsecond=0)
        data_fim = data_inicio + timedelta(days=1)
        
        compromissos = []
        for item in calendario.Items:
            if hasattr(item, 'Start'):
                if data_inicio <= item.Start < data_fim:
                    compromissos.append({
                        'Assunto': item.Subject,
                        'Início': item.Start,
                        'Duração': item.Duration,
                        'Local': item.Location,
                        'Corpo': item.Body
                    })
        
        for comp in compromissos:
            print(f"- {comp['Assunto']} às {comp['Início'].strftime('%H:%M')}")
        
        return compromissos
        
    except com_error as e:
        print(f"Erro: {e}")
        return []

# Uso
listar_compromissos_dia(datetime.now())
```

### 2.3 ContactItem (Contatos)

#### 2.3.1 Criar Novo Contato

```python
def criar_contato(nome_completo, email_pessoal="", email_trabalho="",
                 telefone="", endereco="", empresa="", cargo=""):
    """
    Cria um novo contato no Outlook.
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        contact = outlook.CreateItem(2)  # 2 = ContactItem
        
        # Nome
        contact.FullName = nome_completo
        
        # Emails
        if email_pessoal:
            contact.Email1Address = email_pessoal
            contact.Email1AddressType = "SMTP"
        
        if email_trabalho:
            contact.Email2Address = email_trabalho
            contact.Email2AddressType = "SMTP"
        
        # Telefone
        if telefone:
            contact.MobileTelephoneNumber = telefone
        
        # Informações de trabalho
        if empresa:
            contact.CompanyName = empresa
        if cargo:
            contact.JobTitle = cargo
        
        # Endereço
        if endereco:
            contact.MailingAddress = endereco
        
        contact.Save()
        print(f"Contato '{nome_completo}' criado!")
        
        return contact
        
    except com_error as e:
        print(f"Erro ao criar contato: {e}")
        return None

# Uso
criar_contato(
    nome_completo="Carlos Silva",
    email_trabalho="carlos.silva@company.com",
    telefone="(21) 98765-4321",
    empresa="Tech Solutions",
    cargo="Gerente de Projetos"
)
```

#### 2.3.2 Atualizar Contato Existente

```python
def atualizar_contato(nome_contato, **campos):
    """
    Atualiza campos de um contato existente.
    
    Args:
        nome_contato (str): Nome do contato a encontrar
        **campos: Campos a atualizar (email, telefone, etc)
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        ns = outlook.GetNamespace("MAPI")
        
        # Obter pasta de contatos
        contatos_pasta = ns.GetDefaultFolder(10)  # 10 = Contatos
        
        # Procurar contato por nome
        for item in contatos_pasta.Items:
            if hasattr(item, 'FullName') and item.FullName == nome_contato:
                # Atualizar campos
                for campo, valor in campos.items():
                    if hasattr(item, campo):
                        setattr(item, campo, valor)
                
                item.Save()
                print(f"Contato '{nome_contato}' atualizado!")
                return True
        
        print(f"Contato '{nome_contato}' não encontrado!")
        return False
        
    except com_error as e:
        print(f"Erro: {e}")
        return False

# Uso
atualizar_contato(
    "Carlos Silva",
    MobileTelephoneNumber="(21) 99999-9999",
    Email1Address="carlos.novo@company.com"
)
```

#### 2.3.3 Listar Todos os Contatos

```python
def listar_contatos(filtro_empresa=None):
    """
    Lista todos os contatos ou filtra por empresa.
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        ns = outlook.GetNamespace("MAPI")
        
        contatos_pasta = ns.GetDefaultFolder(10)  # 10 = Contatos
        
        contatos = []
        for item in contatos_pasta.Items:
            if hasattr(item, 'FullName'):
                # Filtrar se necessário
                if filtro_empresa:
                    if hasattr(item, 'CompanyName') and \
                       filtro_empresa.lower() in item.CompanyName.lower():
                        contatos.append(item)
                else:
                    contatos.append(item)
        
        for contato in contatos:
            print(f"\n{contato.FullName}")
            print(f"  Email: {getattr(contato, 'Email1Address', 'N/A')}")
            print(f"  Telefone: {getattr(contato, 'MobileTelephoneNumber', 'N/A')}")
            print(f"  Empresa: {getattr(contato, 'CompanyName', 'N/A')}")
        
        return contatos
        
    except com_error as e:
        print(f"Erro: {e}")
        return []

# Uso
listar_contatos()
listar_contatos(filtro_empresa="Tech Solutions")
```

### 2.4 TaskItem (Tarefas)

#### 2.4.1 Criar Nova Tarefa

```python
def criar_tarefa(assunto, data_vencimento=None, prioridade=1,
                 corpo="", percentual_concluido=0):
    """
    Cria uma nova tarefa no Outlook.
    
    Args:
        assunto (str): Título da tarefa
        data_vencimento (datetime): Quando a tarefa vence
        prioridade (int): 0=Baixa, 1=Normal, 2=Alta
        corpo (str): Descrição detalhada
        percentual_concluido (int): 0-100
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        tarefa = outlook.CreateItem(3)  # 3 = TaskItem
        
        tarefa.Subject = assunto
        tarefa.Body = corpo
        tarefa.Importance = prioridade  # 0=Baixa, 1=Normal, 2=Alta
        tarefa.PercentComplete = percentual_concluido / 100.0
        
        if data_vencimento:
            tarefa.DueDate = data_vencimento
        
        tarefa.Save()
        print(f"Tarefa '{assunto}' criada!")
        
        return tarefa
        
    except com_error as e:
        print(f"Erro ao criar tarefa: {e}")
        return None

# Uso
criar_tarefa(
    assunto="Revisar Proposta de Cliente",
    data_vencimento=datetime.now() + timedelta(days=3),
    prioridade=2,  # Alta
    corpo="Revisar detalhes da proposta e preparar feedback"
)
```

#### 2.4.2 Atualizar Progresso de Tarefa

```python
def atualizar_progresso_tarefa(nome_tarefa, percentual):
    """
    Atualiza o percentual de conclusão de uma tarefa.
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        ns = outlook.GetNamespace("MAPI")
        
        # Obter pasta de tarefas
        tarefas_pasta = ns.GetDefaultFolder(13)  # 13 = Tarefas
        
        for item in tarefas_pasta.Items:
            if hasattr(item, 'Subject') and item.Subject == nome_tarefa:
                item.PercentComplete = percentual / 100.0
                item.Save()
                print(f"Tarefa atualizada: {percentual}%")
                return True
        
        print(f"Tarefa '{nome_tarefa}' não encontrada!")
        return False
        
    except com_error as e:
        print(f"Erro: {e}")
        return False

def marcar_tarefa_concluida(nome_tarefa):
    """Marca uma tarefa como completamente concluída."""
    return atualizar_progresso_tarefa(nome_tarefa, 100)

# Uso
atualizar_progresso_tarefa("Revisar Proposta", 75)
marcar_tarefa_concluida("Revisar Proposta")
```

#### 2.4.3 Listar Tarefas Pendentes

```python
def listar_tarefas_pendentes(dias_futuros=30):
    """
    Lista tarefas pendentes com vencimento nos próximos N dias.
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        ns = outlook.GetNamespace("MAPI")
        
        tarefas_pasta = ns.GetDefaultFolder(13)
        
        hoje = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        limite = hoje + timedelta(days=dias_futuros)
        
        tarefas = []
        for item in tarefas_pasta.Items:
            if hasattr(item, 'Subject'):
                # Filtrar não concluídas
                if item.PercentComplete < 1.0:
                    # Verificar data de vencimento
                    if hasattr(item, 'DueDate') and item.DueDate:
                        if hoje <= item.DueDate <= limite:
                            tarefas.append(item)
        
        # Ordenar por data de vencimento
        tarefas.sort(key=lambda x: x.DueDate if hasattr(x, 'DueDate') 
                                    and x.DueDate else datetime.max)
        
        for tarefa in tarefas:
            vencimento = getattr(tarefa, 'DueDate', 'Sem data')
            progresso = int(getattr(tarefa, 'PercentComplete', 0) * 100)
            print(f"- [{progresso}%] {tarefa.Subject} (Vence: {vencimento})")
        
        return tarefas
        
    except com_error as e:
        print(f"Erro: {e}")
        return []

# Uso
listar_tarefas_pendentes(30)
```

---

## Organização e Navegação

### 3.1 Estrutura de Pastas

#### 3.1.1 Acessar Pastas Padrão

```python
def obter_pastas_padrao():
    """
    Retorna as pastas padrão do Outlook.
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        ns = outlook.GetNamespace("MAPI")
        
        pastas_padrao = {
            'Caixa de Entrada': ns.GetDefaultFolder(6),      # 6
            'Itens Enviados': ns.GetDefaultFolder(5),         # 5
            'Itens Deletados': ns.GetDefaultFolder(3),        # 3
            'Rascunhos': ns.GetDefaultFolder(16),             # 16
            'Spam': ns.GetDefaultFolder(23),                  # 23
            'Calendário': ns.GetDefaultFolder(9),             # 9
            'Contatos': ns.GetDefaultFolder(10),              # 10
            'Tarefas': ns.GetDefaultFolder(13),               # 13
            'Notas': ns.GetDefaultFolder(12),                 # 12
        }
        
        return pastas_padrao
        
    except com_error as e:
        print(f"Erro: {e}")
        return {}

# Uso
pastas = obter_pastas_padrao()
inbox = pastas['Caixa de Entrada']
```

#### 3.1.2 Acessar Pastas Personalizadas

```python
def encontrar_pasta(nome_pasta, pasta_pai=None):
    """
    Encontra uma pasta personalizada pelo nome.
    
    Args:
        nome_pasta (str): Nome da pasta a procurar
        pasta_pai: Pasta na qual buscar (None = raiz)
    
    Retorna: Objeto pasta ou None
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        ns = outlook.GetNamespace("MAPI")
        
        if pasta_pai is None:
            # Buscar nas pastas raiz
            pastas = ns.Folders
        else:
            # Buscar dentro de pasta específica
            pastas = pasta_pai.Folders
        
        for pasta in pastas:
            if pasta.Name.lower() == nome_pasta.lower():
                return pasta
        
        print(f"Pasta '{nome_pasta}' não encontrada!")
        return None
        
    except com_error as e:
        print(f"Erro: {e}")
        return None

def criar_pasta_personalizada(nome_pasta, pasta_pai=None):
    """
    Cria uma nova pasta personalizada.
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        ns = outlook.GetNamespace("MAPI")
        
        if pasta_pai is None:
            pasta_pai = ns.GetDefaultFolder(6)  # Caixa de Entrada
        
        # Verificar se já existe
        if encontrar_pasta(nome_pasta, pasta_pai):
            print(f"Pasta '{nome_pasta}' já existe!")
            return encontrar_pasta(nome_pasta, pasta_pai)
        
        nova_pasta = pasta_pai.Folders.Add(nome_pasta)
        nova_pasta.Save()
        print(f"Pasta '{nome_pasta}' criada!")
        
        return nova_pasta
        
    except com_error as e:
        print(f"Erro ao criar pasta: {e}")
        return None

# Uso
arquivos = criar_pasta_personalizada("Arquivos 2025")
subpasta = criar_pasta_personalizada("Clientes", arquivos)
```

#### 3.1.3 Listar Estrutura de Pastas

```python
def listar_estrutura_pastas(pasta=None, nivel=0):
    """
    Lista recursivamente a estrutura de pastas.
    """
    try:
        if pasta is None:
            outlook = win32com.client.Dispatch('Outlook.Application')
            ns = outlook.GetNamespace("MAPI")
            pastas = ns.Folders
        else:
            pastas = pasta.Folders
        
        for pasta_item in pastas:
            print("  " * nivel + f"├─ {pasta_item.Name}")
            
            # Recursivamente listar subpastas
            if pasta_item.Folders.Count > 0:
                listar_estrutura_pastas(pasta_item, nivel + 1)
    
    except com_error as e:
        print(f"Erro: {e}")

# Uso
print("Estrutura de pastas do Outlook:")
listar_estrutura_pastas()
```

### 3.2 Iteração e Acesso a Itens

#### 3.2.1 Listar E-mails de uma Pasta

```python
def listar_emails_pasta(pasta, limite=None):
    """
    Lista e-mails de uma pasta específica.
    
    Args:
        pasta: Objeto pasta do Outlook
        limite (int): Máximo de e-mails a retornar
    
    Retorna: Lista de e-mails
    """
    try:
        emails = []
        contador = 0
        
        for item in pasta.Items:
            # Verificar se é um MailItem
            if hasattr(item, 'Subject'):
                emails.append({
                    'De': getattr(item, 'SenderName', 'Desconhecido'),
                    'Assunto': item.Subject,
                    'Data': getattr(item, 'ReceivedTime', 'N/A'),
                    'Corpo': item.Body[:100] if item.Body else "",
                    'Objeto': item
                })
                
                contador += 1
                if limite and contador >= limite:
                    break
        
        return emails
        
    except com_error as e:
        print(f"Erro ao listar e-mails: {e}")
        return []

# Uso
outlook = win32com.client.Dispatch('Outlook.Application')
ns = outlook.GetNamespace("MAPI")
inbox = ns.GetDefaultFolder(6)

emails = listar_emails_pasta(inbox, limite=10)
for email in emails:
    print(f"De: {email['De']}")
    print(f"Assunto: {email['Assunto']}")
    print(f"Data: {email['Data']}\n")
```

#### 3.2.2 Contar Itens não Lidos

```python
def contar_nao_lidos(pasta):
    """Conta e-mails não lidos em uma pasta."""
    try:
        nao_lidos = 0
        for item in pasta.Items:
            if hasattr(item, 'UnRead') and item.UnRead:
                nao_lidos += 1
        return nao_lidos
    except com_error as e:
        print(f"Erro: {e}")
        return 0

def obter_emails_nao_lidos(pasta):
    """Retorna apenas e-mails não lidos."""
    try:
        nao_lidos = []
        for item in pasta.Items:
            if hasattr(item, 'UnRead') and hasattr(item, 'Subject') and item.UnRead:
                nao_lidos.append(item)
        return nao_lidos
    except com_error as e:
        print(f"Erro: {e}")
        return []

# Uso
outlook = win32com.client.Dispatch('Outlook.Application')
ns = outlook.GetNamespace("MAPI")
inbox = ns.GetDefaultFolder(6)

print(f"E-mails não lidos: {contar_nao_lidos(inbox)}")

emails_nao_lidos = obter_emails_nao_lidos(inbox)
for email in emails_nao_lidos:
    print(f"- {email.Subject}")
```

### 3.3 Pesquisa e Filtragem com DASL

#### 3.3.1 Fundamentos de DASL

DASL (DAV Searching and Locating) é a linguagem de filtros do Outlook. A sintaxe básica:

```
[Namespace].Items.Restrict("[SchemaName] Operator Value")
```

**Operadores comuns:**
- `=` : Igual
- `<>` : Diferente
- `<` : Menor que
- `>` : Maior que
- `<=` : Menor ou igual
- `>=` : Maior ou igual
- `Like` : Contém (texto com wildcard)

#### 3.3.2 Filtrar E-mails por Remetente

```python
def filtrar_por_remetente(pasta, email_remetente):
    """
    Encontra todos os e-mails de um remetente específico usando DASL.
    """
    try:
        # DASL é case-sensitive, use lowercase
        filtro = f"[SenderEmailAddress] = '{email_remetente}'"
        
        emails = pasta.Items.Restrict(filtro)
        emails.Sort("[ReceivedTime]", True)  # Ordenar por data
        
        resultado = []
        for email in emails:
            resultado.append({
                'De': email.SenderName,
                'Assunto': email.Subject,
                'Data': email.ReceivedTime
            })
        
        return resultado
        
    except com_error as e:
        print(f"Erro na pesquisa DASL: {e}")
        return []

# Uso
outlook = win32com.client.Dispatch('Outlook.Application')
ns = outlook.GetNamespace("MAPI")
inbox = ns.GetDefaultFolder(6)

emails = filtrar_por_remetente(inbox, "cliente@example.com")
for email in emails:
    print(f"{email['Data']}: {email['Assunto']}")
```

#### 3.3.3 Filtrar E-mails por Data

```python
def filtrar_por_data(pasta, data_inicio, data_fim):
    """
    Filtra e-mails dentro de um intervalo de datas.
    
    Args:
        pasta: Pasta para filtrar
        data_inicio (datetime): Data inicial
        data_fim (datetime): Data final
    """
    try:
        # Converter para formato UTC (formato COM)
        inicio_str = data_inicio.strftime("%m/%d/%Y")
        fim_str = data_fim.strftime("%m/%d/%Y")
        
        filtro = (f"[ReceivedTime] >= '{inicio_str}' AND "
                 f"[ReceivedTime] < '{fim_str}'")
        
        emails = pasta.Items.Restrict(filtro)
        emails.Sort("[ReceivedTime]", True)
        
        return list(emails)
        
    except com_error as e:
        print(f"Erro: {e}")
        return []

# Uso
inicio = datetime(2024, 12, 1)
fim = datetime(2024, 12, 31)

emails = filtrar_por_data(inbox, inicio, fim)
print(f"Encontrados {len(emails)} e-mails em dezembro")
```

#### 3.3.4 Filtrar por Palavras-chave no Assunto

```python
def filtrar_por_assunto(pasta, palavra_chave):
    """
    Filtra e-mails que contêm uma palavra no assunto.
    Nota: DASL não suporta Like diretamente em alguns casos.
    """
    try:
        # Usar método alternativo com Find
        emails = []
        
        for item in pasta.Items:
            if hasattr(item, 'Subject'):
                if palavra_chave.lower() in item.Subject.lower():
                    emails.append(item)
        
        return emails
        
    except com_error as e:
        print(f"Erro: {e}")
        return []

# Uso
emails_relatorio = filtrar_por_assunto(inbox, "Relatório")
print(f"Encontrados {len(emails_relatorio)} e-mails com 'Relatório' no assunto")
```

#### 3.3.5 Filtros Combinados Avançados

```python
def filtro_avancado_emailes(pasta, criterios):
    """
    Aplica múltiplos critérios de filtro.
    
    Args:
        pasta: Pasta para filtrar
        criterios (dict): Dicionário com critérios
                         {'remetente': '...', 'desde': datetime, 
                          'ate': datetime, 'nao_lido': True}
    """
    try:
        filtros = []
        
        if 'remetente' in criterios:
            filtros.append(f"[SenderEmailAddress] = '{criterios['remetente']}'")
        
        if 'desde' in criterios:
            data_str = criterios['desde'].strftime("%m/%d/%Y")
            filtros.append(f"[ReceivedTime] >= '{data_str}'")
        
        if 'ate' in criterios:
            data_str = criterios['ate'].strftime("%m/%d/%Y")
            filtros.append(f"[ReceivedTime] < '{data_str}'")
        
        if criterios.get('nao_lido'):
            filtros.append("[UnRead] = True")
        
        # Combinar filtros com AND
        filtro_completo = " AND ".join(filtros)
        
        if filtro_completo:
            items = pasta.Items.Restrict(filtro_completo)
        else:
            items = pasta.Items
        
        return list(items)
        
    except com_error as e:
        print(f"Erro: {e}")
        return []

# Uso
criterios = {
    'remetente': 'gerente@example.com',
    'desde': datetime(2024, 12, 1),
    'ate': datetime(2024, 12, 31),
    'nao_lido': True
}

emails = filtro_avancado_emailes(inbox, criterios)
print(f"Encontrados {len(emails)} e-mails com critérios combinados")
```

### 3.4 Operações em Lote

#### 3.4.1 Mover Múltiplos E-mails

```python
def mover_emails_para_pasta(emails, pasta_destino):
    """
    Move uma lista de e-mails para uma pasta específica.
    
    Args:
        emails (list): Lista de objetos MailItem
        pasta_destino: Objeto pasta de destino
    """
    try:
        movidos = 0
        for email in emails:
            try:
                email.Move(pasta_destino)
                movidos += 1
            except com_error:
                continue
        
        print(f"{movidos} e-mail(s) movido(s) com sucesso!")
        return movidos
        
    except com_error as e:
        print(f"Erro: {e}")
        return 0

# Uso
outlook = win32com.client.Dispatch('Outlook.Application')
ns = outlook.GetNamespace("MAPI")
inbox = ns.GetDefaultFolder(6)
arquivos = encontrar_pasta("Arquivos 2024")

emails = filtrar_por_data(inbox, datetime(2024, 1, 1), datetime(2024, 12, 31))
mover_emails_para_pasta(emails, arquivos)
```

#### 3.4.2 Deletar Múltiplos E-mails

```python
def deletar_emails(emails, permanente=False):
    """
    Deleta uma lista de e-mails.
    
    Args:
        emails (list): Lista de objetos MailItem
        permanente (bool): Se True, exclui permanentemente; 
                          se False, move para Lixeira
    """
    try:
        deletados = 0
        for email in emails:
            try:
                if permanente:
                    email.Delete()
                else:
                    # Mover para lixeira
                    outlook = win32com.client.Dispatch('Outlook.Application')
                    ns = outlook.GetNamespace("MAPI")
                    lixeira = ns.GetDefaultFolder(3)
                    email.Move(lixeira)
                deletados += 1
            except com_error:
                continue
        
        print(f"{deletados} e-mail(s) deletado(s)!")
        return deletados
        
    except com_error as e:
        print(f"Erro: {e}")
        return 0

# Uso com cuidado!
emails_spam = filtrar_por_remetente(inbox, "spammer@example.com")
deletar_emails(emails_spam, permanente=False)  # Move para lixeira
```

#### 3.4.3 Marcar Como Lido/Não Lido em Lote

```python
def marcar_como_lido(emails, lido=True):
    """
    Marca um conjunto de e-mails como lido ou não lido.
    
    Args:
        emails (list): Lista de MailItem
        lido (bool): True = marcar como lido, False = não lido
    """
    try:
        atualizados = 0
        for email in emails:
            try:
                if hasattr(email, 'UnRead'):
                    email.UnRead = not lido
                    email.Save()
                    atualizados += 1
            except com_error:
                continue
        
        status = "lido" if lido else "não lido"
        print(f"{atualizados} e-mail(s) marcado(s) como {status}!")
        return atualizados
        
    except com_error as e:
        print(f"Erro: {e}")
        return 0

# Uso
emails = obter_emails_nao_lidos(inbox)
marcar_como_lido(emails, lido=True)
```

---

## Funcionalidades Avançadas

### 4.1 Event Handling (Manipulação de Eventos)

#### 4.1.1 Monitorar Novos E-mails

```python
import win32com.client
from win32com.client import getevents
import pythoncom

class OutlookEventHandler:
    """
    Classe para lidar com eventos do Outlook como novos e-mails.
    """
    
    def __init__(self):
        self.outlook = win32com.client.Dispatch('Outlook.Application')
        self.namespace = self.outlook.GetNamespace("MAPI")
        self.inbox = self.namespace.GetDefaultFolder(6)
        
        # Configurar sink de eventos
        self.outlook_sink = win32com.client.WithEvents(
            self.outlook, getevents('Outlook.Application')
        )
    
    def OnNewMailEx(self, received_item_ids):
        """
        Disparado quando um novo e-mail é recebido.
        received_item_ids: String com IDs de e-mails recebidos (separado por vírgula)
        """
        print(f"Novo(s) e-mail(s) recebido(s)!")
        
        try:
            # Processar cada novo e-mail
            item_ids = received_item_ids.split(',')
            
            for item_id in item_ids:
                item_id = item_id.strip()
                try:
                    # Obter item pelo ID
                    item = self.namespace.GetItemFromID(item_id)
                    
                    if hasattr(item, 'Subject'):
                        print(f"Assunto: {item.Subject}")
                        print(f"De: {item.SenderName}")
                        print(f"Data: {item.ReceivedTime}")
                        
                        # Aqui você pode executar ações específicas
                        self.processar_novo_email(item)
                
                except com_error as e:
                    print(f"Erro ao processar e-mail: {e}")
        
        except Exception as e:
            print(f"Erro em OnNewMailEx: {e}")
    
    def processar_novo_email(self, email):
        """
        Função para processar o novo e-mail.
        Sobrescreva conforme necessário.
        """
        # Exemplo: salvar anexos automaticamente
        for attachment in email.Attachments:
            caminho = f"C:\\Attachments\\{attachment.FileName}"
            attachment.SaveAsFile(caminho)
            print(f"Anexo salvo: {caminho}")

# Uso
def monitorar_outlook():
    """Executa o monitor de eventos do Outlook."""
    pythoncom.CoInitialize()
    
    try:
        handler = OutlookEventHandler()
        print("Monitorando Outlook para novos e-mails...")
        print("(Pressione Ctrl+C para sair)")
        
        # Loop de pump de mensagens
        while True:
            pythoncom.PumpWaitingMessages()
    
    except KeyboardInterrupt:
        print("\nMonitor encerrado.")
    finally:
        pythoncom.CoUninitialize()

# Executar em thread separada (recomendado)
import threading
thread = threading.Thread(target=monitorar_outlook, daemon=True)
thread.start()
```

#### 4.1.2 Backup Automático de Anexos

```python
def fazer_backup_anexos_automatico(pasta_backup, filtro_remetente=None):
    """
    Salva automaticamente todos os anexos recebidos.
    
    Args:
        pasta_backup (str): Caminho para salvar anexos
        filtro_remetente (str): Opcional - apenas de remetente específico
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        ns = outlook.GetNamespace("MAPI")
        inbox = ns.GetDefaultFolder(6)
        
        os.makedirs(pasta_backup, exist_ok=True)
        
        emails_com_anexos = []
        
        for email in inbox.Items:
            if hasattr(email, 'Attachments') and email.Attachments.Count > 0:
                # Filtrar por remetente se especificado
                if filtro_remetente:
                    if filtro_remetente.lower() not in email.SenderEmailAddress.lower():
                        continue
                
                emails_com_anexos.append(email)
        
        total_salvos = 0
        for email in emails_com_anexos:
            for attachment in email.Attachments:
                try:
                    caminho_arquivo = os.path.join(
                        pasta_backup,
                        f"{email.ReceivedTime.strftime('%Y%m%d_%H%M%S')}_{attachment.FileName}"
                    )
                    
                    attachment.SaveAsFile(caminho_arquivo)
                    total_salvos += 1
                    print(f"Backup: {attachment.FileName}")
                
                except com_error as e:
                    print(f"Erro ao salvar anexo: {e}")
        
        print(f"Total de anexos salvos: {total_salvos}")
        
    except com_error as e:
        print(f"Erro: {e}")

# Uso
fazer_backup_anexos_automatico(r"C:\Backups\Anexos")
fazer_backup_anexos_automatico(r"C:\Backups\ClienteX", "clientex@example.com")
```

### 4.2 Manipulação de Regras

#### 4.2.1 Listar Regras de Caixa de Entrada

```python
def listar_regras():
    """
    Lista todas as regras configuradas no Outlook.
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        ns = outlook.GetNamespace("MAPI")
        
        regras = ns.GetRules()
        
        print(f"Total de regras: {regras.Count}\n")
        
        for regra in regras:
            print(f"Nome: {regra.Name}")
            print(f"Ativa: {regra.Enabled}")
            print(f"Executar em: {'Recebimento' if regra.ExecuteOnReceive else 'Manual'}")
            print("-" * 40)
        
        return regras
        
    except com_error as e:
        print(f"Erro ao listar regras: {e}")
        return None

# Uso
listar_regras()
```

#### 4.2.2 Criar Regra Programaticamente

```python
def criar_regra(nome_regra, condicoes, acoes):
    """
    Cria uma nova regra de Outlook.
    
    Nota: Regras COM têm limitações. Considere usar VBA ou UI do Outlook
    para configurações mais complexas.
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        ns = outlook.GetNamespace("MAPI")
        
        regras = ns.GetRules()
        
        # Verificar se regra já existe
        for regra in regras:
            if regra.Name == nome_regra:
                print(f"Regra '{nome_regra}' já existe!")
                return False
        
        # Criar nova regra
        nova_regra = regras.Create(nome_regra, 0)  # 0 = RecipientRule
        nova_regra.Enabled = True
        nova_regra.ExecuteOnReceive = True
        
        # Nota: Adicionar condições e ações é complexo via COM
        # Recomenda-se usar a UI do Outlook para configurações reais
        
        # Salvar regra
        regras.Save()
        print(f"Regra '{nome_regra}' criada com sucesso!")
        
        return True
        
    except com_error as e:
        print(f"Erro ao criar regra: {e}")
        return False
```

### 4.3 Acesso a Dados do Usuário

#### 4.3.1 Obter Informações do Perfil

```python
def obter_info_perfil():
    """
    Obtém informações do perfil/conta do Outlook.
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        
        # Obter objeto Application
        app_info = {
            'Versão': outlook.Version,
            'Nome': outlook.Name,
        }
        
        # Obter informações da conta padrão
        ns = outlook.GetNamespace("MAPI")
        
        info = {
            'App': app_info,
            'Contas': []
        }
        
        # Iterar sobre contas
        for account in ns.Accounts:
            conta_info = {
                'Nome': account.DisplayName,
                'Email': account.EmailAddress,
                'SMTP': account.SmtpAddress,
                'Tipo': account.AccountType,
            }
            info['Contas'].append(conta_info)
        
        # Exibir informações
        print(f"Outlook: {info['App']['Nome']} (v{info['App']['Versão']})\n")
        
        for conta in info['Contas']:
            print(f"Conta: {conta['Nome']}")
            print(f"  Email: {conta['Email']}")
            print(f"  SMTP: {conta['SMTP']}")
            print()
        
        return info
        
    except com_error as e:
        print(f"Erro: {e}")
        return None

# Uso
obter_info_perfil()
```

#### 4.3.2 Obter Informações de Pasta

```python
def obter_info_pasta(pasta):
    """
    Retorna informações detalhadas sobre uma pasta.
    """
    try:
        info = {
            'Nome': pasta.Name,
            'Total de Itens': pasta.Items.Count,
            'Itens Não Lidos': 0,
            'Tipo': 'Desconhecido'
        }
        
        # Contar não lidos (apenas para pastas de e-mail)
        if hasattr(pasta, 'UnReadItemCount'):
            info['Itens Não Lidos'] = pasta.UnReadItemCount
        
        # Determinar tipo de pasta
        default_folder_index = -1
        try:
            for i in range(22):
                outlook = win32com.client.Dispatch('Outlook.Application')
                ns = outlook.GetNamespace("MAPI")
                if pasta == ns.GetDefaultFolder(i):
                    default_folder_index = i
                    break
        except:
            pass
        
        tipos = {
            3: 'Lixeira',
            5: 'Itens Enviados',
            6: 'Caixa de Entrada',
            9: 'Calendário',
            10: 'Contatos',
            12: 'Notas',
            13: 'Tarefas',
            16: 'Rascunhos',
            23: 'Spam'
        }
        
        if default_folder_index in tipos:
            info['Tipo'] = tipos[default_folder_index]
        
        return info
        
    except com_error as e:
        print(f"Erro: {e}")
        return None

# Uso
outlook = win32com.client.Dispatch('Outlook.Application')
ns = outlook.GetNamespace("MAPI")
inbox = ns.GetDefaultFolder(6)

info = obter_info_pasta(inbox)
for chave, valor in info.items():
    print(f"{chave}: {valor}")
```

---

## Melhores Práticas e Desempenho

### 5.1 Tratamento de Erros

#### 5.1.1 Tratamento Básico de Erros COM

```python
from pywintypes import com_error
import traceback

def operacao_com_outlook_segura():
    """
    Demonstra tratamento robusto de erros COM.
    """
    outlook = None
    
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        ns = outlook.GetNamespace("MAPI")
        inbox = ns.GetDefaultFolder(6)
        
        # Suas operações aqui
        print(f"Caixa de entrada tem {inbox.Items.Count} itens")
    
    except com_error as e:
        # Erro específico de COM
        print(f"Erro COM: {e.hresult}")
        print(f"Mensagem: {e.strerror}")
        
        if e.hresult == -2147221005:  # "Classe não registrada"
            print("Verifique se o Outlook está instalado corretamente")
        
    except AttributeError as e:
        print(f"Erro de Atributo: {e}")
        print("Propriedade ou método não encontrado")
    
    except Exception as e:
        print(f"Erro Geral: {e}")
        traceback.print_exc()
    
    finally:
        # Limpeza
        if outlook:
            del outlook

# Uso
operacao_com_outlook_segura()
```

#### 5.1.2 Validação de Objetos

```python
def validar_objeto_outlook(obj):
    """
    Valida se um objeto Outlook ainda está válido.
    """
    try:
        # Tentar acessar uma propriedade básica
        _ = obj.Name
        return True
    except:
        return False

def operacao_com_validacao(pasta):
    """
    Realiza operações com validação contínua.
    """
    try:
        if not validar_objeto_outlook(pasta):
            print("Objeto pasta inválido ou desconectado!")
            return
        
        # Processar itens
        for item in pasta.Items:
            if not validar_objeto_outlook(item):
                print("Item inválido, pulando...")
                continue
            
            # Processar item
            print(item.Subject if hasattr(item, 'Subject') else "Sem assunto")
    
    except com_error as e:
        print(f"Erro durante processamento: {e}")
```

### 5.2 Otimização de Desempenho

#### 5.2.1 Iteração Eficiente de Itens

```python
def iterar_eficiente(pasta, callback, limite_memoria=1000):
    """
    Itera sobre itens de forma eficiente em memória.
    
    Args:
        pasta: Pasta para iterar
        callback: Função a chamar para cada item
        limite_memoria: Máximo de itens em memória por vez
    """
    try:
        items = list(pasta.Items)  # Criar snapshot
        total = len(items)
        
        print(f"Processando {total} itens...")
        
        for i, item in enumerate(items):
            try:
                callback(item)
            except com_error:
                continue
            
            # Liberar memória periodicamente
            if (i + 1) % limite_memoria == 0:
                import gc
                gc.collect()
                print(f"Processados {i + 1}/{total}")
        
        print("Processamento concluído!")
        
    except com_error as e:
        print(f"Erro: {e}")

def processar_item(item):
    """Função de callback para processar cada item."""
    if hasattr(item, 'Subject'):
        print(f"  - {item.Subject}")

# Uso
outlook = win32com.client.Dispatch('Outlook.Application')
ns = outlook.GetNamespace("MAPI")
inbox = ns.GetDefaultFolder(6)

iterar_eficiente(inbox, processar_item)
```

#### 5.2.2 Cache de Pastas

```python
class CacheOutlook:
    """
    Implementa cache de objetos Outlook para evitar múltiplas 
    buscas.
    """
    
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.pastas_cache = {}
    
    def conectar(self):
        """Conecta ao Outlook."""
        try:
            self.outlook = win32com.client.Dispatch('Outlook.Application')
            self.namespace = self.outlook.GetNamespace("MAPI")
            return True
        except com_error as e:
            print(f"Erro ao conectar: {e}")
            return False
    
    def obter_pasta(self, nome_pasta):
        """Obtém pasta usando cache."""
        if nome_pasta in self.pastas_cache:
            return self.pastas_cache[nome_pasta]
        
        pasta = encontrar_pasta(nome_pasta)
        if pasta:
            self.pastas_cache[nome_pasta] = pasta
        
        return pasta
    
    def obter_inbox(self):
        """Obtém caixa de entrada em cache."""
        if 'Inbox' not in self.pastas_cache:
            self.pastas_cache['Inbox'] = self.namespace.GetDefaultFolder(6)
        return self.pastas_cache['Inbox']
    
    def limpar_cache(self):
        """Limpa o cache de pastas."""
        self.pastas_cache.clear()

# Uso
cache = CacheOutlook()
if cache.conectar():
    inbox = cache.obter_inbox()
    print(f"Itens na caixa: {inbox.Items.Count}")
    
    cache.limpar_cache()
```

### 5.3 Estabilidade e Limpeza

#### 5.3.1 Gerenciador de Contexto

```python
from contextlib import contextmanager

@contextmanager
def outlook_session():
    """
    Context manager para garantir limpeza adequada.
    """
    outlook = None
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        yield outlook
    except com_error as e:
        print(f"Erro na sessão Outlook: {e}")
        raise
    finally:
        if outlook:
            # Tentar fechar com segurança
            try:
                # Não fechar o Outlook se estava rodando antes
                # del outlook
                pass
            except:
                pass

# Uso
with outlook_session() as outlook:
    ns = outlook.GetNamespace("MAPI")
    inbox = ns.GetDefaultFolder(6)
    print(f"Caixa: {inbox.Items.Count}")
```

#### 5.3.2 Timeout de Operações

```python
import signal

def timeout_handler(signum, frame):
    raise TimeoutError("Operação excedeu o tempo limite")

def operacao_com_timeout(funcao, timeout_segundos=30):
    """
    Executa uma função com timeout.
    Nota: Funciona melhor em Unix/Linux
    """
    # Definir handler de signal
    signal.signal(signal.SIGALRM, timeout_handler)
    signal.alarm(timeout_segundos)
    
    try:
        resultado = funcao()
        signal.alarm(0)  # Cancelar alarme
        return resultado
    except TimeoutError:
        print(f"Operação excedeu {timeout_segundos}s!")
        return None

# Uso (cuidado em Windows)
def buscar_emails():
    outlook = win32com.client.Dispatch('Outlook.Application')
    return outlook.GetNamespace("MAPI").GetDefaultFolder(6).Items.Count

# resultado = operacao_com_timeout(buscar_emails, 10)
```

---

## Exemplos Práticos Completos

### 6.1 Processador de Relatórios Automático

```python
def processar_relatoriios_automaticamente(pasta_destino=None):
    """
    Automaticamente processa e-mails com "Relatório" no assunto:
    - Salva anexos
    - Move para pasta de arquivos
    - Registra em log
    """
    import json
    from datetime import datetime
    
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        ns = outlook.GetNamespace("MAPI")
        inbox = ns.GetDefaultFolder(6)
        
        if pasta_destino is None:
            pasta_destino = criar_pasta_personalizada("Relatórios Processados")
        
        # Criar diretório para anexos
        dir_anexos = r"C:\Relatorios\Anexos"
        os.makedirs(dir_anexos, exist_ok=True)
        
        # Log de processamento
        log = []
        
        # Procurar e-mails com "Relatório" no assunto
        emails = filtrar_por_assunto(inbox, "Relatório")
        
        print(f"Processando {len(emails)} relatório(s)...")
        
        for email in emails:
            try:
                registro = {
                    'timestamp': datetime.now().isoformat(),
                    'assunto': email.Subject,
                    'de': email.SenderName,
                    'anexos': []
                }
                
                # Processar anexos
                for attachment in email.Attachments:
                    try:
                        nome_arquivo = (
                            f"{email.ReceivedTime.strftime('%Y%m%d_%H%M%S')}_"
                            f"{attachment.FileName}"
                        )
                        caminho_completo = os.path.join(dir_anexos, nome_arquivo)
                        
                        attachment.SaveAsFile(caminho_completo)
                        
                        registro['anexos'].append({
                            'nome_original': attachment.FileName,
                            'nome_salvo': nome_arquivo,
                            'caminho': caminho_completo
                        })
                        
                    except com_error as e:
                        registro['anexos'].append({
                            'nome': attachment.FileName,
                            'erro': str(e)
                        })
                
                # Mover e-mail
                email.Move(pasta_destino)
                registro['status'] = 'processado'
                
                log.append(registro)
                print(f"✓ {email.Subject}")
            
            except com_error as e:
                print(f"✗ Erro: {e}")
        
        # Salvar log
        log_path = os.path.join(dir_anexos, "log_processamento.json")
        with open(log_path, 'w', encoding='utf-8') as f:
            json.dump(log, f, indent=2, ensure_ascii=False, default=str)
        
        print(f"\nLog salvo: {log_path}")
        return len(emails)
        
    except com_error as e:
        print(f"Erro: {e}")
        return 0

# Uso
processar_relatoriios_automaticamente()
```

### 6.2 Agendador de Reuniões em Lote

```python
def agendar_reunioes_em_lote(dados_reunioes):
    """
    Agenda múltiplas reuniões a partir de dados estruturados.
    
    Args:
        dados_reunioes (list): Lista de dicts com:
            {
                'assunto': str,
                'data': '2025-01-15',
                'hora': '14:00',
                'duracao': 60,
                'participantes': [emails],
                'local': str
            }
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        
        agendadas = 0
        erros = []
        
        for reuniao_data in dados_reunioes:
            try:
                # Parsear data e hora
                data_str = reuniao_data['data']
                hora_str = reuniao_data['hora']
                data_hora = datetime.strptime(
                    f"{data_str} {hora_str}", 
                    "%Y-%m-%d %H:%M"
                )
                
                # Criar reunião
                reuniao = outlook.CreateItem(1)
                reuniao.Subject = reuniao_data['assunto']
                reuniao.Start = data_hora
                reuniao.Duration = reuniao_data.get('duracao', 60)
                reuniao.Location = reuniao_data.get('local', '')
                reuniao.Body = reuniao_data.get('corpo', '')
                
                # Adicionar participantes
                recipients = reuniao.Recipients
                for email in reuniao_data.get('participantes', []):
                    recipient = recipients.Add(email)
                    recipient.Type = 1
                
                recipients.ResolveAll()
                reuniao.MeetingStatus = 1
                reuniao.Send()
                
                agendadas += 1
                print(f"✓ {reuniao_data['assunto']}")
            
            except Exception as e:
                erros.append((reuniao_data.get('assunto', 'Sem assunto'), str(e)))
                print(f"✗ Erro: {e}")
        
        print(f"\nResumo: {agendadas} agendada(s), {len(erros)} erro(s)")
        
        if erros:
            print("\nDetalhes dos erros:")
            for assunto, erro in erros:
                print(f"  - {assunto}: {erro}")
        
        return agendadas
        
    except com_error as e:
        print(f"Erro geral: {e}")
        return 0

# Uso
reunioes = [
    {
        'assunto': 'Planejamento Q1',
        'data': '2025-01-15',
        'hora': '14:00',
        'duracao': 90,
        'participantes': ['maria@example.com', 'joao@example.com'],
        'local': 'Sala A'
    },
    {
        'assunto': 'Follow-up com Cliente',
        'data': '2025-01-16',
        'hora': '10:00',
        'duracao': 30,
        'participantes': ['cliente@example.com'],
        'local': 'Virtual'
    }
]

agendar_reunioes_em_lote(reunioes)
```

### 6.3 Sincronizador de Contatos

```python
def sincronizar_contatos_csv(caminho_csv, atualizar_existentes=True):
    """
    Importa/sincroniza contatos a partir de arquivo CSV.
    
    Formato CSV:
    Nome,Email,Telefone,Empresa,Cargo
    """
    import csv
    
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        ns = outlook.GetNamespace("MAPI")
        contatos_pasta = ns.GetDefaultFolder(10)
        
        contatos_existentes = {}
        for item in contatos_pasta.Items:
            if hasattr(item, 'FullName'):
                contatos_existentes[item.FullName.lower()] = item
        
        importados = 0
        atualizados = 0
        
        with open(caminho_csv, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            
            for row in reader:
                nome = row.get('Nome', '').strip()
                
                if not nome:
                    continue
                
                # Verificar se existe
                contato_chave = nome.lower()
                
                if contato_chave in contatos_existentes:
                    if atualizar_existentes:
                        contato = contatos_existentes[contato_chave]
                        contato.Email1Address = row.get('Email', '')
                        contato.MobileTelephoneNumber = row.get('Telefone', '')
                        contato.CompanyName = row.get('Empresa', '')
                        contato.JobTitle = row.get('Cargo', '')
                        contato.Save()
                        atualizados += 1
                        print(f"✓ Atualizado: {nome}")
                else:
                    # Criar novo
                    contato = outlook.CreateItem(2)
                    contato.FullName = nome
                    contato.Email1Address = row.get('Email', '')
                    contato.MobileTelephoneNumber = row.get('Telefone', '')
                    contato.CompanyName = row.get('Empresa', '')
                    contato.JobTitle = row.get('Cargo', '')
                    contato.Save()
                    importados += 1
                    print(f"✓ Importado: {nome}")
        
        print(f"\nResumo: {importados} novo(s), {atualizados} atualizado(s)")
        return importados + atualizados
        
    except Exception as e:
        print(f"Erro: {e}")
        return 0

# Uso
sincronizar_contatos_csv(r"C:\contatos.csv")
```

---

## Referência Rápida de Enumerações

### Tipo de Itens
```
MailItem = 0
AppointmentItem = 1
ContactItem = 2
TaskItem = 3
JournalItem = 4
NoteItem = 5
```

### Pastas Padrão
```
3 = Lixeira
5 = Itens Enviados
6 = Caixa de Entrada
9 = Calendário
10 = Contatos
12 = Notas
13 = Tarefas
16 = Rascunhos
23 = Spam/Lixo
```

### Prioridade de E-mail
```
1 = Baixa
2 = Normal
3 = Alta
```

### Status de Importância
```
0 = Normal
1 = Pessoal
2 = Privado
3 = Confidencial
```

### Tipo de Recipient
```
1 = Obrigatório (To)
2 = Opcional (CC)
3 = Cópia Oculta (BCC)
```

### Tipo de Recorrência
```
0 = Diário
1 = Semanal
2 = Mensal
3 = Anual
```

---

## Conclusão

Esta documentação cobre os aspectos essenciais e avançados da automação do Outlook com pywin32. Para mais informações:

- **Documentação COM**: https://docs.microsoft.com/de-de/office/vba/api/outlook.application
- **Repositórios GitHub**: Procure por "pywin32 outlook"
- **Stack Overflow**: Tag `pywin32` e `outlook`

Lembre-se sempre de:
1. Tratar erros COM adequadamente
2. Limpar recursos após uso
3. Testar em ambiente controlado
4. Fazer backup antes de deletar dados
5. Usar DASL para filtros complexos
6. Implementar logging para produção
