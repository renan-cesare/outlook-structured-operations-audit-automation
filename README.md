Outlook Audit Automation

Projeto profissional sanitizado de automação de auditoria operacional, utilizado em contexto real de backoffice / risco & compliance.

Contexto

Em ambientes corporativos do mercado financeiro, diversas operações precisam passar por processos formais de auditoria interna, incluindo:

Contato com assessores responsáveis

Coleta de justificativas formais

Registro de evidências

Acompanhamento de respostas

Reenvio de cobranças quando não há retorno

Este projeto automatiza todo esse ciclo de forma integrada ao Microsoft Outlook e planilhas Excel.

Este repositório contém uma versão sanitizada:

Sem nomes reais

Sem e-mails reais

Sem dados internos

Sem estruturas proprietárias

Mas preserva integralmente a lógica real do processo.

O que o sistema faz
1) Módulo de Envio (dispatch)

Lê planilha Excel de operações a serem auditadas

Lê base de dados de profissionais (assessores / líderes)

Gera e envia e-mails automaticamente via Outlook

Insere um token único no corpo do e-mail para rastreio

Localiza o e-mail enviado na pasta “Itens Enviados”

Captura e salva:

ConversationID

InternetMessageID

EntryID

Registra tudo em uma planilha de histórico

2) Módulo de Acompanhamento (followup)

Lê a planilha de histórico

Para cada envio:

Localiza o e-mail original pelo EntryID

Busca respostas na caixa de entrada via ConversationID

Se encontrou resposta:

Marca como Respondido

Salva data e conteúdo da resposta

Se não encontrou:

Gera automaticamente uma cobrança (reply)

Atualiza o status no histórico

Estrutura do Projeto
outlook-audit-automation/
│
├── main.py
├── config.example.json
├── requirements.txt
├── README.md
├── LICENSE
│
└── src/
    └── outlook_audit/
        ├── __init__.py
        ├── config.py
        ├── dispatch.py
        ├── followup.py
        ├── history_store.py
        ├── outlook_client.py
        ├── logging_utils.py
        └── file_lock.py

Configuração

Clone o repositório

Crie um arquivo:

config.json


Baseando-se em:

config.example.json


Ajuste os caminhos das planilhas e parâmetros.

Como rodar

Instalar dependências:

pip install -r requirements.txt

Teste seguro (não envia e-mail):
python main.py dispatch --dry-run

Mostrar e-mails antes de enviar:
python main.py dispatch --display-only

Rodar acompanhamento:
python main.py followup --display-only

Segurança

O projeto:

Bloqueia planilhas abertas em uso

Nunca sobrescreve histórico manualmente

Usa tokens únicos por envio

O .gitignore impede subir:

config.json real

planilhas reais

logs

O que este projeto demonstra tecnicamente

Automação corporativa real

Integração com Outlook via COM

Controle de estado e histórico

Idempotência e rastreabilidade

Arquitetura modular

Separação de responsabilidades

Processamento de Excel com pandas

Padrões de projeto aplicados a backoffice / compliance

Observação importante

Este projeto não é um script de estudo.
Ele é a formalização sanitizada de uma automação real de produção usada em ambiente corporativo.

Licença

MIT License.
