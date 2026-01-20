# outlook-audit-automation

Automação **(sanitizada)** para rotina de auditoria operacional/compliance:

* **Envio automático** de e-mails via Microsoft Outlook (COM / pywin32)
* Captura de IDs (**ConversationID / InternetMessageID / EntryID**)
* Registro de histórico em **Excel**
* **Acompanhamento** de respostas e **cobrança automática** (Reply)

> **Aviso de sanitização (importante)**
> Este repositório é uma versão **profissional e sanitizada** de uma automação real usada em rotina operacional.
> Não há dados reais, e-mails reais, nomes reais, estruturas internas, nem credenciais.

---

## Problema real de negócio

Em auditorias internas, é comum precisar:

* enviar questionários padronizados para responsáveis (ex.: assessores),
* manter rastreabilidade (o que foi enviado, quando e para quem),
* capturar identificadores confiáveis do e-mail para localizar respostas,
* cobrar novamente quando não houver retorno.

Executar isso manualmente é repetitivo e sujeito a falhas. Esta automação padroniza, rastreia e registra o processo.

---

## O que o projeto faz

### 1) Dispatch (envio)

* Lê uma planilha Excel de **operações para auditar**
* Lê uma planilha Excel de **base de profissionais**
* Envia e-mail via **Outlook**
* Insere um **token único** no corpo do e-mail
* Busca no **Sent Items** e captura:

  * `ConversationID`
  * `InternetMessageID`
  * `EntryID`
* Registra tudo em uma planilha Excel de **histórico** (append seguro)

### 2) Follow-up (acompanhamento/cobrança)

* Lê o histórico (Excel)
* Filtra registros do mês e status “Enviado”
* Abre o e-mail original via `EntryID` e verifica resposta via `ConversationID` na Inbox
* Se houver resposta: registra conteúdo e data e marca como respondido
* Se não houver: cria cobrança via `Reply()` (com opção `.Display()` ou `.Send()`)

---

## Requisitos

* Windows
* Microsoft Outlook instalado e configurado
* Python 3.10+
* Acesso às planilhas de entrada/saída

---

## Configuração

### 1) Instalar dependências

```bash
pip install -r requirements.txt
```

### 2) Criar `config.json`

Copie:

* `config.example.json` → `config.json`

Edite o `config.json` com os caminhos reais do seu computador.

---

## Como usar

### Enviar e registrar (dispatch)

```bash
python main.py dispatch
```

Modo seguro (não envia):

```bash
python main.py dispatch --dry-run
```

Modo seguro (abre o e-mail no Outlook e não envia automaticamente):

```bash
python main.py dispatch --display-only
```

### Acompanhar e cobrar (followup)

```bash
python main.py followup
```

Sobrescrever mês de referência:

```bash
python main.py followup --month 2026-01
```

Modo seguro (abre a cobrança e não envia automaticamente):

```bash
python main.py followup --display-only
```

---

## Entradas esperadas (planilhas)

### A) Operações (`paths.operations_xlsx`)

Colunas obrigatórias (nomes esperados):

* `Código Cliente`
* `Nome do Cliente`
* `Estrutura`
* `Ativo`
* `% PL`
* `Assessor da Operação`
* `Assessor do Cliente`

### B) Base de profissionais (`paths.professionals_xlsx`)

Colunas obrigatórias:

* `Código Assessor`
* `Nome Completo`
* `E-mail`
* `Código do Líder`

---

## Saída (histórico Excel)

Arquivo: `paths.history_xlsx`
Aba: `paths.history_sheet`

Campos rastreados incluem:

* `Email Assessor`, `Email Lider`
* `Status`, `Data Envio`, `Assunto`, `Token Identificador`
* `ConversationID`, `InternetID`, `EntryID`
* Campos do follow-up: `Status Auditoria`, `Data da Resposta`, `Conteúdo da Resposta`, `Data da Nova Cobrança`

---

## Observações e limitações

* Automação baseada em Outlook COM: **Windows-only**
* Caixas muito grandes podem tornar a busca mais lenta; o follow-up ordena por `ReceivedTime` desc para acelerar

---

## Licença

MIT
