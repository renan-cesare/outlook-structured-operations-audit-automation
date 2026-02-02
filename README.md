# Outlook Structured Operations Audit Automation

Automação em Python para **auditoria operacional estruturada**, com **envio de e-mails via Outlook**, **captura de identificadores (IDs)**, **histórico em Excel** e **cobrança automática (follow-up)** baseada em status de resposta.

> **English (short):** Python automation for structured operations audit workflows using Outlook, with ID tracking, Excel history, and automated follow-up.

---

## Principais recursos

* Automação de envio de e-mails via **Outlook (COM automation)**
* Geração de mensagens a partir de **templates HTML**
* Captura e persistência de **IDs de mensagens**
* Registro estruturado de histórico em **Excel**
* **Follow-up automático** para casos sem resposta
* Controle de execução e rastreabilidade operacional
* Separação clara entre:

  * lógica de auditoria
  * templates de comunicação
  * histórico de controle

---

## Diferença em relação a outros projetos similares

Este projeto é focado em **auditoria operacional e cobrança estruturada**, enquanto outros fluxos de auditoria podem focar em **análise de desempenho**.

Aqui, o objetivo principal é:

* garantir retorno operacional
* registrar evidências de contato
* automatizar cobranças recorrentes
* manter histórico auditável de interações

---

## Contexto

Em rotinas de **operações, risco e compliance interno**, é comum a necessidade de:

* envio estruturado de solicitações
* identificação única de cada contato
* acompanhamento de pendências
* cobrança automática após prazos definidos

Este projeto automatiza esse processo, reduzindo esforço manual e aumentando **controle, padronização e rastreabilidade**.

---

## Aviso importante (uso autorizado)

Este repositório é apresentado como **exemplo técnico e portfólio**.

* Utilize apenas **ambientes e contas autorizadas**
* Não publique dados reais, e-mails corporativos ou informações sensíveis
* Respeite políticas internas, LGPD e regras de uso do Outlook

---

## Estrutura do projeto

```text
.
├─ src/
│  └─ outlook_audit/
│     ├─ __init__.py
│     ├─ app.py
│     └─ dispatch.py
├─ templates/
│  └─ email_body.html
├─ config.example.json
├─ main.py
├─ requirements.txt
├─ LICENSE
└─ README.md
```

---

## Requisitos

* Python 3.10+
* **Windows**
* Microsoft Outlook instalado e configurado

> Este projeto utiliza automação COM, sendo compatível apenas com ambiente Windows.

---

## Instalação

```bash
python -m venv .venv

# Windows
.venv\\Scripts\\activate

pip install -r requirements.txt
```

---

## Configuração

Crie um arquivo local de configuração:

```bash
copy config.example.json config.json
```

O arquivo de configuração define:

* parâmetros de envio
* caminhos de arquivos Excel
* prazos de cobrança
* regras de acompanhamento

> O arquivo `config.json` deve permanecer fora do versionamento.

---

## Execução

```bash
python main.py
```

O processo:

* envia os e-mails iniciais
* registra IDs e histórico
* identifica pendências
* executa cobranças automáticas conforme regras

---

## Saídas geradas

* Histórico estruturado em Excel
* Controle de pendências
* Evidências de auditoria operacional

---

## Sanitização de dados

Este repositório **não contém dados reais**.

* Bases Excel reais devem permanecer fora do Git
* Templates HTML podem ser versionados normalmente
* Identificadores sensíveis são gerados apenas em tempo de execução

---

## Licença

MIT
