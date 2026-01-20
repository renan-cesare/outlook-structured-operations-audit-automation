from datetime import datetime
import pandas as pd

from .config import AppConfig, get
from .file_lock import assert_files_closed
from .history_store import HistoryStore
from .logging_utils import make_logger
from .outlook_client import OutlookClient


def build_email_body(
    nome_assessor: str,
    nome_cliente: str,
    estrutura: str,
    ativo: str,
    alocacao_pct: str,
    token: str,
) -> str:
    # Template mais completo (sanitizado), alinhado ao modelo que você usava.
    # Obs.: Mantém placeholders e campos de resposta para o assessor preencher.
    return f"""
Prezado(a) {nome_assessor},

Estamos entrando em contato no âmbito de processo de auditoria interna referente à alocação da operação estruturada **{estrutura}**, vinculada ao ativo **{ativo}**, realizada na carteira do cliente **{nome_cliente}**.

A referida alocação representa aproximadamente **{alocacao_pct}%** do portfólio do cliente. Solicitamos, portanto, seu retorno com o preenchimento das informações abaixo, para conclusão do processo de auditoria.

================================================================================
SEÇÃO 1 – CONTEXTO E PERFIL DO CLIENTE
1.1) Qual o contexto do cliente e objetivo da recomendação (horizonte, necessidade, perfil)?
R:
[______________________________________________]

1.2) Houve alguma restrição/condição específica considerada (liquidez, tolerância a perdas, prazo, etc.)?
R:
[______________________________________________]

================================================================================
SEÇÃO 2 – PROCESSO DE VENDA E COMUNICAÇÃO
2.1) Descreva detalhadamente o processo de venda desta operação (abordagem, explicação da mecânica e riscos):
R:
[______________________________________________]

2.2) Como foram apresentados os principais riscos (cenários, perdas possíveis, marcação a mercado, vencimento, etc.)?
R:
[______________________________________________]

2.3) Houve comunicação por escrito e/ou registro de aceite? Onde?
R:
[______________________________________________]

================================================================================
SEÇÃO 3 – JUSTIFICATIVA PARA A ALOCAÇÃO
3.1) Qual foi o racional para alocar essa operação na carteira do cliente?
R:
[______________________________________________]

3.2) Quais alternativas foram consideradas e por que esta foi escolhida?
R:
[______________________________________________]

3.3) Como foi definida a quantidade/alocação (por que {alocacao_pct}% do PL)?
R:
[______________________________________________]

================================================================================
SEÇÃO 4 – ADEQUAÇÃO E GOVERNANÇA (SUITABILITY / CONTROLES)
4.1) Em sua avaliação, a operação é adequada ao perfil e objetivo do cliente? Justifique.
R:
[______________________________________________]

4.2) Existe algum ponto de atenção que você considera relevante mencionar para registro?
R:
[______________________________________________]

================================================================================
Observações adicionais (se aplicável):
[______________________________________________]

Prazo: solicitamos devolutiva em até **3 (três) dias úteis**, para finalizarmos o processo de auditoria.

Atenciosamente,
Equipe de Auditoria (Sanitized)

{token}
""".lstrip()


def run_dispatch(cfg: AppConfig, dry_run: bool, display_only: bool) -> int:
    log = make_logger()

    operations_path = get(cfg, "paths", "operations_xlsx")
    professionals_path = get(cfg, "paths", "professionals_xlsx")
    history_path = get(cfg, "paths", "history_xlsx")
    history_sheet = get(cfg, "paths", "history_sheet", default="Auditoria De Estruturadas")

    send_delay = int(get(cfg, "outlook", "send_delay_seconds", default=3))
    sent_max = int(get(cfg, "outlook", "search_sent_max_items", default=300))
    status_label = str(get(cfg, "dispatch", "status_sent_label", default="Enviado"))
    subject_template = str(
        get(
            cfg,
            "dispatch",
            "email_subject_template",
            default="Análise de Alocação em Operações Estruturadas – Cliente {nome_cliente} – {cod_cliente}",
        )
    )

    display_default = bool(get(cfg, "run_mode", "display_only_default", default=False))
    display_effective = bool(display_only or display_default)

    if not operations_path or not professionals_path or not history_path:
        log.error("Config inválida: paths.operations_xlsx / paths.professionals_xlsx / paths.history_xlsx são obrigatórios.")
        return 2

    try:
        assert_files_closed([operations_path, professionals_path, history_path])
    except Exception as e:
        log.error(str(e))
        return 2

    base = pd.read_excel(professionals_path)
    base.columns = base.columns.str.strip()
    base = base.rename(
        columns={
            "Código Assessor": "codigo_assessor",
            "Nome Completo": "nome",
            "E-mail": "email",
            "Código do Líder": "codigo_lider",
        }
    )

    ops = pd.read_excel(operations_path)
    ops.columns = ops.columns.str.strip()

    required = ["Código Cliente", "Nome do Cliente", "Estrutura", "Ativo", "% PL", "Assessor da Operação", "Assessor do Cliente"]
    missing = [c for c in required if c not in ops.columns]
    if missing:
        log.error(f"Planilha de operações sem colunas obrigatórias: {missing}")
        return 2

    outlook = OutlookClient()
    store = HistoryStore(history_path=history_path, sheet_name=history_sheet)

    for i, row in ops.iterrows():
        try:
            cod_cliente = row.get("Código Cliente")
            nome_cliente = row.get("Nome do Cliente")
            estrutura = row.get("Estrutura")
            ativo = row.get("Ativo")
            alocacao_pct = row.get("% PL")
            cod_assessor = row.get("Assessor da Operação")
            cod_lider = row.get("Assessor do Cliente")

            if any(pd.isna(x) for x in [cod_cliente, nome_cliente, estrutura, ativo, cod_assessor, cod_lider]):
                log.error(f"Linha {i+2}: dados obrigatórios ausentes.")
                continue

            try:
                dados_assessor = base[base["codigo_assessor"] == cod_assessor].iloc[0]
                dados_lider = base[base["codigo_assessor"] == cod_lider].iloc[0]
            except Exception:
                log.error(f"Cliente {nome_cliente} - {cod_cliente}: assessor/líder não encontrados na base.")
                continue

            email_assessor = str(dados_assessor.get("email", "")).strip()
            nome_assessor = str(dados_assessor.get("nome", "")).strip()
            email_lider = str(dados_lider.get("email", "")).strip()

            if not email_assessor or not email_lider:
                log.error(f"Cliente {nome_cliente} - {cod_cliente}: e-mail do assessor ou líder ausente.")
                continue

            token = f"#audit_token:{cod_cliente}_{datetime.now().strftime('%Y%m%d%H%M%S')}"
            subject = subject_template.format(nome_cliente=nome_cliente, cod_cliente=cod_cliente)

            body = build_email_body(
                nome_assessor=nome_assessor,
                nome_cliente=nome_cliente,
                estrutura=str(estrutura),
                ativo=str(ativo),
                alocacao_pct=str(alocacao_pct),
                token=token,
            )

            if dry_run:
                log.info(f"[DRY-RUN] Para: {email_assessor} | CC: {email_lider} | Cliente {cod_cliente}")
                continue

            outlook.send_mail(email_assessor, email_lider, subject, body, display_only=display_effective)
            ids = outlook.find_sent_ids_by_subject_and_token(subject, token, delay_seconds=send_delay, max_items=sent_max)

            store.append_dispatch_record(
                operation_row=row.to_dict(),
                email_assessor=email_assessor,
                email_lider=email_lider,
                assunto=subject,
                token=token,
                status=status_label,
                conversation_id=ids.conversation_id,
                internet_id=ids.internet_message_id,
                entry_id=ids.entry_id,
            )

            log.ok(f"Cliente {cod_cliente}: enviado e registrado.")
        except Exception as e:
            log.error(f"Linha {i+2}: {e}")

    return 0
