from datetime import datetime

import pandas as pd

from .config import AppConfig, get
from .file_lock import assert_files_closed
from .history_store import HistoryStore
from .logging_utils import make_logger
from .outlook_client import OutlookClient


def load_html_template(path: str) -> str:
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        return f.read()


def build_email_body_from_template(
    template_html: str,
    nome_assessor: str,
    nome_cliente: str,
    cod_cliente: str,
    estrutura: str,
    ativo: str,
    alocacao_pct: str,
    token: str,
) -> str:
    """
    Monta o corpo do e-mail a partir do template HTML (sanitizado).
    Mantém o conteúdo do questionário em HTML e injeta os campos dinâmicos.
    """
    return template_html.format(
        nome_assessor=nome_assessor,
        nome_cliente=nome_cliente,
        cod_cliente=cod_cliente,
        estrutura=estrutura,
        ativo=ativo,
        alocacao_pct=alocacao_pct,
        token=token,
    )


def run_dispatch(cfg: AppConfig, dry_run: bool, display_only: bool) -> int:
    log = make_logger()

    operations_path = get(cfg, "paths", "operations_xlsx")
    professionals_path = get(cfg, "paths", "professionals_xlsx")
    history_path = get(cfg, "paths", "history_xlsx")
    history_sheet = get(cfg, "paths", "history_sheet", default="Auditoria De Estruturadas")

    # NOVO: template HTML
    email_body_html_path = get(cfg, "paths", "email_body_html", default="templates/email_body.html")

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

    # Carrega template HTML uma única vez
    try:
        template_html = load_html_template(email_body_html_path)
    except Exception as e:
        log.error(f"Falha ao carregar template HTML em '{email_body_html_path}': {e}")
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

            body_html = build_email_body_from_template(
                template_html=template_html,
                nome_assessor=nome_assessor,
                nome_cliente=str(nome_cliente),
                cod_cliente=str(cod_cliente),
                estrutura=str(estrutura),
                ativo=str(ativo),
                alocacao_pct=str(alocacao_pct),
                token=token,
            )

            if dry_run:
                log.info(f"[DRY-RUN] Para: {email_assessor} | CC: {email_lider} | Cliente {cod_cliente}")
                continue

            # IMPORTANTE: enviar como HTML
            outlook.send_mail(
                to=email_assessor,
                cc=email_lider,
                subject=subject,
                body=body_html,
                display_only=display_effective,
                is_html=True,
            )

            ids = outlook.find_sent_ids_by_subject_and_token(
                subject,
                token,
                delay_seconds=send_delay,
                max_items=sent_max,
            )

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
