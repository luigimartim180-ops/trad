"""pt-PT -> pt-BR conversion pipeline for the Ulyses review workbook.

Reads ``Revision_traducciones_PT.xlsx`` and produces:

* ``revisao_pt_BR_completa.xlsx`` -- the review sheets with the corrected
  ``Estado`` / ``Corrección PT`` / ``Comentario`` columns;
* ``duvidas_PO.xlsx`` -- the open questions for the Product Owner, untouched;
* ``correcoes_backend.csv`` -- backend key/state/correction/comment list.

The workbook layout is detected generically:

* the ``Instrucciones`` sheet is informational only;
* review sheets are identified by having a key column (``Clave`` or
  ``Clave técnica (no editar)``) and a pt-PT text column (``Traducción PT
  actual``);
* the header row may be preceded by title rows (detected by searching for
  ``Clave``/``Código``).

Only the ``Estado``, ``Corrección PT`` and ``Comentario`` columns are written;
the ES/EN source text is never regenerated.
"""

from __future__ import annotations

import csv
import re
import sys
from pathlib import Path

import pandas as pd

XLSX_IN = "Revision_traducciones_PT.xlsx"
OUT_COMPLETA = "revisao_pt_BR_completa.xlsx"
OUT_REVISAR = "revisao_pt_BR_REVISAR.xlsx"
OUT_DUVIDAS = "duvidas_PO.xlsx"
OUT_CSV = "correcoes_backend.csv"

KEY_COLUMNS = ["Clave", "Clave técnica (no editar)", "Código"]
PT_COLUMNS = ["Traducción PT actual", "Texto PT", "Traducción PT"]
COL_ESTADO = "Estado"
COL_CORRECAO = "Corrección PT"
COL_COMENTARIO = "Comentario"
COL_AVISO = "Aviso"
COL_PRIORIDADE = "Prioridade"
INFO_SHEET = "Instrucciones"
DUVIDAS_SHEET = "Dúvidas"
DEAD_COLUMNS = [
    "A1_Léxico",
    "A2_Gramática",
    "A3_Estilo",
    "A4_Formato",
    "A5_Consistência",
]

# --- glossary decisions -----------------------------------------------------

#: B2B glossary keeps the English term "allotment"; no global cota substitution.
KEEP_AS_IS = {"allotment"}

#: Terms that may legitimately stay identical to ES/EN and need no review.
WHITELIST_IGUAL_ES_EN = {
    "ok",
    "overbooking",
    "revpar",
    "json",
    "check-in",
    "check-out",
    "day use",
    "voucher",
    "pax",
    "fiscal",
    "central",
    "empresa",
    "canal",
}

#: Keys without a translation in the source workbook (kept as-is).
TRADUCOES_EM_FALTA: list[str] = []

#: Keys whose translation is under discussion with the PO (never converted).
#: Grouped by the four open questions in ``glossario_pt_BR.md`` §3.
DUVIDAS_GRUPOS: dict[str, list[str]] = {
    "D1 pensão/regime": ["PropertyClassType t WHERE t.Code = '9'"],
    "D2 allotment": ["shared.forecast.nonDeductedAllotment"],
    "D3 OnHold/OutOfOrder/OutOfService": [
        "apps.chainSpaceLockType.radio.onHold",
        "apps.chainSpaceLockType.radio.outOfOrder",
        "apps.chainSpaceLockType.radio.outOfService",
    ],
    "D4 aiAssistant hotel/property": ["apps.aiAssistant.label.hotel"],
}

DUVIDAS_KEYS = [
    "core.label.billing",
    "core.label.chain",
    "shared.forecast.nonDeductedAllotment",
]

#: Surgical, key-scoped overrides that bypass the generic rules. ``None``
#: marks a key that must never be auto-converted (open PO question).
CORRECOES: dict[str, str | None] = {
    "card.hojaRegistro": "Ficha de registro de hóspedes (FNRH)",
    "core.button.settlement": "Cobranças automáticas",
    "invoice.direction.label": "Endereço",
    # "to download" is "baixar" in pt-BR; only the UI download buttons change,
    # never the account/folio "transfer" actions.
    "core.button.download": "Baixar",
    "core.button.downloadDocument": "Baixar documento",
    "core.button.downloadFileAutomatically": "Baixar arquivo automaticamente",
    "core.button.downloadNDocuments": "Baixar documentos ({{n}})",
    "core.button.downloadTemplate": "Baixar modelo",
    "core.navbar.user.logs": "Baixar logs",
    "apps.report.downloadCSV": "Baixar CSV",
    "apps.report.downloadExcel": "Baixar Excel",
    "apps.report.downloadPDF": "Baixar PDF",
    "apps.management.button.exportExcel": "Baixar Excel",
    "apps.management.customExcel": "Baixar Excel com fórmulas",
    "apps.budget.importExcelModal.downloadExcel": "Baixar modelo",
    "apps.rate.importExcelModal.downloadExcel": "Baixar modelo",
    "apps.integration.septeoChannel.button.downloadTemplateFiles": "Baixar modelo",
    "apps.integration.siteminder.button.downloadTemplateFiles": "Baixar modelo",
    "apps.integration.accounting.downloadFile": "Baixar arquivos",
    "apps.integration.nf525.action.downloadArchives": "Baixar individualmente",
    "apps.integration.nf525.action.downloadArchivesPeriod": "Baixar agrupado",
    "shared.modal.bulkDownloadFile.title": "Baixando arquivos em massa",
    "shared.modal.printInvoice.downloadNInvoice": "Baixar {{count}} faturas",
    "shared.switch.allowDownloadInvoice.title": "Baixar faturas",
    "shared.switch.allowInHouseDownloadInvoice.title": "Baixar faturas",
    "shared.switch.downloadReservation.title": "Baixar reserva",
    "shared.switch.downloadReservation.message": (
        "Ative esta opção para baixar automaticamente as reservas"
    ),
    # AO90 spelling -> Brazilian spelling
    "core.button.transferEntity": "Transferir {{entity}}",
    # refined pt-BR wording (pt-PT residue or idiom fixes in the suggestion)
    "core.label.amountOtherBreakdownAfterTax": "Outros valores",
    "core.label.amountPaidOut": "Valor retirado",
    "core.label.totalAmountEntered": "Valor total informado",
    "core.label.taxExcludedShort": "Imposto excl.",
    "core.label.taxIncludedShort": "Imposto incl.",
    "apps.integration.accounting.trevenque.check": "Tem certeza de que deseja verificar?",
    "apps.import.importJob.import.error": (
        "Existem alguns registros que não puderam ser importados. "
        "Utilize esta opção se quiser tentar novamente."
    ),
    "shared.modal.information.anonymize.message": (
        "Este processo excluirá todos os dados pessoais deste cliente para "
        "torná-los anônimos. Tem certeza de que deseja continuar?"
    ),
    # multiline messages with pt-PT residue
    "apps.budget.importExcelModal.instructions": (
        "Envie aqui a previsão atualizada (receita, PAX e número de acomodações). "
        "Confirme abaixo para tornar as alterações visíveis. Tenha em conta que "
        "as alterações não serão salvas até clicar em Salvar na tela principal."
        "\n\nBaixe o modelo Excel caso ainda não o tenha feito."
    ),
    "before.checkout.message.intro": (
        "Esperamos que tenha gostado da sua estadia em {0}. Para tornar a sua "
        "partida mais cômoda, sem contato e sem passar pela recepção amanhã, "
        "você já pode fazer o check-out online."
    ),
    "apps.login.section.passwordReset.descriptionSystem": (
        "Por motivos de segurança, a senha do Ulyses Suite tem de ser atualizada "
        "a cada 90 dias. Lembre-se de não reutilizar nenhuma das últimas 4 senhas "
        "que utilizou para iniciar sessão no Ulyses Suite."
    ),
    "apps.integration.publicApiRest.refreshEventEntity.message": (
        "Aplique esta ação se deseja reenviar informações sobre a entidade de "
        "evento selecionada.\nTenha em conta que as alterações não serão "
        "atualizadas de imediato; por isso, recomendamos que consulte a aba Fila "
        "de eventos para acompanhar a evolução."
    ),
    "apps.integration.publicApiRest.removeEventEntity.message": (
        "Aplique esta ação se deseja excluir informações sobre a entidade de "
        "evento selecionada.\nTenha em conta que as alterações não serão "
        "atualizadas de imediato; por isso, recomendamos que consulte a aba Fila "
        "de eventos para acompanhar a evolução."
    ),
    # translations missing in the source workbook
    "shared.assistant.propertyAutomatedCheckinScheduler.step0": "Resumo",
    "shared.assistant.propertyAutomatedCheckinScheduler.step1": "Filtros",
    "shared.assistant.propertyAutomatedCheckinScheduler.step2": "Ação",
    "shared.assistant.propertyAutomatedCheckinScheduler.step3": "Revisão",
    "core.object.entity.propertyCheckInAutomationScheduler": "Check-in automático",
    "apps.import.importJob.relatedMap.chainProductSubGroup": "Subgrupo de produtos",
    "apps.import.importJob.relatedMap.chainVIPLevel": "Nível VIP",
    "apps.import.importJob.relatedMap.communicationLocationType": "Tipo de contato",
    "apps.import.importJob.relatedMap.depositStatusType": "Status do depósito",
    "shared.aggregateGrid.empty.chainProductTaxIncluded.noIncludedProductTax": (
        "Adicione os produtos que podem ser utilizados neste produto."
    ),
    # open PO questions: never converted
    "apps.profileCommissionTransaction.gridAction.offset": None,
    "PropertyClassType t WHERE t.Code = '9'": None,
    "apps.chainSpaceLockType.radio.onHold": None,
    "apps.chainSpaceLockType.radio.outOfOrder": None,
    "apps.chainSpaceLockType.radio.outOfService": None,
    "apps.aiAssistant.label.hotel": None,
}

#: Keys closed as OK by the reviewer: the term is already valid pt-BR and the
#: "Igual a ES/EN" warning is a false positive.
OK_MANUAL: set[str] = {
    # core.label
    "core.label.total",
    "core.label.totalADR",
    "core.label.totalPriceConsumed",
    "core.label.tpv",
    "core.label.tpvpc",
    "core.label.tvaCode",
    "core.label.manually",
    "core.label.maximumTicketNumber",
    "core.label.minimumTicketNumber",
    "core.label.mix",
    "core.label.menu",
    "core.label.method",
    "core.label.iata",
    "core.label.iban",
    "core.label.ical",
    "core.label.id",
    "core.label.ideas",
    "core.label.iistat",
    "core.label.ine",
    "core.label.inquiry",
    "core.label.inHouse",
    "core.label.incomplete",
    "core.label.mcrMelia",
    # shared.sectionPanel / shared.assistant / shared.switch / apps
    "shared.sectionPanel.notes",
    "shared.sectionPanel.offer",
    "shared.assistant.productTypeCancellationFee.step2",
    "shared.assistant.propertyAutomatedCheckInScheduler.step1",
    "shared.assistant.propertyPaymentScheduler.step1",
    "shared.switch.amountIncluded.title",
    "apps.cache.headers.lastStatus",
    "apps.availability.footer.confirm",
    "apps.billing.folio.action.export",
    "apps.billing.folio.action.group",
    "apps.billing.folio.action.import",
    "apps.billing.folio.action.refund",
    "apps.integration.retailForce.mapping.button.documentTypeMapping",
    "core.object.entity.propertyBENotice",
    "core.object.entity.propertyBETheme",
}


def _rule(pt: str, br: str) -> tuple[re.Pattern[str], str]:
    # Word boundaries are essential: without them "a carregar" would also match
    # inside "par(a carregar)" and "utilizador" inside "utilizadores".
    pattern = r"\b" + re.escape(pt).replace(r"\ ", r"\s+") + r"\b"
    return re.compile(pattern, re.IGNORECASE), br


#: Ordered substitution rules. Longer/compound forms MUST come before the
#: shorter forms they contain ("a carregar" before "carregar"; "utilizadores"
#: before "utilizador") so the first matching rule wins.
SUBSTITUICOES: list[tuple[re.Pattern[str], str]] = [
    # 1. compound expressions first
    _rule("por defeito", "por padrão"),
    _rule("a carregar", "carregando"),
    _rule("a digitalizar", "digitalizando"),
    _rule("tem a certeza", "tem certeza"),
    _rule("gestor de conta", "gerente de contas"),
    _rule("folha de cálculo", "planilha"),
    # 2. plurals / feminine forms before their singulars
    _rule("palavras-passe", "senhas"),
    _rule("utilizadores", "usuários"),
    _rule("utilizadora", "usuária"),
    _rule("ficheiros", "arquivos"),
    _rule("montantes", "valores"),
    _rule("telemóveis", "celulares"),
    _rule("predefinições", "padrões"),
    _rule("receções", "recepções"),
    _rule("subscrições", "assinaturas"),
    _rule("faturações", "faturamentos"),
    _rule("moradas", "endereços"),
    _rule("cadeias", "redes"),
    _rule("fechos", "fechamentos"),
    _rule("registos", "registros"),
    _rule("planeamentos", "planejamentos"),
    _rule("ecrãs", "telas"),
    _rule("secções", "seções"),
    _rule("demasiados", "muitos"),
    _rule("demasiadas", "muitas"),
    _rule("contactos", "contatos"),
    _rule("actividades", "atividades"),
    _rule("detetados", "detectados"),
    _rule("detetadas", "detectadas"),
    # 3. remaining single verbs / nouns
    _rule("palavra-passe", "senha"),
    _rule("secção", "seção"),
    _rule("contacto", "contato"),
    _rule("carregar", "enviar"),
    _rule("utilizador", "usuário"),
    _rule("ficheiro", "arquivo"),
    _rule("guardar", "salvar"),
    _rule("eliminar", "excluir"),
    _rule("transferir", "transferir"),
    _rule("ecrã", "tela"),
    _rule("telemóvel", "celular"),
    _rule("registo", "registro"),
    # 4. extended glossary
    _rule("morada", "endereço"),
    _rule("predefinição", "padrão"),
    _rule("fecho", "fechamento"),
    _rule("receção", "recepção"),
    _rule("detetada", "detectada"),
    _rule("detetado", "detectado"),
    _rule("montante", "valor"),
    _rule("faturação", "faturamento"),
    _rule("planeamento", "planejamento"),
    _rule("enviámos", "enviamos"),
    _rule("subscrição", "assinatura"),
    _rule("cadeia", "rede"),
    _rule("demasiado", "muito"),
    _rule("demasiada", "muita"),
    _rule("Letónia", "Letônia"),
    _rule("Roménia", "Romênia"),
    _rule("Iémen", "Iêmen"),
    _rule("Zimbabué", "Zimbábue"),
    # 5. AO90 -> Brazilian spelling
    _rule("acção", "ação"),
    _rule("acções", "ações"),
    _rule("actividade", "atividade"),
    _rule("actividades", "atividades"),
    _rule("activo", "ativo"),
    _rule("activa", "ativa"),
    _rule("actual", "atual"),
    _rule("objectivo", "objetivo"),
    _rule("objectivos", "objetivos"),
    _rule("objecto", "objeto"),
    _rule("directo", "direto"),
    _rule("director", "diretor"),
    _rule("colectivo", "coletivo"),
    _rule("protecção", "proteção"),
    _rule("correcção", "correção"),
    _rule("seleccionar", "selecionar"),
    _rule("colecção", "coleção"),
    # 6. pt-PT verb forms / participles that leak into the suggestions
    _rule("de que pretende", "de que deseja"),
    _rule("se pretender", "se quiser"),
    _rule("se pretende", "se deseja"),
    _rule("pretendidos", "desejados"),
    _rule("pretendidas", "desejadas"),
    _rule("pretendido", "desejado"),
    _rule("pretendida", "desejada"),
    _rule("pretender", "querer"),
    _rule("pretenda", "deseje"),
    _rule("pretende", "deseja"),
    _rule("encontrámos", "encontramos"),
    _rule("detetámos", "detectamos"),
    _rule("guardados", "salvos"),
    _rule("guardadas", "salvas"),
    _rule("guardado", "salvo"),
    _rule("guardada", "salva"),
    _rule("eliminados", "excluídos"),
    _rule("eliminadas", "excluídas"),
    _rule("eliminado", "excluído"),
    _rule("eliminada", "excluída"),
    _rule("introduzidos", "informados"),
    _rule("introduzidas", "informadas"),
    _rule("introduzido", "informado"),
    _rule("introduzida", "informada"),
    _rule("introduzir", "informar"),
    _rule("introduzam", "informem"),
    _rule("introduza", "informe"),
    _rule("bebés", "bebês"),
    _rule("bebé", "bebê"),
    _rule("cómodos", "cômodos"),
    _rule("cómodas", "cômodas"),
    _rule("cómodo", "cômodo"),
    _rule("cómoda", "cômoda"),
    _rule("anulações", "cancelamentos"),
    _rule("anulação", "cancelamento"),
    _rule("eliminação", "exclusão"),
    _rule("detetar", "detectar"),
    _rule("carregadas", "enviadas"),
    _rule("carregada", "enviada"),
    _rule("todos os separadores", "todas as abas"),
    _rule("separadores", "abas"),
    _rule("separador", "aba"),
]

_PLACEHOLDER_RE = re.compile(r"\{\{.*?\}\}|\{\d+\}|<[^>]+>")
_TOKEN = "\x00PH{}\x00"


def _adapt_case(template: str, source: str) -> str:
    """Mirror the capitalisation of ``source`` onto ``template``."""
    if source.isupper():
        return template.upper()
    if source[:1].isupper():
        return template[:1].upper() + template[1:]
    return template


def _protect(text: str) -> tuple[str, list[str]]:
    saved: list[str] = []

    def _stash(match: re.Match[str]) -> str:
        saved.append(match.group(0))
        return _TOKEN.format(len(saved) - 1)

    return _PLACEHOLDER_RE.sub(_stash, text), saved


def _restore(text: str, saved: list[str]) -> str:
    for index, value in enumerate(saved):
        text = text.replace(_TOKEN.format(index), value)
    return text


def converter(texto: str) -> str:
    """Convert a single pt-PT string to pt-BR."""
    if not isinstance(texto, str) or not texto:
        return texto

    protected, saved = _protect(texto)

    for pattern, repl in SUBSTITUICOES:
        # Direct re.sub(pattern, repl, text): never a group reference, so a
        # None group can never raise TypeError.
        protected = pattern.sub(
            lambda m, r=repl: _adapt_case(r, m.group(0)), protected
        )

    return _restore(protected, saved)


def _find_header_row(raw: pd.DataFrame) -> int:
    for idx in range(min(len(raw), 30)):
        cells = [str(v).strip() for v in raw.iloc[idx].tolist() if pd.notna(v)]
        if any(col in cells for col in KEY_COLUMNS) or "Código" in cells:
            return idx
    raise ValueError("header row with 'Clave'/'Código' not found")


def _read_sheet(xl: pd.ExcelFile, sheet: str) -> pd.DataFrame:
    raw = pd.read_excel(xl, sheet_name=sheet, header=None)
    header_idx = _find_header_row(raw)
    header = [str(v).strip() for v in raw.iloc[header_idx].tolist()]
    df = raw.iloc[header_idx + 1 :].reset_index(drop=True)
    df.columns = header
    return df


def _is_review_sheet(columns: list[str]) -> bool:
    has_key = any(col in columns for col in KEY_COLUMNS)
    has_pt = any(col in columns for col in PT_COLUMNS)
    return has_key and has_pt


def _pick(columns: list[str], candidates: list[str]) -> str | None:
    for name in candidates:
        if name in columns:
            return name
    return None


def _key_value(row: pd.Series, key_col: str | None) -> str:
    if not key_col:
        return ""
    value = row.get(key_col)
    return "" if pd.isna(value) else str(value).strip()


def _resolve_estado(aviso: str, pt: str, converted: str) -> str:
    aviso_low = (aviso or "").lower()
    pt_norm = (pt or "").strip().lower()

    if "sin traducción" in aviso_low or "sem tradução" in aviso_low:
        return "REVISAR"
    if "igual a es" in aviso_low or "igual a en" in aviso_low:
        if pt_norm in WHITELIST_IGUAL_ES_EN:
            return "OK"
        return "REVISAR"
    if not pt.strip():
        return "REVISAR"
    return "OK" if converted == pt else "REVISAR"


def _clean(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value)


def _prioridade(estado: str, tem_correcao: bool, e_duvida: bool) -> str:
    if estado == "OK":
        return ""
    if e_duvida:
        return "P0 - dúvida PO"
    if not tem_correcao:
        return "P1 - decisão manual"
    return "P2 - validar sugestão"


def process_sheet(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    columns = list(out.columns)
    key_col = _pick(columns, KEY_COLUMNS)
    pt_col = _pick(columns, PT_COLUMNS)
    aviso_col = _pick(columns, [COL_AVISO])

    for col in (COL_ESTADO, COL_CORRECAO, COL_COMENTARIO, COL_PRIORIDADE):
        if col not in out.columns:
            out[col] = ""
        out[col] = out[col].map(_clean)

    for idx, row in out.iterrows():
        clave = _key_value(row, key_col)
        texto_pt = _clean(row.get(pt_col)) if pt_col else ""
        aviso = _clean(row.get(aviso_col)) if aviso_col else ""

        if clave in OK_MANUAL:
            # closed by the reviewer: already-valid pt-BR term, warning is a
            # false positive
            out.at[idx, COL_ESTADO] = "OK"
            out.at[idx, COL_CORRECAO] = ""
            out.at[idx, COL_PRIORIDADE] = "Fechado"
            continue

        if clave in DUVIDAS_KEYS or CORRECOES.get(clave, "") is None:
            # open question for the PO: never auto-converted, listed in
            # duvidas_PO.xlsx for a human decision
            out.at[idx, COL_ESTADO] = "REVISAR"
            out.at[idx, COL_CORRECAO] = ""
            out.at[idx, COL_PRIORIDADE] = _prioridade("REVISAR", False, True)
            continue

        if clave in CORRECOES:
            converted = CORRECOES[clave]
        else:
            converted = converter(texto_pt)

        estado = _resolve_estado(aviso, texto_pt, converted)

        if estado == "OK":
            out.at[idx, COL_CORRECAO] = ""
        else:
            out.at[idx, COL_CORRECAO] = (
                converted if converted != texto_pt else ""
            )
        out.at[idx, COL_ESTADO] = estado
        if clave in CORRECOES:
            # reviewed and closed by the reviewer
            out.at[idx, COL_PRIORIDADE] = "Fechado"
        else:
            out.at[idx, COL_PRIORIDADE] = _prioridade(
                estado, bool(_clean(out.at[idx, COL_CORRECAO])), False
            )

    return out.drop(columns=[c for c in DEAD_COLUMNS if c in out.columns])


def _review_sheet_names(xl: pd.ExcelFile) -> list[str]:
    names: list[str] = []
    for sheet in xl.sheet_names:
        if sheet in (INFO_SHEET, DUVIDAS_SHEET):
            continue
        head = pd.read_excel(xl, sheet_name=sheet, header=None, nrows=30)
        try:
            header_idx = _find_header_row(head)
        except ValueError:
            continue
        columns = [str(v).strip() for v in head.iloc[header_idx].tolist()]
        if _is_review_sheet(columns):
            names.append(sheet)
    return names


def _duvida_grupo(clave: str) -> str:
    for grupo, keys in DUVIDAS_GRUPOS.items():
        if clave in keys:
            return grupo
    if clave in DUVIDAS_KEYS:
        return "glossário"
    return ""


def _extract_duvidas(
    xl: pd.ExcelFile, processed: dict[str, pd.DataFrame]
) -> pd.DataFrame:
    if DUVIDAS_SHEET in xl.sheet_names:
        return pd.read_excel(xl, sheet_name=DUVIDAS_SHEET)

    rows: list[dict[str, object]] = []
    for sheet, df in processed.items():
        columns = list(df.columns)
        key_col = _pick(columns, KEY_COLUMNS)
        pt_col = _pick(columns, PT_COLUMNS)
        for _, row in df.iterrows():
            clave = _key_value(row, key_col)
            grupo = _duvida_grupo(clave)
            if not grupo:
                continue
            rows.append(
                {
                    "Clave": clave,
                    "Hoja": sheet,
                    "Texto PT": _clean(row.get(pt_col)) if pt_col else "",
                    "Grupo": grupo,
                    "Comentario": _clean(row.get(COL_COMENTARIO)),
                    "Destino": "PO",
                }
            )
    return pd.DataFrame(
        rows,
        columns=["Clave", "Hoja", "Texto PT", "Grupo", "Comentario", "Destino"],
    )


def build() -> int:
    root = Path(__file__).resolve().parent
    xlsx_path = root / XLSX_IN
    if not xlsx_path.exists():
        print(f"ERRO: {XLSX_IN} não encontrado em {root}", file=sys.stderr)
        return 1

    xl = pd.ExcelFile(xlsx_path)
    review_sheets = _review_sheet_names(xl)
    if not review_sheets:
        print("ERRO: nenhuma aba de revisão encontrada", file=sys.stderr)
        return 1

    processed: dict[str, pd.DataFrame] = {}
    for sheet in review_sheets:
        processed[sheet] = process_sheet(_read_sheet(xl, sheet))

    with pd.ExcelWriter(root / OUT_COMPLETA) as writer:
        for sheet, df in processed.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)

    with pd.ExcelWriter(root / OUT_REVISAR) as writer:
        for sheet, df in processed.items():
            pendentes = df[df[COL_ESTADO] != "OK"]
            if not pendentes.empty:
                pendentes.to_excel(writer, sheet_name=sheet[:31], index=False)

    duvidas = _extract_duvidas(xl, processed)
    duvidas.to_excel(root / OUT_DUVIDAS, index=False)

    rows = pd.concat(processed.values(), ignore_index=True)
    backend = rows[rows[COL_ESTADO] != "OK"]
    with open(root / OUT_CSV, "w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh, delimiter=";")
        writer.writerow(
            [KEY_COLUMNS[0], COL_ESTADO, COL_CORRECAO, COL_COMENTARIO, COL_PRIORIDADE]
        )
        for _, row in backend.iterrows():
            key = next(
                (_key_value(row, c) for c in KEY_COLUMNS if c in row.index), ""
            )
            writer.writerow(
                [
                    key,
                    _clean(row.get(COL_ESTADO)),
                    _clean(row.get(COL_CORRECAO)),
                    _clean(row.get(COL_COMENTARIO)),
                    _clean(row.get(COL_PRIORIDADE)),
                ]
            )

    contagem = rows[COL_ESTADO].value_counts().to_dict()
    print("Contagem final por estado:")
    for estado, total in sorted(contagem.items()):
        print(f"  {estado}: {total}")
    print(f"Total de linhas: {len(rows)}")
    print(f"Substituições distintas: {len(SUBSTITUICOES)}")
    print(f"Dúvidas (não convertidas): {len(duvidas)}")

    prioridades = (
        rows[rows[COL_ESTADO] != "OK"][COL_PRIORIDADE].value_counts().to_dict()
    )
    print("Prioridades dos pendentes:")
    for prio, total in sorted(prioridades.items()):
        print(f"  {prio}: {total}")
    return 0


if __name__ == "__main__":
    raise SystemExit(build())
