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

import argparse
import csv
import json
import re
import sys
from pathlib import Path

import pandas as pd

XLSX_IN = "Revision_traducciones_PT.xlsx"
OUT_COMPLETA = "revisao_pt_BR_completa.xlsx"
OUT_FINAL = "revisao_pt_BR_FINAL.xlsx"
OUT_DUVIDAS = "duvidas_PO.xlsx"
OUT_CSV = "correcoes_backend.csv"
OUT_DECISOES = "decisoes_revisao.jsonl"

#: colunas com texto em pt-PT que NÃO podem aparecer no ficheiro de entrega
_PT_SOURCE_COLUMNS = ("Traducción PT actual", "Texto origen (IA)")
#: coluna única com o texto final, no ficheiro de entrega limpo
COL_FINAL = "Texto final pt-BR"

KEY_COLUMNS = ["Clave", "Clave técnica (no editar)", "Código"]
PT_COLUMNS = ["Traducción PT actual", "Texto PT", "Traducción PT"]
COL_ESTADO = "Estado"
COL_CORRECAO = "Corrección PT"
COL_COMENTARIO = "Comentario"
COL_AVISO = "Aviso"
COL_PRIORIDADE = "Prioridade"
INFO_SHEET = "Instrucciones"
DUVIDAS_SHEET = "Dúvidas"

#: Estado values. ``OK`` = already valid pt-BR (nothing to patch);
#: ``CORRIGIDO`` = validated pt-BR suggestion, exported to the backend;
#: ``PENDENTE_PO`` = glossary decision owned by the PO, never auto-converted.
ESTADO_OK = "OK"
ESTADO_CORRIGIDO = "CORRIGIDO"
ESTADO_PENDENTE_PO = "PENDENTE_PO"

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
    # valid icons/short labels identical to ES that a reviewer would close
    "total",
    "tipo",
    "por",
    "estado",
    "código",
    "hora",
    "ver",
    "abrir",
    "editar",
    "enviar",
    "cliente",
    "grupo",
    "hotel",
    "aviso",
    "nota",
    "ofertas",
    "exportar",
    "importar",
    "reembolsar",
    "confirmar",
    "reserva",
    "reservas",
    "depósito",
    "suplemento",
    "número",
    "valor",
}

#: Words that are wrong in pt-BR (Spanish-only or pt-PT spelling). They stay
#: PENDENTE_PO even if a reviewer mass-closes the "Igual a ES/EN" bucket.
#: NOTE: "Cancelar", "Transferir", "Compensar" are valid pt-BR and are NOT
#: listed here, despite being identical to Spanish.
NOT_OK_IGUAL_ES_EN = {
    # pt-PT spelling that changes in pt-BR
    "acrónimo",
    "acrónimos",
    "género",
    # Spanish-only vocabulary with a different pt-BR term
    "cerrar",
    "borrar",
    "crear",
    "adjunto",
    "importe",
    "cupo",
    "habitación",
    "huésped",
    "guardar",
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
    # report generation option: "download now" sense, not an account transfer
    "apps.report.options.information": (
        "Selecione como deseja gerar o seu relatório:\n\n"
        "Baixar agora: obtenha o seu relatório imediatamente.\n\n"
        "Processar em fila: o seu relatório será adicionado a uma fila de "
        "processamento e gerado em segundo plano. Será notificado quando "
        "estiver pronto para baixar."
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
        "atualizadas imediatamente; por isso, recomendamos que consulte a aba Fila "
        "de eventos para acompanhar a evolução."
    ),
    "apps.integration.publicApiRest.removeEventEntity.message": (
        "Aplique esta ação se deseja excluir informações sobre a entidade de "
        "evento selecionada.\nTenha em conta que as alterações não serão "
        "atualizadas imediatamente; por isso, recomendamos que consulte a aba Fila "
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
    _rule("de imediato", "imediatamente"),
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
    _rule("predefinidos", "padrão"),
    _rule("predefinidas", "padrão"),
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
    _rule("casas de banho", "banheiros"),
    _rule("apelidos", "sobrenomes"),
    _rule("contactaram", "contataram"),
    _rule("contactámos", "contatamos"),
    # 3. remaining single verbs / nouns
    _rule("palavra-passe", "senha"),
    _rule("secção", "seção"),
    _rule("contacto", "contato"),
    _rule("casa de banho", "banheiro"),
    _rule("apelido", "sobrenome"),
    _rule("contactar", "contatar"),
    _rule("contacte", "contate"),
    _rule("contactem", "contatem"),
    _rule("contactado", "contatado"),
    _rule("contactada", "contatada"),
    _rule("contactável", "contatável"),
    _rule("connosco", "conosco"),
    _rule("convosco", "com você"),
    _rule("pequenos-almoços", "cafés da manhã"),
    _rule("pequeno-almoço", "café da manhã"),
    _rule("registe", "registre"),
    _rule("registem", "registrem"),
    _rule("registou", "registrou"),
    _rule("registaram", "registraram"),
    _rule("registámos", "registramos"),
    _rule("registando", "registrando"),
    _rule("registado", "registrado"),
    _rule("registada", "registrada"),
    _rule("registadora", "registradora"),
    _rule("registar", "registrar"),
    # "carregar" -> "enviar" only in the upload sense: followed by a file/media
    # object (optionally with an article) or standing alone as a button label.
    # Pagination ("carregar mais") and load errors ("não foi possível
    # carregar") keep the pt-BR verb, which is valid in pt-BR.
    (
        re.compile(
            r"\bcarregar\b(?=\s*(?:$|\x00|\{\{|(?:(?:o|a|os|as|um|uma)\s+)?"
            r"(?:ficheiro|ficheiros|arquivo|arquivos|documento|documentos|"
            r"imagem|imagens|m[ée]dia|v[íi]deo|v[íi]deos|folha|planilha|modelo|"
            r"excel)))",
            re.IGNORECASE | re.MULTILINE,
        ),
        "enviar",
    ),
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
    _rule("predefinido", "padrão"),
    _rule("predefinida", "padrão"),
    _rule("fecho", "fechamento"),
    _rule("receção", "recepção"),
    _rule("detetada", "detectada"),
    _rule("detetado", "detectado"),
    _rule("deteção", "detecção"),
    _rule("conceção", "concepção"),
    _rule("recetivo", "receptivo"),
    _rule("recetiva", "receptiva"),
    _rule("infetado", "infectado"),
    _rule("infetada", "infectada"),
    _rule("infeção", "infecção"),
    _rule("infetar", "infectar"),
    _rule("recetor", "receptor"),
    _rule("recetores", "receptores"),
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
    # 7. pt-PT verbs/nouns identified by the native review panel
    _rule("aceder", "acessar"),
    _rule("acede", "acessa"),
    _rule("acedem", "acessam"),
    _rule("acedeu", "acessou"),
    _rule("aceda", "acesse"),
    _rule("acedam", "acessem"),
    # must run AFTER aceda->acesse: pt-BR does not use "aos" after "acesse"
    _rule("acesse aos", "acesse os"),
    _rule("acesse às", "acesse as"),
    _rule("gerir", "gerenciar"),
    _rule("gerido", "gerenciado"),
    _rule("gerida", "gerenciada"),
    _rule("geridos", "gerenciados"),
    _rule("geridas", "gerenciadas"),
    _rule("equipas", "equipes"),
    _rule("equipa", "equipe"),
    _rule("percentagens", "porcentagens"),
    _rule("percentagem", "porcentagem"),
    _rule("grelhas", "grades"),
    _rule("grelha", "grade"),
    _rule("partilhados", "compartilhados"),
    _rule("partilhadas", "compartilhadas"),
    _rule("partilhado", "compartilhado"),
    _rule("partilhada", "compartilhada"),
    _rule("partilhe", "compartilhe"),
    _rule("partilhar", "compartilhar"),
    _rule("partilhável", "compartilhável"),
    _rule("partilháveis", "compartilháveis"),
    _rule("prima", "clique"),
    _rule("recolha", "coleta"),
    _rule("recolhida", "coletada"),
    _rule("recolhidas", "coletadas"),
    _rule("conta-corrente", "conta corrente"),
    # 8. "transferir" = download in these compound forms only; the bare verb
    # stays for account/folio moves (gridAction.transfer, folioProduct.transfer)
    _rule("para transferir", "para baixar"),
    # "a transferir" is only a download when a downloadable object follows;
    # "a transferir dados" is a data migration and keeps the verb
    (
        re.compile(
            r"\ba transferir\b(?=\s+(?:(?:o|a|os|as)\s+)?"
            r"(?:ficheiro|ficheiros|arquivo|arquivos|documento|documentos|"
            r"fatura|faturas|log|logs|excel|csv|pdf|modelo|relat[óo]rio))",
            re.IGNORECASE,
        ),
        "a baixar",
    ),
    _rule("transferir o ficheiro", "baixar o arquivo"),
    _rule("transferir o arquivo", "baixar o arquivo"),
    _rule("transferir o documento", "baixar o documento"),
    _rule("transferir todas as reservas", "baixar todas as reservas"),
    _rule("transferir as reservas", "baixar as reservas"),
    # 9. Brazilian spelling the panel wrongly reported as already valid
    _rule("géneros", "gêneros"),
    _rule("género", "gênero"),
    _rule("acrónimos", "acrônimos"),
    _rule("acrónimo", "acrônimo"),
    _rule("prémios", "prêmios"),
    _rule("prémio", "prêmio"),
    _rule("quilómetros", "quilômetros"),
    _rule("quilómetro", "quilômetro"),
]

_PLACEHOLDER_RE = re.compile(r"\{\{.*?\}\}|\{\d+\}|<[^>]+>")
_TOKEN = "\x00PH{}\x00"

#: pt-PT progressive ("está a carregar") -> pt-BR gerund ("está carregando").
#: The gerund is built by dropping the final "r" and appending "ndo", which
#: also covers every irregular verb (ser->sendo, ter->tendo, ir->indo...).
_PROGRESSIVE_RE = re.compile(
    r"\b(est[áã]o?|estavam|estava|estiverem?|estivermos)\s+a\s+(\w{2,}?)r\b",
    re.IGNORECASE,
)


def _fix_progressive(text: str) -> str:
    def _repl(match: re.Match[str]) -> str:
        aux, verb = match.group(1), match.group(2)
        return f"{aux} {verb}ndo"

    return _PROGRESSIVE_RE.sub(_repl, text)


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

    # Runs after SUBSTITUICOES so a lemma rule ("gerir"->"gerenciar") is applied
    # first; the gerund is then derived from the pt-BR stem ("gerenciando").
    protected = _fix_progressive(protected)

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
    """Estado for a row whose pt-BR text is ``converted``.

    Returns ``CORRIGIDO`` when there is a real change to export, ``OK`` when
    the text is already valid pt-BR, and ``PENDENTE_PO`` when no automatic
    decision is possible (missing translation / ambiguous warning).
    """
    aviso_low = (aviso or "").lower()

    if "sin traducción" in aviso_low or "sem tradução" in aviso_low:
        return ESTADO_PENDENTE_PO
    if not pt.strip():
        return ESTADO_PENDENTE_PO
    if converted != pt:
        return ESTADO_CORRIGIDO
    if "igual a es" in aviso_low or "igual a en" in aviso_low:
        if _igual_es_en_is_safe(pt):
            return ESTADO_OK
        return ESTADO_PENDENTE_PO
    return ESTADO_OK


#: pt-PT spelling/vocabulary that must never be closed as "already valid
#: pt-BR". Narrow on purpose: "regist\w+", "fat\w+" etc. are identical in
#: pt-BR and would produce false positives.
_PT_PT_MARKER = re.compile(
    r"\b(acrónim\w+|acrónic\w+|géner\w+|morad\w+|montante|montantes|cadeia\w*|"
    r"equipa|equipas|utilizador\w*|ficheir\w+|contacto\w*|receç\w+|secç\w+|"
    r"planeament\w*|detet\w+|gerir|gerid\w+|aced\w+|partilh\w+|prima|percentagem\w*|"
    r"grelha\w*|recolha|demasiad\w*|telemóvel\w*|ecrã\w*|bebé\w*|cómod\w+|"
    r"subscrição|subscrições|anulaç\w+|pretend\w+|guardad\w+|eliminad\w+|"
    r"encontrámos)\b",
    re.IGNORECASE,
)


def _igual_es_en_is_safe(pt: str) -> bool:
    """True when an "Igual a ES/EN" line is already valid pt-BR.

    The warning is a false positive for codes, acronyms, numbers, product
    names, ``lorem ipsum``/``//TODO`` and cognates that are spelled the same
    in Spanish, Portuguese and pt-BR (Total, Menu, Grupo, Cliente...).
    """
    text = (pt or "").strip()
    if not text:
        return False
    norm = text.lower()
    words = set(re.findall(r"\w+", norm, re.UNICODE))
    if words & NOT_OK_IGUAL_ES_EN:
        return False
    if norm in WHITELIST_IGUAL_ES_EN:
        return True
    if _PT_PT_MARKER.search(text):
        return False
    return True


def load_decisoes(path: Path | None = None) -> dict[tuple[str, str], str]:
    """Load pt-BR decisions validated by the review agents.

    File is JSONL, one object per line:
    ``{"clave": ..., "pt_atual": ..., "verdict": "APROVAR"|"CORRIGIR",
    "pt_br": ..., "motivo": ...}``. ``APROVAR`` keeps the deterministic
    suggestion; ``CORRIGIR`` replaces it with the agent's ``pt_br``.

    The key is ``(clave, pt_atual)`` because a clave is NOT unique across the
    workbook (the same key can appear in several sheets with different text);
    matching on the source text keeps a decision scoped to the row it was made
    for. Missing file -> no overrides.
    """
    if path is None:
        path = Path(__file__).resolve().parent / OUT_DECISOES
    decisoes: dict[tuple[str, str], str] = {}
    if not path.exists():
        return decisoes
    with open(path, encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line:
                continue
            record = json.loads(line)
            clave = record.get("clave")
            pt_atual = _clean(record.get("pt_atual"))
            pt_br = _clean(record.get("pt_br"))
            if clave and pt_br and record.get("verdict") in (
                "APROVAR",
                "CORRIGIR",
            ):
                decisoes[(clave, pt_atual)] = pt_br
    return decisoes


#: Filled in ``build()`` from ``decisoes_revisao.jsonl``; empty by default so
#: unit tests of ``converter``/``process_sheet`` stay deterministic.
DECISOES: dict[tuple[str, str], str] = {}


def _clean(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value)


def _prioridade(estado: str, tem_correcao: bool, e_duvida: bool) -> str:
    if estado == ESTADO_OK:
        return ""
    if estado == ESTADO_PENDENTE_PO or e_duvida:
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
            out.at[idx, COL_ESTADO] = ESTADO_OK
            out.at[idx, COL_CORRECAO] = ""
            out.at[idx, COL_PRIORIDADE] = "Fechado"
            continue

        if clave in DUVIDAS_KEYS or CORRECOES.get(clave, "") is None:
            # open question for the PO: never auto-converted, listed in
            # duvidas_PO.xlsx for a human decision
            out.at[idx, COL_ESTADO] = ESTADO_PENDENTE_PO
            out.at[idx, COL_CORRECAO] = ""
            out.at[idx, COL_PRIORIDADE] = _prioridade(
                ESTADO_PENDENTE_PO, False, True
            )
            continue

        dec = DECISOES.get((clave, texto_pt))
        if dec is not None:
            # validated by the review agents: accept their pt-BR verbatim
            converted = dec
        elif clave in CORRECOES:
            converted = CORRECOES[clave]
        else:
            converted = converter(texto_pt)

        estado = _resolve_estado(aviso, texto_pt, converted)

        if estado == ESTADO_OK:
            out.at[idx, COL_CORRECAO] = ""
        else:
            out.at[idx, COL_CORRECAO] = (
                converted if converted != texto_pt else ""
            )
        out.at[idx, COL_ESTADO] = estado
        if dec is not None or clave in CORRECOES:
            # reviewed and closed by the reviewer / review agents
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


def _tem_residuo_pt_pt(texto: str) -> bool:
    """True quando ``texto`` ainda contém pt-PT que não foi convertido.

    Serve a porta do ficheiro de entrega: uma linha ``PENDENTE_PO`` que não
    tenha tradução (ou que conserve pt-PT) não pode aparecer como texto final.
    """
    if not texto:
        return False
    if _PT_PT_MARKER.search(texto):
        return True
    # A rule whose replacement equals what it matched (e.g. transferir->
    # transferir, kept on purpose) changes nothing and is not residue.
    # Case is compared loosely: rules that only re-case a word (Transferir ->
    # transferir) are not residue either.
    return any(
        pat.sub(r, texto).lower() != texto.lower() for pat, r in SUBSTITUICOES
    )


def _construir_final(processed: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    """Constrói o ficheiro de entrega: uma única coluna de texto final em pt-BR.

    Elimina as colunas de texto pt-PT (`Traducción PT actual`, `Texto origen
    (IA)`) e as colunas de trabalho (`Estado`, `Corrección PT`, `Comentario`,
    `Prioridade`), deixando o texto final consolidado numa única coluna
    `Texto final pt-BR`. Assim não há forma de alguém entregar, por engano,
    a coluna pt-PT antiga.

    O texto final é: `Corrección PT` quando existe, senão `Traducción PT actual`
    (linhas OK, já válidas em pt-BR). Fica vazio quando não há texto de origem
    ou quando o texto conserva pt-PT que não foi possível converter — nesses
    casos a decisão pertence ao PO (`duvidas_PO.xlsx`).
    """
    final: dict[str, pd.DataFrame] = {}
    for sheet, df in processed.items():
        out = df.copy()
        pt_col = next((c for c in _PT_SOURCE_COLUMNS if c in out.columns), None)
        if pt_col is None:
            pt_col = next((c for c in PT_COLUMNS if c in out.columns), None)
        pt_original = out[pt_col].map(_clean) if pt_col else ""
        correcao = out[COL_CORRECAO].map(_clean) if COL_CORRECAO in out.columns else ""
        estado = out[COL_ESTADO].map(_clean) if COL_ESTADO in out.columns else ""

        def _final(p: str, c: str, e: str) -> str:
            if c:
                return c
            if not p:
                return ""
            # Only unresolved rows are gated on pt-PT residue; OK/CORRIGIDO text
            # is already validated and must never be blanked by a false positive.
            if e == ESTADO_PENDENTE_PO and _tem_residuo_pt_pt(p):
                return ""
            return p

        out[COL_FINAL] = [
            _final(p, c, e) for p, c, e in zip(pt_original, correcao, estado)
        ]
        out = out.drop(
            columns=[
                c
                for c in (*_PT_SOURCE_COLUMNS, *PT_COLUMNS,
                          COL_ESTADO, COL_CORRECAO, COL_COMENTARIO, COL_PRIORIDADE)
                if c in out.columns
            ]
        )
        final[sheet] = out
    return final


def build(xlsx_in: str | None = None, out_dir: str | None = None) -> int:
    root = Path(__file__).resolve().parent
    entrada = Path(xlsx_in) if xlsx_in else root / XLSX_IN
    if not entrada.is_absolute():
        entrada = root / entrada
    if not entrada.exists():
        print(f"ERRO: ficheiro de entrada não encontrado: {entrada}", file=sys.stderr)
        return 1
    destino = Path(out_dir) if out_dir else entrada.parent
    if not destino.is_absolute():
        destino = root / destino
    destino.mkdir(parents=True, exist_ok=True)

    global DECISOES
    DECISOES = load_decisoes(destino / OUT_DECISOES)

    xl = pd.ExcelFile(entrada)
    review_sheets = _review_sheet_names(xl)
    if not review_sheets:
        print("ERRO: nenhuma aba de revisão encontrada", file=sys.stderr)
        return 1

    processed: dict[str, pd.DataFrame] = {}
    for sheet in review_sheets:
        processed[sheet] = process_sheet(_read_sheet(xl, sheet))

    with pd.ExcelWriter(destino / OUT_COMPLETA) as writer:
        for sheet, df in processed.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)

    final = _construir_final(processed)
    with pd.ExcelWriter(destino / OUT_FINAL) as writer:
        for sheet, df in final.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)

    duvidas = _extract_duvidas(xl, processed)
    duvidas.to_excel(destino / OUT_DUVIDAS, index=False)

    rows = pd.concat(processed.values(), ignore_index=True)
    # only validated corrections go to the backend; OK has nothing to patch and
    # PENDENTE_PO is a glossary decision owned by the PO
    backend = rows[rows[COL_ESTADO] == ESTADO_CORRIGIDO]
    with open(destino / OUT_CSV, "w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh, delimiter=";")
        writer.writerow([KEY_COLUMNS[0], COL_ESTADO, COL_CORRECAO, COL_COMENTARIO])
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
                ]
            )

    contagem = rows[COL_ESTADO].value_counts().to_dict()
    print("Contagem final por estado:")
    for estado in (ESTADO_OK, ESTADO_CORRIGIDO, ESTADO_PENDENTE_PO):
        print(f"  {estado}: {contagem.get(estado, 0)}")
    print(f"Total de linhas: {len(rows)}")
    print(f"Substituições distintas: {len(SUBSTITUICOES)}")
    print(f"Dúvidas (não convertidas): {len(duvidas)}")
    print(f"Decisões dos agentes aplicadas: {len(DECISOES)}")
    return 0


def _cli() -> int:
    parser = argparse.ArgumentParser(
        description="Converte pt-PT → pt-BR num ficheiro de revisão Ulyses."
    )
    parser.add_argument(
        "entrada",
        nargs="?",
        help=f"XLSX de entrada (por omissão: {XLSX_IN}).",
    )
    parser.add_argument(
        "-o",
        "--out-dir",
        help="Diretório de saída (por omissão: o do ficheiro de entrada).",
    )
    args = parser.parse_args()
    return build(args.entrada, args.out_dir)


if __name__ == "__main__":
    raise SystemExit(_cli())
