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
OUT_PREENCHIDO = "Revision_traducciones_PT_pt-BR.xlsx"
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

#: Final tab names (pt-BR) shown in the delivered workbooks; the input file
#: keeps the legacy ES/PT labels.
SHEET_NAMES_PT_BR = {
    "Frontend (pantallas)": "Frontend (telas)",
    "Entidades de sistema": "Entidades do sistema",
    "Mensajes e informes": "Mensagens e relatórios",
}

#: Estado values. ``OK`` = already valid pt-BR (nothing to patch);
#: ``CORRIGIDO`` = validated pt-BR suggestion, exported to the backend;
#: ``PENDENTE_PO`` = glossary decision owned by the PO, never auto-converted.
#: The labels match the source workbook dropdown (``OK,Corregir``) so the
#: delivered file can be filtered as-is.
ESTADO_OK = "OK"
ESTADO_CORRIGIDO = "Corregir"
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
#: Every question was closed by the PO sign-off (see ``CORRECOES_PO``), so the
#: maps below stay empty; they are kept as hooks for future rounds.
DUVIDAS_GRUPOS: dict[str, list[str]] = {}

DUVIDAS_KEYS: list[str] = []

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
        "Confirme abaixo para tornar as alterações visíveis. Observe que "
        "as alterações não serão salvas até clicar em Salvar na tela principal."
        "\n\nBaixe o modelo Excel caso ainda não o tenha feito."
    ),
    "before.checkout.message.intro": (
        "Esperamos que tenha gostado da sua estadia em {0}. Para tornar a sua "
        "saída mais confortável, sem contato e sem passar pela recepção amanhã, "
        "você já pode fazer o check-out online."
    ),
    "apps.login.section.passwordReset.descriptionSystem": (
        "Por motivos de segurança, a senha do Ulyses Suite tem de ser atualizada "
        "a cada 90 dias. Lembre-se de não reutilizar nenhuma das últimas 4 senhas "
        "que utilizou para fazer login no Ulyses Suite."
    ),
    "apps.integration.publicApiRest.refreshEventEntity.message": (
        "Aplique esta ação se deseja reenviar informações sobre a entidade de "
        "evento selecionada.\nObserve que as alterações não serão "
        "atualizadas imediatamente; por isso, recomendamos que consulte a aba Fila "
        "de eventos para acompanhar a evolução."
    ),
    "apps.integration.publicApiRest.removeEventEntity.message": (
        "Aplique esta ação se deseja excluir informações sobre a entidade de "
        "evento selecionada.\nObserve que as alterações não serão "
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
}

#: PO sign-off (2026-10): every previously open question is now a binding
#: glossary decision, so the ``None`` markers above are superseded here. These
#: entries also close rows that the source flags as "Sin traducción PT" /
#: "Igual a ES/EN": an explicit override is the decision and wins over the
#: warning.
CORRECOES_PO: dict[str, str] = {
    # A. glossary decisions
    "apps.aiAssistant.label.hotel": "Propriedade",
    "apps.profileCommissionTransaction.gridAction.offset": "Compensar",
    "core.label.billing": "Faturamento",
    "core.label.chain": "Rede",
    "shared.forecast.nonDeductedAllotment": "Allotment não deduzido",
    "apps.chainSpaceLockType.radio.onHold": (
        "OnHold (Ative esta opção para bloquear a acomodação)"
    ),
    "apps.chainSpaceLockType.radio.outOfOrder": (
        "OutOfOrder (Ative esta opção para permitir retirar a acomodação "
        "do inventário)"
    ),
    "apps.chainSpaceLockType.radio.outOfService": (
        "OutOfService (Ative esta opção para permitir retirar a acomodação "
        "da disponibilidade)"
    ),
    "PropertyClassType t WHERE t.Code = '9'": "Pensão",
    # B. translations missing in the source workbook
    "apps.billing.folio.refund.paymentMethod": "Método de pagamento do reembolso",
    "apps.import.importJob.relatedMap.accountReceivableStatementSendType": (
        "Forma de envio do extrato de contas a receber"
    ),
    "apps.import.importJob.relatedMap.allotment": "Allotment",
    "apps.import.importJob.relatedMap.cardStatusType": "Status do cartão",
    "apps.import.importJob.relatedMap.chainContactCategory": "Categoria de contato",
    "apps.import.importJob.relatedMap.chainLoyalty": (
        "Fidelização - Programa de afiliação"
    ),
    "apps.import.importJob.relatedMap.chainProductGroup": "Grupo de produtos",
    "apps.import.importJob.relatedMap.chainProductSubGroup": "Subgrupo de produtos",
    "apps.import.importJob.relatedMap.chainVIPLevel": "Nível VIP",
    "apps.import.importJob.relatedMap.communicationLocationType": "Tipo de contato",
    "apps.import.importJob.relatedMap.depositStatusType": "Status do depósito",
    "apps.import.importJob.relatedMap.emailAddressType": "Tipo de e-mail",
    "apps.import.importJob.relatedMap.familyRelationship": "Parentesco",
    "apps.import.importJob.relatedMap.genderType": "Gênero",
    "apps.import.importJob.relatedMap.municipality": "Município",
    "apps.import.importJob.relatedMap.phoneTechType": (
        "Tipo de tecnologia telefônica"
    ),
    "apps.import.importJob.relatedMap.productDeposit": "Produto - depósito",
    "apps.import.importJob.relatedMap.productPayment": "Produto - pagamento",
    "apps.import.importJob.relatedMap.propertyRoomFeatureType": (
        "Característica de acomodação"
    ),
    "apps.import.importJob.relatedMap.propertyRoomType": "Tipologia",
    "apps.import.importJob.relatedMap.reservationPriceModelType": (
        "Tipo de modelo de preços de reserva"
    ),
    "apps.import.importJob.relatedMap.reservationStatusType": (
        "Tipo de status da estadia"
    ),
    "apps.import.importJob.relatedMap.spaceStatusType": "Tipo de status dos espaços",
    "apps.import.importJob.relatedMap.stateProv": "Estado/província",
    "apps.import.importJob.relatedMap.webSiteLocationType": (
        "Tipo de localização do site"
    ),
    "apps.management.guestLedger.grid.header.totalFunctionRoomAmountAfterTax": (
        "Salão de eventos"
    ),
    "apps.propertyCommunicationEmailTypeConfiguration.entity.reservationGroupContract.info": (
        "Ative esta opção se você quiser que seus usuários possam enviar "
        "comunicações de contratos de reservas de grupo aos seus clientes. "
        "Se quiser personalizar a comunicação, acesse Modelos de e-mail na "
        "Configuração da propriedade."
    ),
    "apps.propertyCommunicationEmailTypeConfiguration.entity.reservationGroupContract.title": (
        "Contrato de reserva de grupo"
    ),
    "apps.reservationGroup.action.printDocument": "Imprimir documento",
    "apps.reservationGroup.action.sendConfirmation": "Enviar confirmação",
    "apps.reservationGroup.action.sendContract": "Enviar contrato",
    "apps.reservationGroup.action.sendProposal": "Enviar proposta",
    "core.label.billingInstruction": "Instruções de faturamento",
    "core.object.entity.propertyCheckInAutomationScheduler": "Check-in automático",
    "core.object.entity.reservationGroupContract": "Contrato de reserva de grupo",
    "core.userMessage.server.ratediscountnotallowedforpricemodeltype.default": (
        "Não é permitido aplicar desconto ao modelo de preços desta tarifa"
    ),
    "core.userMessage.server.rateoverwritenotallowedforpricemodeltype.default": (
        "Não é permitido sobrescrever o preço no modelo de preços desta tarifa"
    ),
    "shared.aggregateGrid.empty.chainProductTaxIncluded.noIncludedProductTax": (
        "Adicione os produtos que podem ser utilizados neste produto."
    ),
    "shared.aggregateGrid.empty.product.noProductTaxIncluded": (
        "Adicione os impostos aplicáveis a este produto"
    ),
    "shared.aggregateGrid.empty.propertyKioskDevic ePropertyServiceConsumptionControl.noPropertyServiceConsumptionControl": (
        "Adicione controles de serviços para este dispositivo; caso não "
        "adicione nenhum, todos serão exibidos"
    ),
    "shared.assistant.propertyAutomatedCheckinScheduler.step0": "Resumo",
    "shared.assistant.propertyAutomatedCheckinScheduler.step1": "Filtros",
    "shared.assistant.propertyAutomatedCheckinScheduler.step2": "Ação",
    "shared.assistant.propertyAutomatedCheckinScheduler.step3": "Revisão",
    "shared.emailTemplate.section.holderBillToGroupContract": (
        "Detalhe do titular e faturamento"
    ),
    "shared.emailTemplate.section.htmlReservationGroupContract": (
        "Contrato de reserva de grupo HTML"
    ),
    "shared.emailTemplate.section.propertyBillingInformation": (
        "Dados de faturamento da propriedade"
    ),
    "shared.emailTemplate.section.textReservationGroupContract": (
        "Texto do contrato de reserva de grupo"
    ),
    "shared.grid.header.overwriteAmountAfterTax": "Sobrescrita",
    # C. empty in the source workbook, closed by the PO (BR fiscal wording)
    "shared.radioButton.propertyEmailSchedulerEventINVCAMD.message": (
        "Fatura alterada"
    ),
    "shared.radioButton.propertyEmailSchedulerEventINVCISS.message": (
        "Fatura emitida"
    ),
    "shared.radioButton.propertyEmailSchedulerEventINVCVOID.message": (
        "Fatura anulada"
    ),
    "txt.space": "Mensagens e relatórios",
}

CORRECOES.update(CORRECOES_PO)

#: Correções descobertas na auditoria BR (2026-10): pt-PT que passou intacto,
#: sobretudo "Descargar/Download" traduzido como "Transferir/Transfira" (o pt-BR
#: usa "Baixar"), além de "eliminá-lo"->"excluí-lo" e "A aguardar"->"Aguardando".
#: Têm precedência sobre as regras genéricas; não colidem com decisões dos agentes.
CORRECOES_BR: dict[str, str] = {
    "apps.accountReceivableTransaction.gridAction.download": "Baixar",
    "apps.endOfDay.process.emergencyReports.description": (
        "Baixe os relatórios de emergência e, caso ocorra uma falha generalizada, estará preparado para gerenciar a sua propriedade temporariamente sem o Ulyses Suite."
    ),
    "apps.integration.barcelo.downloadInvoice": "Baixar fatura",
    "apps.integration.bavel.downloadInvoice": "Baixar fatura",
    "apps.integration.description.a3": (
        "Baixe arquivos e carregue-os no sistema de contabilidade A3Con. Poupe tempo e analise dados com esta solução da Wolters Kluwer."
    ),
    "apps.integration.description.gsp": (
        "O Ulyses Guest Services Portal permite que os seus hóspedes façam o check-in online, consultem a sua conta, baixem faturas e muito mais. Poupe tempo, aumente a distância e proteja a sua equipe."
    ),
    "apps.integration.description.sage200": (
        "O Sage 200 é um conjunto de produtos de contabilidade e gestão desenvolvido pelo Sage Group, destinado a médias empresas. Baixe os arquivos necessários e carregue-os no seu sistema de contabilidade."
    ),
    "apps.integration.ertzaintza.button.downloadParteViajerosXML": "Baixar XML de partes de viajantes",
    "apps.integration.ertzaintza.button.downloadReservationXML": "Baixar XML de reservas",
    "apps.integration.fatturaPA.downloadInvoice": "Baixar faturas",
    "apps.rate.importExcelModal.instructions": (
        "Carregue o documento preenchido com todas as condições do contrato atualizadas (Preços, Encerramentos, CTA, CTD, Estadia máx./mín.). Baixe o modelo Excel caso ainda não o tenha feito."
    ),
    "apps.report.options.downloadNow": "Baixar agora",
    "apps.reservation.reservationRoomStaySidebar.await": "Aguardando localizador",
    "apps.reservationGroup.event.await": "Aguardando localizador",
    "apps.reservationGroupInquiry.await": "Aguardando localizador",
    "apps.rooming.importExcel": (
        "Baixe o modelo abaixo e preencha-o com toda a informação necessária para criar todas as acomodações. Pode compartilhar este arquivo com outras empresas, pelo que deve garantir que o carrega quando o rooming estiver confirmado."
    ),
    "apps.systemSettings.import.downloadTemplate": (
        "Baixe o seguinte arquivo, preencha os seus dados e carregue-o para importar dados para o UlysesCloud"
    ),
    "core.button.downloadExcel": "Baixar Excel",
    "core.button.downloadExcelWithAllMappings": "Baixar todos os mapeamentos Excel",
    "core.button.downloadInvoice": "Baixar fatura",
    "core.label.waitingFor": "Aguardando desde",
    "core.userMessage.server.paynopain.credorax.pending_3ds_response": "Aguardando processamento da resposta em 3DS",
    "core.userMessage.server.paynopain.credorax.pending_confirmation": "Aguardando confirmação",
    "core.userMessage.server.paynopain.credorax.pending_processor_response": "Resposta aguardando processamento",
    "shared.grid.noViewAllDataInGridReport": (
        "O resultado tem mais de 5000 linhas, altere os filtros aplicados ou baixe o relatório em PDF, Excel ou CSV."
    ),
    "shared.modal.bulkAssign.importExcel.instructions": (
        "Carregue o documento preenchido com todas as estadias atualizadas. Baixe o modelo Excel caso ainda não o tenha feito."
    ),
    "shared.modal.importAccountingMappingExcel.importExcel": (
        "Baixe o seguinte arquivo, defina as suas acomodações com os respetivos tipos de acomodação e carregue-o para criar a estrutura de acomodações da sua propriedade."
    ),
    "shared.modal.importItemExcel.importExcel": (
        "Baixe o seguinte arquivo, defina os seus produtos com os respetivos tipos de itens e carregue-o para criar a estrutura de itens da sua propriedade."
    ),
    "shared.modal.importMeterExcel.importExcel": (
        "Baixe o arquivo, especifique os contadores com os respetivos tipos e carregue-o para criar a estrutura de contadores do seu estabelecimento."
    ),
    "shared.modal.importProductExcel.importExcel": (
        "Baixe o seguinte arquivo, defina os seus produtos com os respetivos tipos de produto e carregue-o para criar a estrutura de produtos da sua propriedade."
    ),
    "shared.modal.importRoomExcel.importExcel": (
        "Baixe o seguinte arquivo, defina as suas acomodações com os respetivos tipos de acomodação e carregue-o para criar a estrutura de acomodações da sua propriedade."
    ),
    "shared.modal.importTranslateExcel.importExcel": (
        "Baixe o seguinte arquivo, defina as suas acomodações com os respetivos tipos de acomodação e carregue-o para criar a estrutura de acomodações da sua propriedade."
    ),
    "shared.modal.information.createCutOff.message": (
        "Vai executar o processo de Cut-Off. Após este processo, as configurações e o bloqueio de acomodação ficarão bloqueadas. Pode revertê-lo excluindo o Cut-Off."
    ),
    "shared.modal.information.deleteMeterNotAllowed.message": (
        "Este contador tem leituras adicionadas e não é possível excluá-lo"
    ),
    "shared.modal.information.genericAccountingAlreadyExported.title": "Período baixado",
    "shared.modal.information.siteminderImportReservation.message": (
        "Tem certeza de que deseja prosseguir com esta ação?\n\nEste processo irá pedir ao Channel Manager que volte a baixar todas as reservas, tanto atuais como futuras. Como resultado, todas as reservas geradas pelo Channel Manager serão baixadas novamente e as reservas existentes serão atualizadas em conformidade.\n\nObserve que este processo é irreversível."
    ),
    "shared.modal.propagateVehicle.deleteMessage": (
        "Este veículo é utilizado em {{count}} estadias atuais ou futuras. Deseja também excluí-lo dessas estadias?"
    ),
    "shared.modal.rooming.downloadExcel.roomingListType1": "Baixar atual",
    "shared.modal.rooming.downloadExcel.roomingListType2": "Baixar exemplo",
    "shared.modal.rooming.downloadExcel.roomingListType3": "Baixar bloqueio de acomodações",
    "shared.modal.rooming.downloadExcel.roomingListType4": "Baixar bloqueio de acomodações alargado",
    "shared.modal.rooming.importExcel": (
        "Baixe o modelo abaixo e preencha-o com todas as informações necessárias para criar todas as acomodações. Pode compartilhar este arquivo com outras empresas, mas certifique-se de que o carrega quando a rooming list estiver confirmada."
    ),
    "shared.switch.allowDownloadInvoice.message": (
        "Ative esta opção para permitir que os usuários baixem as faturas relacionadas com as suas reservas"
    ),
    "shared.switch.allowInHouseDownloadInvoice.message": (
        "Ative esta opção para permitir que os usuários baixem as faturas relacionadas com as suas reservas"
    ),
    "shared.switch.generateRoomingListByBlock.message": (
        "Ative esta opção se quiser que o arquivo a baixar tenha as rooming lists separadas por data e blocos de quantidade, em vez do mínimo possível de rooming lists"
    ),
    "shared.switch.profileMappingAllowAddDeposit.message": (
        "Ative esta opção para permitir a criação de depósitos quando são baixados na reserva."
    ),
}

CORRECOES.update(CORRECOES_BR)

#: Explanatory notes for the ``Comentario`` column. Only rows that are
#: polemic, ambiguous or need a reason get a note; everything else stays
#: empty on purpose (the correction itself is self-explanatory).
COMENTARIOS: dict[str, str] = {
    # --- glossary / PO decisions (B2B jargon, regional variants) -----------
    "shared.forecast.nonDeductedAllotment": (
        "Allotment mantido (decisão de glossário B2B): é termo corrente em "
        "contratos com operadores estrangeiros. Não traduzir por cota."
    ),
    "apps.import.importJob.relatedMap.allotment": (
        "Allotment mantido por consistência com a decisão de glossário B2B."
    ),
    "apps.integration.ulysesSync.mapping.allotment": (
        "Allotment mantido por consistência com a decisão de glossário B2B."
    ),
    "apps.importJob.category.importAllotmentMapList": (
        "Allotment mantido por consistência com a decisão de glossário B2B."
    ),
    "PropertyClassType t WHERE t.Code = '9'": (
        "Pensão mantido: em pt-BR hoteleiro é a forma usada para o regime de "
        "refeições (pensão completa / meia pensão). Não confundir com o "
        "sentido pt-PT de casa de hóspedes."
    ),
    "apps.aiAssistant.label.hotel": (
        "Propriedade (e não Hotel): o assistente cobre hotéis e outros tipos "
        "de propriedade, seguindo o EN \"Property\". Só o ES usa \"Hotel\"."
    ),
    "core.label.billing": (
        "Faturamento: em pt-BR a faturação é \"faturamento\" (o pt-PT "
        "\"facturação/faturação\" é forma lusa)."
    ),
    "core.label.chain": (
        "Rede: pt-BR usa \"rede\" para cadeia hoteleira; \"cadeia\" é "
        "interpretado como cadeia de caracteres/produto."
    ),
    "core.label.billingInstruction": "Coerente com core.label.billing = Faturamento.",
    "shared.emailTemplate.section.propertyBillingInformation": (
        "Coerente com a decisão core.label.billing = Faturamento."
    ),
    "apps.profileCommissionTransaction.gridAction.offset": (
        "Compensar: ação contabilística de compensar valores; o pt-PT "
        "\"compensar/equilibrar\" aqui é ambíguo."
    ),
    # --- BR fiscal wording: anulada != cancelada --------------------------
    "shared.radioButton.propertyEmailSchedulerEventINVCVOID.message": (
        "Fatura anulada (VOID) mantido literal: em BR fiscal \"anulada\" "
        "(invalidação/nulidade) é distinto de \"cancelada\". A regra genérica "
        "anulada->cancelada NÃO se aplica aqui."
    ),
    "shared.radioButton.propertyEmailSchedulerEventINVCAMD.message": (
        "Fatura alterada: evento de alteração de fatura (INVCAMD)."
    ),
    "shared.radioButton.propertyEmailSchedulerEventINVCISS.message": (
        "Fatura emitida: evento de emissão de fatura (INVCISS)."
    ),
    # --- technical tokens kept verbatim -----------------------------------
    "apps.chainSpaceLockType.radio.onHold": (
        "OnHold mantido: valor técnico do sistema (estado de bloqueio). "
        "Traduzido apenas o texto entre parênteses."
    ),
    "apps.chainSpaceLockType.radio.outOfOrder": (
        "OutOfOrder mantido: valor técnico do sistema (fora de ordem). "
        "Traduzido apenas o texto entre parênteses."
    ),
    "apps.chainSpaceLockType.radio.outOfService": (
        "OutOfService mantido: valor técnico do sistema (fora de serviço). "
        "Traduzido apenas o texto entre parênteses."
    ),
    # --- surgical fixes from the BR source ---------------------------------
    "card.hojaRegistro": (
        "Ficha de registro de hóspedes (FNRH): designação oficial brasileira "
        "(Polícia Federal/Embratur). Não traduzir literalmente \"folha de "
        "registo\"."
    ),
    "core.button.settlement": (
        "Cobranças automáticas: o ES \"Cargos automáticos\" refere a cobrança "
        "automática, não a liquidação de conta."
    ),
    "invoice.direction.label": (
        "Endereço: pt-BR usa \"endereço\" (pt-PT \"morada\")."
    ),
    "core.label.taxExcludedShort": (
        "Imposto excl.: no Brasil não existe IVA; usa-se \"imposto\"."
    ),
    "core.label.taxIncludedShort": (
        "Imposto incl.: no Brasil não existe IVA; usa-se \"imposto\"."
    ),
    "core.label.amountPaidOut": (
        "Valor retirado: segue o ES \"cantidad retirada\"; \"montante\" é "
        "pt-PT."
    ),
    "core.label.totalAmountEntered": (
        "Valor total informado: pt-BR prefere \"informado\" a \"introduzido\"."
    ),
    "core.label.amountOtherBreakdownAfterTax": (
        "Outros valores: forma idiomática em pt-BR; evita \"montante\"."
    ),
    # --- "to download" -> baixar (não confundir com transferir) -----------
    "core.button.download": (
        "Baixar: em pt-BR \"download\" é \"baixar\". Regra aplicada só aos "
        "botões de download, nunca às transferências de conta/folio."
    ),
    "core.button.downloadExcel": (
        "Baixar: \"Descargar\" é \"baixar\" em pt-BR, não \"transferir\" "
        "(transferir = mover entre contas/folios)."
    ),
    "core.button.downloadInvoice": (
        "Baixar fatura: \"Download invoice\" é descarregar o PDF, não "
        "transferir valores."
    ),
    "shared.modal.rooming.downloadExcel.roomingListType1": (
        "Baixar atual: botão de descarregar a rooming list; ES \"Descargar\"."
    ),
    "shared.modal.rooming.downloadExcel.roomingListType2": (
        "Baixar exemplo: botão de descarregar o modelo; ES \"Descargar\"."
    ),
    "shared.modal.rooming.downloadExcel.roomingListType3": (
        "Baixar bloqueio de acomodações: ES \"Descargar roomblock\"; é "
        "descarregar o ficheiro, não transferir."
    ),
    "shared.modal.rooming.downloadExcel.roomingListType4": (
        "Baixar bloqueio de acomodações alargado: ES \"Descargar roomblock "
        "extendido\"; é descarregar o ficheiro."
    ),
    "apps.report.options.downloadNow": (
        "Baixar agora: opção de descarregar o relatório imediatamente."
    ),
    "shared.switch.profileMappingAllowAddDeposit.message": (
        "Baixados: os depósitos são descarregados (ES \"descargados\"), não "
        "transferidos entre contas."
    ),
    "shared.modal.information.genericAccountingAlreadyExported.title": (
        "Período baixado: ES \"Periodo descargado\"; refere a exportação/"
        "descarga do período contabilístico."
    ),
    "shared.switch.generateRoomingListByBlock.message": (
        "Arquivo a baixar: ES \"que se descargue el archivo\"; é descarregar, "
        "não transferir."
    ),
    "shared.switch.allowDownloadInvoice.message": (
        "Baixem: ES \"la descarga de facturas\"; os clientes descarregam as "
        "faturas."
    ),
    "shared.switch.allowInHouseDownloadInvoice.message": (
        "Baixem: ES \"la descarga de facturas\"; os clientes descarregam as "
        "faturas."
    ),
    "shared.modal.information.siteminderImportReservation.message": (
        "Baixar novamente: ES \"vuelva a descargar\"; o Channel Manager volta "
        "a descarregar as reservas (não as transfere)."
    ),
    "shared.grid.noViewAllDataInGridReport": (
        "Baixe o relatório: ES \"descargue el report\"; acima de 5000 linhas "
        "é preciso descarregar."
    ),
    "apps.rate.importExcelModal.instructions": (
        "Baixe o modelo: ES \"descargar la plantilla\"."
    ),
    "shared.modal.bulkAssign.importExcel.instructions": (
        "Baixe o modelo: ES \"Descarga el modelo Excel\"."
    ),
    "apps.integration.description.a3": (
        "Baixe arquivos: ES \"Descarga asientos contables y expórtalos\"; "
        "refere a exportação de ficheiros para o sistema A3Con."
    ),
    "apps.integration.description.sage200": (
        "Baixe os arquivos: ES \"Descarga los ficheros necesarios\"."
    ),
    "apps.integration.description.gsp": (
        "Baixem faturas: ES \"descargar sus facturas\"; os hóspedes "
        "descarregam as suas faturas."
    ),
    "apps.endOfDay.process.emergencyReports.description": (
        "Baixe os relatórios: ES \"Descarga los informes de emergencia\"; é "
        "descarregar para uso offline, não transferir."
    ),
    "apps.rooming.importExcel": (
        "Baixe o modelo: ES \"Descargue la plantilla a continuación\"."
    ),
    "shared.modal.rooming.importExcel": (
        "Baixe o modelo: ES \"Descarga la plantilla y rellénala\"."
    ),
    "apps.systemSettings.import.downloadTemplate": (
        "Baixe o ficheiro: ES \"Descarga el siguiente fichero\"; descarregar "
        "o modelo, não transferir."
    ),
    "shared.modal.importItemExcel.importExcel": (
        "Baixe o arquivo: ES \"Descarga el fichero\"; descarregar o modelo."
    ),
    "shared.modal.importMeterExcel.importExcel": (
        "Baixe o arquivo: ES \"Descargue el archivo\"; descarregar o modelo."
    ),
    "shared.modal.importProductExcel.importExcel": (
        "Baixe o arquivo: ES \"Descarga el fichero\"; descarregar o modelo."
    ),
    "shared.modal.importRoomExcel.importExcel": (
        "Baixe o arquivo: ES \"Descarga el fichero\"; descarregar o modelo."
    ),
    "shared.modal.importTranslateExcel.importExcel": (
        "Baixe o arquivo: ES \"Descargue el siguiente archivo\"; descarregar "
        "o modelo."
    ),
    "shared.modal.importAccountingMappingExcel.importExcel": (
        "Baixe o arquivo: ES \"Descarga el siguiente archivo\"; descarregar o "
        "modelo."
    ),
    "apps.accountReceivableTransaction.gridAction.download": (
        "Baixar: ação de descarregar o extrato (ES \"Descargar\"). O "
        "genérico \"Transferir\" era erro; existe outra chave .transfer para "
        "a transferência real entre contas."
    ),
    "apps.integration.barcelo.downloadInvoice": (
        "Baixar fatura: ES \"Descargar factura\"; é descarregar o PDF."
    ),
    "apps.integration.bavel.downloadInvoice": (
        "Baixar fatura: ES \"Descargar factura\"; é descarregar o PDF."
    ),
    "apps.integration.fatturaPA.downloadInvoice": (
        "Baixar faturas: ES \"Descargar facturas\"; é descarregar os PDF."
    ),
    "apps.integration.ertzaintza.button.downloadParteViajerosXML": (
        "Baixar XML: ES \"Descarga parte viajeros XML\"; é descarregar o "
        "ficheiro."
    ),
    "apps.integration.ertzaintza.button.downloadReservationXML": (
        "Baixar XML: ES \"Descarga reservation XML\"; é descarregar o "
        "ficheiro."
    ),
    "core.button.downloadExcelWithAllMappings": (
        "Baixar todos os mapeamentos Excel: ES \"Descargar Excel de mapeos\"; "
        "é descarregar o ficheiro."
    ),
    "apps.reservation.reservationRoomStaySidebar.await": (
        "Aguardando: em pt-BR usa-se o gerúndio (\"Aguardando\"), não \"A "
        "aguardar\"; ES \"Esperando\"."
    ),
    "apps.reservationGroup.event.await": (
        "Aguardando: em pt-BR usa-se o gerúndio; ES \"Esperando\"."
    ),
    "apps.reservationGroupInquiry.await": (
        "Aguardando: em pt-BR usa-se o gerúndio; ES \"Esperando\"."
    ),
    "core.label.waitingFor": (
        "Aguardando: em pt-BR usa-se o gerúndio; ES \"Esperando desde\"."
    ),
    "core.userMessage.server.paynopain.credorax.pending_3ds_response": (
        "Aguardando: gerúndio pt-BR para \"Pendiente de\"."
    ),
    "core.userMessage.server.paynopain.credorax.pending_confirmation": (
        "Aguardando: gerúndio pt-BR para \"Pendiente de\"."
    ),
    "core.userMessage.server.paynopain.credorax.pending_processor_response": (
        "Aguardando: gerúndio pt-BR para \"pending\"."
    ),
    "shared.modal.information.deleteMeterNotAllowed.message": (
        "Excluí-lo: pt-BR usa \"excluir\" para delete; \"eliminar\" é pt-PT."
    ),
    "shared.modal.propagateVehicle.deleteMessage": (
        "Excluí-lo: pt-BR usa \"excluir\" para delete; \"eliminar\" é pt-PT."
    ),
    "shared.modal.information.createCutOff.message": (
        "Excluindo: pt-BR usa \"excluir\" para delete; \"eliminar\" é pt-PT."
    ),
    "core.navbar.user.logs": (
        "Baixar logs: verbo \"baixar\" (pt-BR) em vez de \"transferir\"."
    ),
    "shared.modal.bulkDownloadFile.title": (
        "Baixando arquivos em massa: gerúndio pt-BR; \"a transferir\" seria "
        "pt-PT."
    ),
    "apps.report.options.information": (
        "Baixar/baixe: mantém o sentido de descarregar o relatório, não de "
        "transferir valores."
    ),
    # --- pt-PT residue in the suggestion itself ---------------------------
    "apps.budget.importExcelModal.instructions": (
        "Sugestão continha pt-PT residual (\"tenha em conta\") no próprio "
        "texto de destino; corrigido para \"observe\"/\"note\"."
    ),
    "apps.integration.publicApiRest.refreshEventEntity.message": (
        "Sugestão continha pt-PT residual; corrigido para pt-BR (observe/"
        "atualizadas)."
    ),
    "apps.integration.publicApiRest.removeEventEntity.message": (
        "Sugestão continha pt-PT residual; corrigido para pt-BR (observe/"
        "atualizadas)."
    ),
    # --- form of address: pt-PT "tu/você" vs pt-BR "você" -----------------
    "apps.propertyCommunicationEmailTypeConfiguration.entity.reservationGroupContract.info": (
        "Tratamento por \"você\" (pt-BR) em vez de \"tu\"; mantidos os termos "
        "de UI (Modelos de e-mail, Configuração da propriedade)."
    ),
    "before.checkout.message.intro": (
        "Tratamento por \"você\" (pt-BR) e \"recepção\"; placeholders {0} e "
        "HTML preservados."
    ),
    "apps.login.section.passwordReset.descriptionSystem": (
        "Senha (pt-BR) em vez de palavra-passe; tratamento por \"você\"."
    ),
    # --- plural/spelling normalisation in this batch -----------------------
    "shared.grid.header.overwriteAmountAfterTax": (
        "Sobrescrita: pt-BR prefere \"sobrescrita\" a \"sobreposição\"."
    ),
    "apps.management.guestLedger.grid.header.totalFunctionRoomAmountAfterTax": (
        "Salão de eventos: \"function room\" é o salão para eventos; evita o "
        "literal \"sala de funções\"."
    ),
    "apps.reservationGroup.action.printDocument": (
        "Imprimir documento: \"print\" é \"imprimir\"; \"imprimir documento\" "
        "já é pt-BR válido."
    ),
    "txt.space": (
        "Mensagens e relatórios: identificador interno da aba, preenchido "
        "para não ficar vazio."
    ),
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
    # 10. lacunas encontradas na auditoria cruzada (relatório do painel nativo)
    # "folha de registo" must run before the "registo"->"registro" rule, so both
    # the pt-PT and the already-converted spelling are mapped.
    _rule("folha de registo", "ficha de registro"),
    _rule("folha de registro", "ficha de registro"),
    _rule("faturar a", "faturar para"),
    _rule("NIF", "CPF/CNPJ"),
    _rule("código postal", "CEP"),
    _rule("códigos postais", "CEPs"),
    _rule("hora de fim", "horário de término"),
    _rule("hora de início", "horário de início"),
    _rule("informação bancária", "informações bancárias"),
    _rule("informação fiscal", "informações fiscais"),
    _rule("iniciar sessão", "fazer login"),
    _rule("iniciar a sessão", "fazer o login"),
    _rule("terminar sessão", "sair"),
    _rule("encerrar sessão", "sair"),
    _rule("menu de bar", "cardápio do bar"),
    _rule("menu de restaurante", "cardápio do restaurante"),
    _rule("lavandaria", "lavanderia"),
    _rule("artigos", "itens"),
    _rule("artigo", "item"),
    _rule("desportivos", "esportivos"),
    _rule("desportivas", "esportivas"),
    _rule("desportivo", "esportivo"),
    _rule("desportiva", "esportiva"),
    _rule("descarregue", "baixe"),
    _rule("descarregar", "baixar"),
    _rule("partidas", "saídas"),
    _rule("partida", "saída"),
    # "alojamento(s)" -> "acomodação(ões)" changes gender (masc -> fem), so the
    # determiner and the adjective must change with it. Every phrase below was
    # measured in the workbook; the bare word is converted last. Longer
    # determiner phrases must precede the shorter ones that they contain
    # ("todos os alojamentos" before "os alojamentos").
    _rule("todo o alojamento", "toda a acomodação"),
    _rule("todos os alojamentos", "todas as acomodações"),
    _rule("o alojamento", "a acomodação"),
    _rule("os alojamentos", "as acomodações"),
    _rule("um alojamento", "uma acomodação"),
    _rule("uns alojamentos", "umas acomodações"),
    _rule("do alojamento", "da acomodação"),
    _rule("dos alojamentos", "das acomodações"),
    _rule("no alojamento", "na acomodação"),
    _rule("nos alojamentos", "nas acomodações"),
    _rule("ao alojamento", "à acomodação"),
    _rule("aos alojamentos", "às acomodações"),
    _rule("deste alojamento", "desta acomodação"),
    _rule("destes alojamentos", "destas acomodações"),
    _rule("neste alojamento", "nesta acomodação"),
    _rule("nestes alojamentos", "nestas acomodações"),
    _rule("este alojamento", "esta acomodação"),
    _rule("estes alojamentos", "estas acomodações"),
    _rule("esse alojamento", "essa acomodação"),
    _rule("esses alojamentos", "essas acomodações"),
    _rule("aquele alojamento", "aquela acomodação"),
    _rule("aqueles alojamentos", "aquelas acomodações"),
    _rule("seu alojamento", "sua acomodação"),
    _rule("seus alojamentos", "suas acomodações"),
    _rule("novo alojamento", "nova acomodação"),
    _rule("novos alojamentos", "novas acomodações"),
    _rule("último alojamento", "última acomodação"),
    _rule("últimos alojamentos", "últimas acomodações"),
    _rule("determinado alojamento", "determinada acomodação"),
    _rule("determinados alojamentos", "determinadas acomodações"),
    _rule("respetivo alojamento", "respetiva acomodação"),
    _rule("respetivos alojamentos", "respetivas acomodações"),
    _rule("nenhum alojamento", "nenhuma acomodação"),
    _rule("alojamento atribuído", "acomodação atribuída"),
    _rule("alojamentos atribuídos", "acomodações atribuídas"),
    _rule("alojamento selecionado", "acomodação selecionada"),
    _rule("alojamentos selecionados", "acomodações selecionadas"),
    _rule("alojamento físico", "acomodação física"),
    _rule("alojamento combinado", "acomodação combinada"),
    _rule("alojamento reservado", "acomodação reservada"),
    _rule("alojamento provisório", "acomodação provisória"),
    _rule("alojamentos provisórios", "acomodações provisórias"),
    _rule("alojamento turístico", "acomodação turística"),
    _rule("alojamentos turísticos", "acomodações turísticas"),
    _rule("alojamento consumido", "acomodação consumida"),
    _rule("alojamento dedicado", "acomodação dedicada"),
    _rule("alojamento associado", "acomodação associada"),
    _rule("alojamento liquidado", "acomodação liquidada"),
    _rule("alojamentos consecutivos", "acomodações consecutivas"),
    _rule("alojamentos sujos", "acomodações sujas"),
    _rule("alojamentos", "acomodações"),
    _rule("alojamento", "acomodação"),
    # 11. lacunas apontadas pelo painel nativo (aba Frontend)
    # "definições" = app settings in the UI, not a mathematical definition.
    _rule("definições", "configurações"),
    _rule("definição", "configuração"),
    _rule("tenha em conta que", "observe que"),
    _rule("tenha em conta", "observe"),
    _rule("carateres", "caracteres"),
    _rule("caráter", "caractere"),
    _rule("eletrónicas", "eletrônicas"),
    _rule("eletrónicos", "eletrônicos"),
    _rule("eletrónica", "eletrônica"),
    _rule("eletrónico", "eletrônico"),
    # 11b. clitic forms the bare-verb rules cannot reach
    _rule("contactá-lo", "contatá-lo"),
    _rule("contactá-la", "contatá-la"),
    _rule("controlos", "controles"),
    _rule("controlo", "controle"),
    # 12. orthographic correction of country names
    _rule("Arménia", "Armênia"),
    _rule("Estónia", "Estônia"),
    _rule("Polónia", "Polônia"),
    _rule("logótipo", "logotipo"),
    # 13. "anular" -> "cancelar": the blanket verb rule runs last, after the
    # inflected forms and collocations it would otherwise mangle. The
    # "fatura anulada(s)" collocations are NOT converted: the PO keeps the BR
    # fiscal distinction "fatura anulada" (void) vs "fatura cancelada".
    _rule("anulado por", "cancelado por"),
    _rule("anulada por", "cancelada por"),
    _rule("anulados", "cancelados"),
    _rule("anuladas", "canceladas"),
    _rule("anulado", "cancelado"),
    _rule("anulada", "cancelada"),
    _rule("anulará", "cancelará"),
    _rule("anular", "cancelar"),
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


#: pt-PT "bare" progressive ("A processar...") -> pt-BR gerund ("Processando...").
#: Only the leading position is rewritten, and only for known verbs: matching
#: any word would corrupt proper nouns ("A Cuiner", "A Icar") and prepositional
#: phrases ("a partir de", "a receber"). Auxiliary forms ("está a carregar")
#: are handled by _PROGRESSIVE_RE, which is safe because of the auxiliary.
_BARE_PROGRESSIVE_VERBS = frozenset(
    """
    processar verificar reconstruir expirar enviar atualizar transferir gerar
    validar executar carregar bloquear desbloquear excluir reverter mostrar
    aplicar adicionar criar cancelar redirecionar selecionar mapear reenviar
    resolver dividir guardar importar exportar sincronizar digitalizar preparar
    converter apagar duplicar imprimir calcular gravar
    """.split()
)
_BARE_PROGRESSIVE_RE = re.compile(
    r"^(\s*)a\s+(" + "|".join(sorted(_BARE_PROGRESSIVE_VERBS, key=len, reverse=True)) + r")\b",
    re.IGNORECASE,
)


def _fix_bare_progressive(text: str) -> str:
    """Rewrite a leading pt-PT "A <verbo>r" as the pt-BR gerund.

    The match is anchored to the start of the string, so the gerund always
    opens the sentence and is therefore capitalised.
    """
    match = _BARE_PROGRESSIVE_RE.match(text)
    if match is None:
        return text
    lead, verb = match.group(1), match.group(2)
    gerund = verb[:-1] + "ndo"
    return lead + gerund[:1].upper() + gerund[1:] + text[match.end():]


def _fix_progressive(text: str) -> str:
    def _repl(match: re.Match[str]) -> str:
        aux, verb = match.group(1), match.group(2)
        return f"{aux} {verb}ndo"

    return _PROGRESSIVE_RE.sub(_repl, text)


#: "acomodação/acomodações" is feminine, but the pt-PT wording it replaces
#: ("alojamento") was masculine, and some source strings already paired
#: "acomodações" with a masculine adjective. Fix the agreement after the fact,
#: whatever order the substitution rules ran in.
_ACOMODACAO_FEM = {
    "provisórios": "provisórias", "provisório": "provisória",
    "selecionados": "selecionadas", "selecionado": "selecionada",
    "atribuídos": "atribuídas", "atribuído": "atribuída",
    "consumidos": "consumidas", "consumido": "consumida",
    "combinados": "combinadas", "combinado": "combinada",
    "reservados": "reservadas", "reservado": "reservada",
    "turísticos": "turísticas", "turístico": "turística",
    "físicos": "físicas", "físico": "física",
    "sujos": "sujas", "sujo": "suja", "consecutivos": "consecutivas",
    "dedicados": "dedicadas", "dedicado": "dedicada",
    "associados": "associadas", "associado": "associada",
    "liquidados": "liquidadas", "liquidado": "liquidada",
    "bloqueados": "bloqueadas", "bloqueado": "bloqueada",
    "ocupados": "ocupadas", "ocupado": "ocupada",
    "criados": "criadas", "criado": "criada",
    "cancelados": "canceladas", "cancelado": "cancelada",
}
_MASC_ADJ_RE = re.compile(
    r"\b(" + "|".join(sorted(_ACOMODACAO_FEM, key=len, reverse=True)) + r")\b",
    re.IGNORECASE,
)
_ACOMODACAO_COMO_RESERVADO_RE = re.compile(
    r"(a\s+acomodaç(?:ão|ões)\b[^.]*?\bcomo\s+)reservado\b", re.IGNORECASE
)
#: Determiner disagreement: "Nenhum das acomodações".
_ACOMODACAO_DET_RE = re.compile(r"\b(Nenh)um(\s+das\s+acomodaç)", re.IGNORECASE)
#: Agreement window: after "acomodação(s)", walk over words that can only
#: belong to its own noun phrase (articles, prepositions, auxiliaries,
#: negations, conjunctions) and re-gender every masculine adjective found.
#: The walk stops at the first other word, so an adjective that agrees with a
#: different noun ("Bloqueio de acomodação provisória consumido") is left
#: untouched.
_ACOM_FILLER_WORDS = (
    "não", "nao", "nem", "já", "ja", "ainda", "estão", "estao", "está", "esta",
    "estavam", "estava", "são", "sao", "é", "eram", "era", "foram", "serão",
    "serao", "seria", "ficarão", "ficarao", "ficaram", "ficam", "fica", "têm",
    "tem", "de", "da", "do", "das", "dos", "com", "para", "em", "no", "na",
    "nos", "nas", "uns", "umas", "um", "uma", "seu", "sua", "seus", "suas",
    "este", "esta", "estes", "estas", "esse", "essa", "esses", "essas",
    "todo", "toda", "todos", "todas", "ou", "os", "as", "o", "a", "e",
)
#: Words the window may walk over: fillers and the masculine adjectives being
#: fixed. Feminine adjectives are NOT transparent, so a participle that agrees
#: with another head noun ("Bloqueio de acomodação provisória consumido") is
#: left alone. Boundaries are enforced so "de" does not swallow the "de" of
#: "dedicado".
_ACOM_WORD_ALT = "|".join(
    re.escape(w)
    for w in sorted(
        set(_ACOM_FILLER_WORDS) | set(_ACOMODACAO_FEM), key=len, reverse=True
    )
)
_ACOMODACAO_WINDOW_RE = re.compile(
    r"\bacomodaç(?:ão|ões)\b((?:\s+(?:" + _ACOM_WORD_ALT + r")\b)*)",
    re.IGNORECASE,
)


def _fix_agreement(text: str) -> str:
    def _window(match: re.Match[str]) -> str:
        head = match.group(0)[: len(match.group(0)) - len(match.group(1))]
        tail = _MASC_ADJ_RE.sub(
            lambda m: _ACOMODACAO_FEM[m.group(1).lower()], match.group(1)
        )
        return head + tail

    text = _ACOMODACAO_WINDOW_RE.sub(_window, text)
    text = _ACOMODACAO_COMO_RESERVADO_RE.sub(r"\1reservada", text)
    return _ACOMODACAO_DET_RE.sub(r"Nenhuma\2", text)


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
    protected = _fix_bare_progressive(protected)
    protected = _fix_agreement(protected)

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


def _pt_br_sheet(sheet: str) -> str:
    """pt-BR label for a source tab, keeping the internal keys untouched."""
    return SHEET_NAMES_PT_BR.get(sheet, sheet)


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


#: Keys whose approved pt-BR text legitimately contains a substring that the
#: generic rules would rewrite, so the automatic residue gate must skip them.
#: ``propertyEmailSchedulerEventINVCVOID.message`` is "Fatura anulada": BR
#: fiscal wording keeps "anulada" (void) distinct from "cancelada".
FALSE_POSITIVOS_PT_PT: set[str] = {
    "shared.radioButton.propertyEmailSchedulerEventINVCVOID.message",
}


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
        if clave in CORRECOES_PO:
            # binding PO decision: wins over both the source warning ("Sin
            # traducción PT" / "Igual a ES/EN") and the review agents, so the
            # row closes as CORRIGIDO.
            converted = CORRECOES_PO[clave]
            estado = (
                ESTADO_OK
                if converted == texto_pt
                else ESTADO_CORRIGIDO
            )
        elif clave in CORRECOES_BR:
            # audit fix for pt-PT the generic rules missed (mostly
            # "Descargar/Download" rendered as "Transferir"): wins over the
            # agents and over the source warning.
            converted = CORRECOES_BR[clave]
            estado = (
                ESTADO_OK
                if converted == texto_pt
                else ESTADO_CORRIGIDO
            )
        elif dec is not None:
            # validated by the review agents: accept their pt-BR verbatim
            converted = dec
            estado = _resolve_estado(aviso, texto_pt, converted)
        elif clave in CORRECOES:
            converted = CORRECOES[clave]
            estado = _resolve_estado(aviso, texto_pt, converted)
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
        # Explanatory note for polemic / ambiguous rows; keep any note the
        # source workbook already carried.
        comentario = _clean(row.get(COL_COMENTARIO))
        if not comentario and clave in COMENTARIOS:
            comentario = COMENTARIOS[clave]
        out.at[idx, COL_COMENTARIO] = comentario
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
        chave_col = next((c for c in KEY_COLUMNS if c in out.columns), None)
        chave = (
            out[chave_col].map(_clean)
            if chave_col
            else pd.Series([""] * len(out), index=out.index)
        )

        def _final(p: str, c: str, e: str, k: str) -> str:
            if c:
                return c
            if not p:
                return ""
            # Only unresolved rows are gated on pt-PT residue; OK/CORRIGIDO text
            # is already validated and must never be blanked by a false positive.
            if (
                e == ESTADO_PENDENTE_PO
                and k not in FALSE_POSITIVOS_PT_PT
                and _tem_residuo_pt_pt(p)
            ):
                return ""
            return p

        out[COL_FINAL] = [
            _final(p, c, e, k)
            for p, c, e, k in zip(pt_original, correcao, estado, chave)
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


def _write_preenchido(
    entrada: Path, destino: Path, processed: dict[str, pd.DataFrame]
) -> None:
    """Write the source workbook with Estado/Corrección PT/Comentario filled.

    The original file is the base: sheet order, headers, formatting, column
    widths and the ``OK,Corregir`` dropdown are preserved. Only the three
    review columns are written, matched by key so the row order never shifts.
    """
    from openpyxl import load_workbook

    wb = load_workbook(entrada)
    for sheet, df in processed.items():
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        # header row is not necessarily row 1 (the source may carry title rows)
        header_row = 1
        for r in range(1, min(ws.max_row, 30) + 1):
            cells = [str(c.value).strip() for c in ws[r] if c.value is not None]
            if any(c in cells for c in KEY_COLUMNS):
                header_row = r
                break
        header = {
            str(c.value).strip(): c.column
            for c in ws[header_row]
            if c.value is not None
        }
        cols = {name: header[name] for name in
                (COL_ESTADO, COL_CORRECAO, COL_COMENTARIO) if name in header}
        if COL_ESTADO not in cols:
            continue
        key_name = next((c for c in KEY_COLUMNS if c in header), None)
        if key_name is None:
            continue
        key_col = header[key_name]

        estado = df[COL_ESTADO].tolist() if COL_ESTADO in df.columns else []
        correcao = df[COL_CORRECAO].tolist() if COL_CORRECAO in df.columns else []
        comentario = df[COL_COMENTARIO].tolist() if COL_COMENTARIO in df.columns else []
        # The key alone is not unique (the technical key column is truncated and
        # repeats across rows). Both the sheet and the frame keep the source row
        # order, so the nth occurrence of a key maps to the nth frame row.
        pending: dict[str, list[int]] = {}
        for i, key in enumerate(df[key_name].tolist()):
            pending.setdefault(str(key).strip(), []).append(i)

        seen: dict[str, int] = {}
        for row in ws.iter_rows(min_row=header_row + 1):
            key = str(row[key_col - 1].value or "").strip()
            slots = pending.get(key)
            if not slots:
                continue
            pos = seen.get(key, 0)
            if pos >= len(slots):
                continue
            seen[key] = pos + 1
            idx = slots[pos]
            for name, col in cols.items():
                src = {
                    COL_ESTADO: estado,
                    COL_CORRECAO: correcao,
                    COL_COMENTARIO: comentario,
                }[name]
                value = src[idx] if idx < len(src) else ""
                row[col - 1].value = value if value else None
    wb.save(destino)


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
            df.to_excel(writer, sheet_name=_pt_br_sheet(sheet)[:31], index=False)

    # entrega principal: o próprio ficheiro original, preenchido nas três
    # colunas de revisão (abas, formatação e dropdown OK/Corregir intactos)
    _write_preenchido(entrada, destino / OUT_PREENCHIDO, processed)

    final = _construir_final(processed)
    with pd.ExcelWriter(destino / OUT_FINAL) as writer:
        for sheet, df in final.items():
            df.to_excel(writer, sheet_name=_pt_br_sheet(sheet)[:31], index=False)

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
