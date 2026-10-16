"""Unit tests for the pt-PT -> pt-BR pipeline, using real workbook cases."""

from __future__ import annotations

import re
import unittest
from pathlib import Path

import pandas as pd

import pt_pt_para_pt_br as conv

ROOT = Path(__file__).resolve().parent
PLACEHOLDER_RE = re.compile(r"\{\{.*?\}\}|\{\d+\}|<[^>]+>")


def _clean(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value)


class TestConverter(unittest.TestCase):
    def test_utilizador_uppercase(self) -> None:
        # real case: Mensajes e informes / field.user
        self.assertEqual(conv.converter("UTILIZADOR:"), "USUÁRIO:")

    def test_utilizadores_plural(self) -> None:
        self.assertEqual(conv.converter("Utilizadores"), "Usuários")

    def test_a_carregar_compound_before_carregar(self) -> None:
        # real case: Frontend "A carregar..."
        self.assertEqual(conv.converter("A carregar…"), "Carregando…")
        self.assertNotIn("enviar", conv.converter("A carregar…").lower())

    def test_a_carregar_does_not_match_para_carregar(self) -> None:
        # real case: "Clique aqui para carregar mais" -> no "parcarregando"
        out = conv.converter("Clique aqui para carregar mais")
        self.assertEqual(out, "Clique aqui para enviar mais")

    def test_tem_a_certeza(self) -> None:
        self.assertEqual(
            conv.converter("Tem a certeza de que pretende…?"),
            "Tem certeza de que deseja…?",
        )

    def test_enviamos(self) -> None:
        self.assertEqual(
            conv.converter("Enviámos um e-mail…"), "Enviamos um e-mail…"
        )

    def test_placeholders_and_html_preserved(self) -> None:
        source = "Olá {{value}}<br>Obrigado"
        out = conv.converter(source)
        self.assertIn("{{value}}", out)
        self.assertIn("<br>", out)
        self.assertNotIn("\x00", out)

    def test_indexed_placeholder_preserved(self) -> None:
        source = "Olá {0}, tem {1} mensagens"
        self.assertEqual(conv.converter(source), source)

    def test_strong_placeholder_preserved(self) -> None:
        out = conv.converter("Olá <strong>{{name}}</strong>")
        self.assertEqual(out, "Olá <strong>{{name}}</strong>")

    def test_linhas_untouched(self) -> None:
        # real case: shared.grid.noViewAllDataInGridReport ("5000 linhas")
        self.assertEqual(conv.converter("5000 linhas"), "5000 linhas")

    def test_allotment_kept_intact(self) -> None:
        # real case: shared.forecast.nonDeductedAllotment
        self.assertEqual(
            conv.converter("Allotment não deduzido"), "Allotment não deduzido"
        )

    def test_no_typeerror_on_varied_input(self) -> None:
        for text in ["", "texto normal", "{{x}}", "<br>", "A carregar", None]:
            if text is None:
                self.assertIsNone(conv.converter(text))
            else:
                self.assertIsInstance(conv.converter(text), str)

    def test_contacto_ao90(self) -> None:
        self.assertEqual(conv.converter("Contacto do perfil"), "Contato do perfil")
        self.assertEqual(conv.converter("Contactos"), "Contatos")

    def test_ao90_spelling(self) -> None:
        for pt, br in [
            ("Actividade", "Atividade"),
            ("Acção", "Ação"),
            ("Objectivo", "Objetivo"),
            ("Protecção", "Proteção"),
            ("Seleccionar", "Selecionar"),
        ]:
            self.assertEqual(conv.converter(pt), br)

    def test_por_defeito_compound(self) -> None:
        self.assertEqual(
            conv.converter("por defeito, se nada for configurado"),
            "por padrão, se nada for configurado",
        )

    def test_detetado_and_plural(self) -> None:
        self.assertEqual(conv.converter("detetado"), "detectado")
        self.assertEqual(conv.converter("detetados"), "detectados")
        self.assertEqual(conv.converter("detetadas"), "detectadas")

    def test_transferir_kept_for_account_actions(self) -> None:
        # "transferir" is a real account/folio move, not a download
        self.assertEqual(
            conv.converter("Transferir para outra conta"),
            "Transferir para outra conta",
        )

    def test_download_keys_use_baixar(self) -> None:
        self.assertEqual(conv.CORRECOES["core.button.download"], "Baixar")
        self.assertEqual(
            conv.CORRECOES["apps.report.downloadExcel"], "Baixar Excel"
        )

    def test_duvida_keys_never_converted(self) -> None:
        for key in [
            "apps.aiAssistant.label.hotel",
            "apps.chainSpaceLockType.radio.onHold",
            "apps.chainSpaceLockType.radio.outOfOrder",
            "apps.chainSpaceLockType.radio.outOfService",
            "PropertyClassType t WHERE t.Code = '9'",
        ]:
            self.assertIsNone(conv.CORRECOES[key])

    def test_pensao_kept(self) -> None:
        self.assertEqual(conv.converter("Pensão"), "Pensão")

    def test_allotment_still_kept(self) -> None:
        self.assertEqual(
            conv.converter("Allotment não deduzido"), "Allotment não deduzido"
        )

    def test_verb_forms_pt_pt_to_pt_br(self) -> None:
        cases = {
            "de que pretende prosseguir": "de que deseja prosseguir",
            "opção se pretender tentar": "opção se quiser tentar",
            "opção se pretende": "opção se deseja",
            "Encontrámos alguns erros": "Encontramos alguns erros",
            "detetámos": "detectamos",
            "Reserva guardada": "Reserva salva",
            "Mensagem eliminada": "Mensagem excluída",
            "O código introduzido": "O código informado",
            "Permitir bebé": "Permitir bebê",
            "mais cómoda": "mais cômoda",
            "Todos os separadores": "Todas as abas",
        }
        for pt, br in cases.items():
            self.assertEqual(conv.converter(pt), br, pt)

    def test_anulacao_and_eliminacao(self) -> None:
        self.assertEqual(conv.converter("anulação"), "cancelamento")
        self.assertEqual(conv.converter("anulações"), "cancelamentos")
        self.assertEqual(conv.converter("eliminação"), "exclusão")

    def test_detetar_and_carregada(self) -> None:
        self.assertEqual(conv.converter("detetar"), "detectar")
        self.assertEqual(
            conv.converter("será carregada na seção"), "será enviada na seção"
        )

    def test_panel_identified_terms(self) -> None:
        cases = {
            "Pode aceder ao detalhe": "Pode acessar ao detalhe",
            "Aceda à seção": "Acesse à seção",
            "pronto para gerir a propriedade": "pronto para gerenciar a propriedade",
            "proteja a sua equipa": "proteja a sua equipe",
            "um montante ou percentagem": "um valor ou porcentagem",
            "coluna da grelha": "coluna da grade",
            "dados pessoais com terceiros, se pretender partilhar": (
                "dados pessoais com terceiros, se quiser compartilhar"
            ),
            "prima pesquisar": "clique pesquisar",
            "recolha de caixa": "coleta de caixa",
        }
        for pt, br in cases.items():
            self.assertEqual(conv.converter(pt), br, pt)

    def test_gerir_derivatives(self) -> None:
        for pt, br in {
            "gerir": "gerenciar",
            "gerido": "gerenciado",
            "gerida": "gerenciada",
            "geridos": "gerenciados",
            "geridas": "gerenciadas",
        }.items():
            self.assertEqual(conv.converter(pt), br, pt)

    def test_progressive_becomes_gerund(self) -> None:
        for pt, br in {
            "está a carregar os dados": "está carregando os dados",
            "está a gerir as reservas": "está gerenciando as reservas",
            "estão a gerir": "estão gerenciando",
            "está a ser configurada": "está sendo configurada",
            "está a tentar excluir": "está tentando excluir",
            "está a demorar muito": "está demorando muito",
        }.items():
            self.assertEqual(conv.converter(pt), br, pt)

    def test_noun_aplicacao_is_not_verb(self) -> None:
        # "de aplicação" / "Intervalo de aplicação" = application OF a deposit,
        # not the software "aplicação" -> must stay untouched.
        for text in (
            "Data de aplicação",
            "Intervalo de aplicação",
            "Modo de aplicação do depósito",
        ):
            self.assertEqual(conv.converter(text), text)


class TestRules(unittest.TestCase):
    def test_compound_rules_come_first(self) -> None:
        patterns = [p.pattern for p, _ in conv.SUBSTITUICOES]
        compound = patterns.index(r"\ba\s+carregar\b")
        single = patterns.index(r"\bcarregar\b")
        self.assertLess(compound, single)

    def test_plural_before_singular(self) -> None:
        order = [p.pattern for p, _ in conv.SUBSTITUICOES]
        plural = next(i for i, p in enumerate(order) if p == r"\butilizadores\b")
        singular = next(i for i, p in enumerate(order) if p == r"\butilizador\b")
        self.assertLess(plural, singular)

    def test_all_rules_have_word_boundaries(self) -> None:
        for pattern, _ in conv.SUBSTITUICOES:
            self.assertTrue(pattern.pattern.startswith(r"\b"))
            self.assertTrue(pattern.pattern.endswith(r"\b"))

    def test_no_global_allotment_rule(self) -> None:
        for pattern, repl in conv.SUBSTITUICOES:
            self.assertNotEqual(repl.lower(), "cota")
            self.assertNotIn("allotment", pattern.pattern.lower())


class TestPipeline(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        xlsx = ROOT / conv.XLSX_IN
        if not xlsx.exists():
            raise unittest.SkipTest(f"{xlsx} not available")
        cls.xl = pd.ExcelFile(xlsx)
        cls.sheets = conv._review_sheet_names(cls.xl)
        cls.processed = {
            s: conv.process_sheet(conv._read_sheet(cls.xl, s))
            for s in cls.sheets
        }
        cls.all_rows = pd.concat(cls.processed.values(), ignore_index=True)

    def test_review_sheets_detected(self) -> None:
        self.assertIn("Frontend (pantallas)", self.sheets)
        self.assertIn("Entidades de sistema", self.sheets)
        self.assertIn("Mensajes e informes", self.sheets)
        self.assertNotIn(conv.INFO_SHEET, self.sheets)

    def test_header_detected_below_title_rows(self) -> None:
        df = conv._read_sheet(self.xl, "Frontend (pantallas)")
        self.assertIn("Clave", df.columns)

    def test_no_ok_row_has_correction(self) -> None:
        ok_rows = self.all_rows[self.all_rows[conv.COL_ESTADO] == "OK"]
        filled = ok_rows[
            ok_rows[conv.COL_CORRECAO].notna()
            & ok_rows[conv.COL_CORRECAO].astype(str).str.strip().ne("")
        ]
        self.assertTrue(filled.empty, f"OK rows with correction:\n{filled}")

    def test_placeholder_integrity_across_workbook(self) -> None:
        for _, row in self.all_rows.iterrows():
            src = str(row.get("Traducción PT actual", ""))
            dst = row.get(conv.COL_CORRECAO)
            if dst is None or (isinstance(dst, float) and pd.isna(dst)):
                dst = src
            dst = str(dst)
            if not dst.strip():
                # empty correction means the pt-PT text is kept unchanged
                dst = src
            self.assertNotIn("\x00", dst)
            self.assertEqual(
                sorted(PLACEHOLDER_RE.findall(src)),
                sorted(PLACEHOLDER_RE.findall(dst)),
                f"placeholder mismatch for {row.get('Clave')}",
            )

    def test_dead_columns_removed(self) -> None:
        for df in self.processed.values():
            for col in conv.DEAD_COLUMNS:
                self.assertNotIn(col, df.columns)

    def test_surgical_corrections(self) -> None:
        by_key = self.all_rows.groupby("Clave").first()
        self.assertEqual(
            by_key.at["card.hojaRegistro", conv.COL_CORRECAO],
            "Ficha de registro de hóspedes (FNRH)",
        )
        self.assertEqual(
            by_key.at["core.button.settlement", conv.COL_CORRECAO],
            "Cobranças automáticas",
        )
        self.assertEqual(
            by_key.at["invoice.direction.label", conv.COL_CORRECAO], "Endereço"
        )

    def test_allotment_row_not_converted(self) -> None:
        # dúvida: "Allotment não deduzido" must stay intact (never converted)
        by_key = self.all_rows.groupby("Clave").first()
        self.assertEqual(
            by_key.at["shared.forecast.nonDeductedAllotment", "Traducción PT actual"],
            "Allotment não deduzido",
        )
        self.assertEqual(
            by_key.at["shared.forecast.nonDeductedAllotment", conv.COL_CORRECAO], ""
        )

    def test_igual_es_en_reviewed_unless_whitelisted(self) -> None:
        fe = self.processed["Frontend (pantallas)"]
        igual = fe[fe["Aviso"].astype(str).str.startswith("Igual")]
        self.assertFalse(igual.empty)
        for _, row in igual.iterrows():
            pt = conv._clean(row["Traducción PT actual"])
            clave = row["Clave"]
            if (
                not pt
                or clave in conv.DUVIDAS_KEYS
                or conv.CORRECOES.get(clave, "") is None
            ):
                self.assertEqual(row[conv.COL_ESTADO], conv.ESTADO_PENDENTE_PO)
                continue
            converted = conv.CORRECOES.get(clave) or conv.converter(pt)
            if clave in conv.OK_MANUAL or (
                converted == pt and conv._igual_es_en_is_safe(pt)
            ):
                self.assertEqual(row[conv.COL_ESTADO], conv.ESTADO_OK)
            else:
                self.assertEqual(row[conv.COL_ESTADO], conv.ESTADO_CORRIGIDO)

    def test_no_pt_pt_marker_is_closed_as_ok(self) -> None:
        for _, row in self.all_rows.iterrows():
            if row[conv.COL_ESTADO] != "OK":
                continue
            mark = conv._PT_PT_MARKER.search(str(row["Traducción PT actual"]))
            self.assertIsNone(
                mark,
                f"{row['Clave']}: fechado OK com pt-PT {mark and mark.group(0)!r}",
            )

    def test_priority_column_present(self) -> None:
        for df in self.processed.values():
            self.assertIn(conv.COL_PRIORIDADE, df.columns)

    def test_priority_values_consistent(self) -> None:
        for _, row in self.all_rows.iterrows():
            estado = row[conv.COL_ESTADO]
            prio = row[conv.COL_PRIORIDADE]
            if estado == "OK":
                self.assertIn(prio, {"", "Fechado"})
                continue
            self.assertIn(prio, {
                "P0 - dúvida PO",
                "P1 - decisão manual",
                "P2 - validar sugestão",
                "Fechado",
            })

    def test_ok_manual_closed_as_ok(self) -> None:
        by_key = self.all_rows.groupby("Clave").first()
        for key in ("core.label.total", "core.label.inHouse"):
            self.assertEqual(by_key.loc[key, conv.COL_ESTADO], "OK", key)
            self.assertEqual(_clean(by_key.loc[key, conv.COL_CORRECAO]), "", key)
            self.assertEqual(by_key.loc[key, conv.COL_PRIORIDADE], "Fechado", key)

    def test_reviewed_corrections_marked_fechado(self) -> None:
        by_key = self.all_rows.groupby("Clave").first()
        for key in (
            "core.label.taxExcludedShort",
            "shared.modal.information.anonymize.message",
            "before.checkout.message.intro",
        ):
            self.assertEqual(by_key.loc[key, conv.COL_PRIORIDADE], "Fechado", key)
        self.assertEqual(
            by_key.loc["core.label.taxExcludedShort", conv.COL_CORRECAO],
            "Imposto excl.",
        )

    def test_missing_translations_filled(self) -> None:
        by_key = self.all_rows.groupby("Clave").first()
        for key, expected in {
            "shared.assistant.propertyAutomatedCheckinScheduler.step0": "Resumo",
            "shared.assistant.propertyAutomatedCheckinScheduler.step2": "Ação",
            "core.object.entity.propertyCheckInAutomationScheduler": (
                "Check-in automático"
            ),
        }.items():
            self.assertEqual(by_key.loc[key, conv.COL_CORRECAO], expected, key)

    def test_anonymize_restores_es_question(self) -> None:
        by_key = self.all_rows.groupby("Clave").first()
        value = by_key.loc[
            "shared.modal.information.anonymize.message", conv.COL_CORRECAO
        ]
        self.assertTrue(value.endswith("deseja continuar?"))

    def test_budget_instructions_multiline_placeholders(self) -> None:
        by_key = self.all_rows.groupby("Clave").first()
        value = by_key.loc[
            "apps.budget.importExcelModal.instructions", conv.COL_CORRECAO
        ]
        self.assertIn("acomodações", value)
        self.assertIn("Salvar na tela principal", value)
        self.assertIn("Baixe o modelo Excel", value)

    def test_no_pt_pt_residue_in_suggestions(self) -> None:
        residue = re.compile(
            r"\b(pretend\w+|encontrámos|guardad\w+|eliminad\w+|separador\w*|bebé\w*"
            r"|cómod\w+|detet\w+|morad\w+|cadeia\w*|faturação|receção|ecrã\w*"
            r"|ficheiro\w*|palavra-passe\w*|montant\w+|planeament\w*|registo\w*"
            r"|activ\w+|acç\w+|objectiv\w+|contacto\w*|aceder|aceda|gerir|equipa\w*"
            r"|partilhar\w*|prima|percentagem\w*|grelha\w*|recolha|demasiado\w*"
            r"|telemóvel\w*|utilizador\w*|subscrição|anulação\w*)\b",
            re.IGNORECASE,
        )
        for _, row in self.all_rows.iterrows():
            suggestion = _clean(row[conv.COL_CORRECAO])
            if not suggestion:
                continue
            match = residue.search(suggestion)
            self.assertIsNone(
                match,
                None if match is None else f"{row['Clave']}: resíduo {match.group(0)!r}",
            )

    def test_duvida_rows_are_p0_and_empty(self) -> None:
        by_key = self.all_rows.groupby("Clave").first()
        for grupo, keys in conv.DUVIDAS_GRUPOS.items():
            for key in keys:
                if key not in by_key.index:
                    continue
                row = by_key.loc[key]
                self.assertEqual(row[conv.COL_ESTADO], conv.ESTADO_PENDENTE_PO, key)
                self.assertEqual(_clean(row[conv.COL_CORRECAO]), "", key)
                self.assertEqual(row[conv.COL_PRIORIDADE], "P0 - dúvida PO", key)

    def test_pendentes_file_contains_only_non_ok(self) -> None:
        path = ROOT / conv.OUT_REVISAR
        if not path.exists():
            raise unittest.SkipTest(f"{path} not generated")
        sheets = pd.read_excel(path, sheet_name=None)
        total = 0
        for name, df in sheets.items():
            self.assertIn(conv.COL_ESTADO, df.columns, name)
            self.assertTrue((df[conv.COL_ESTADO] != conv.ESTADO_OK).all(), name)
            total += len(df)
        expected = (self.all_rows[conv.COL_ESTADO] != conv.ESTADO_OK).sum()
        self.assertEqual(total, expected)

    def test_backend_csv_only_corrigido(self) -> None:
        path = ROOT / conv.OUT_CSV
        if not path.exists():
            raise unittest.SkipTest(f"{path} not generated")
        csv_rows = pd.read_csv(path, sep=";", dtype=str).fillna("")
        self.assertFalse(csv_rows.empty)
        self.assertTrue((csv_rows[conv.COL_ESTADO] == conv.ESTADO_CORRIGIDO).all())
        self.assertTrue((csv_rows[conv.COL_CORRECAO].str.strip() != "").all())
        expected = (self.all_rows[conv.COL_ESTADO] == conv.ESTADO_CORRIGIDO).sum()
        self.assertEqual(len(csv_rows), expected)

    def test_no_revisar_state_left(self) -> None:
        self.assertEqual((self.all_rows[conv.COL_ESTADO] == "REVISAR").sum(), 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
