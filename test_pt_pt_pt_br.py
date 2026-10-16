"""Unit tests for the pt-PT -> pt-BR pipeline, using real workbook cases."""

from __future__ import annotations

import re
import unittest
from pathlib import Path

import pandas as pd

import pt_pt_para_pt_br as conv

ROOT = Path(__file__).resolve().parent
PLACEHOLDER_RE = re.compile(r"\{\{.*?\}\}|\{\d+\}|<[^>]+>")


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
            "Tem certeza de que pretende…?",
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
            pt = str(row["Traducción PT actual"]).strip().lower()
            if pt in conv.WHITELIST_IGUAL_ES_EN:
                self.assertEqual(row[conv.COL_ESTADO], "OK")
            else:
                self.assertEqual(row[conv.COL_ESTADO], "REVISAR")


if __name__ == "__main__":
    unittest.main(verbosity=2)
