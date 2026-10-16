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
OUT_DUVIDAS = "duvidas_PO.xlsx"
OUT_CSV = "correcoes_backend.csv"

KEY_COLUMNS = ["Clave", "Clave técnica (no editar)", "Código"]
PT_COLUMNS = ["Traducción PT actual", "Texto PT", "Traducción PT"]
COL_ESTADO = "Estado"
COL_CORRECAO = "Corrección PT"
COL_COMENTARIO = "Comentario"
COL_AVISO = "Aviso"
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
DUVIDAS_KEYS = [
    "core.label.billing",
    "core.label.chain",
    "shared.forecast.nonDeductedAllotment",
]

#: Surgical, key-scoped overrides that bypass the generic rules.
CORRECOES = {
    "card.hojaRegistro": "Ficha de registro de hóspedes (FNRH)",
    "core.button.settlement": "Cobranças automáticas",
    "invoice.direction.label": "Endereço",
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
    # 3. remaining single verbs / nouns
    _rule("palavra-passe", "senha"),
    _rule("secção", "seção"),
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


def process_sheet(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    columns = list(out.columns)
    key_col = _pick(columns, KEY_COLUMNS)
    pt_col = _pick(columns, PT_COLUMNS)
    aviso_col = _pick(columns, [COL_AVISO])

    for col in (COL_ESTADO, COL_CORRECAO, COL_COMENTARIO):
        if col not in out.columns:
            out[col] = ""
        out[col] = out[col].map(_clean)

    for idx, row in out.iterrows():
        clave = _key_value(row, key_col)
        texto_pt = _clean(row.get(pt_col)) if pt_col else ""
        aviso = _clean(row.get(aviso_col)) if aviso_col else ""

        if clave in DUVIDAS_KEYS:
            # open question for the PO: never auto-converted, listed in
            # duvidas_PO.xlsx for a human decision
            out.at[idx, COL_ESTADO] = "REVISAR"
            out.at[idx, COL_CORRECAO] = ""
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
            if clave in DUVIDAS_KEYS:
                rows.append(
                    {
                        "Clave": clave,
                        "Hoja": sheet,
                        "Texto PT": _clean(row.get(pt_col)) if pt_col else "",
                        "Comentario": _clean(row.get(COL_COMENTARIO)),
                        "Destino": "PO",
                    }
                )
    return pd.DataFrame(
        rows, columns=["Clave", "Hoja", "Texto PT", "Comentario", "Destino"]
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

    duvidas = _extract_duvidas(xl, processed)
    duvidas.to_excel(root / OUT_DUVIDAS, index=False)

    rows = pd.concat(processed.values(), ignore_index=True)
    backend = rows[rows[COL_ESTADO] != "OK"]
    with open(root / OUT_CSV, "w", newline="", encoding="utf-8") as fh:
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
    for estado, total in sorted(contagem.items()):
        print(f"  {estado}: {total}")
    print(f"Total de linhas: {len(rows)}")
    print(f"Substituições distintas: {len(SUBSTITUICOES)}")
    print(f"Dúvidas (não convertidas): {len(duvidas)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(build())
