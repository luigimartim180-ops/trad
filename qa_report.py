"""Validações de aceitação e relatório de sign-off (Agente 3 — QA).

Corre sobre os outputs do pipeline e o JSONL consolidado dos agentes, e emite
``PASS``/``FAIL`` por verificação. Sem ``PASS`` global, o patch não é aplicado.

Uso::

    python qa_report.py            # imprime o relatório
    python qa_report.py --strict   # sai com código 1 em qualquer FAIL
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

import pandas as pd

import pt_pt_para_pt_br as conv

ROOT = Path(__file__).resolve().parent
PLACEHOLDER_RE = re.compile(r"\{\{.*?\}\}|\{\d+\}|<[^>]+>")
LIMITE_REPROVACAO_AMOSTRA = 0.03


def _clean(value: object) -> str:
    # must match conv._clean exactly (no strip) so worklist keys line up
    return conv._clean(value)


def _carregar_tudo() -> pd.DataFrame:
    sheets = pd.read_excel(ROOT / conv.OUT_COMPLETA, sheet_name=None)
    frames = []
    for name, df in sheets.items():
        df = df.copy()
        # key/PT column names differ per sheet, so normalise before concat
        df["clave"] = _series(df, conv.KEY_COLUMNS)
        df["texto_pt"] = _series(df, conv.PT_COLUMNS)
        df["hoja"] = name
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


def _series(df: pd.DataFrame, candidatos: list[str]) -> pd.Series:
    for c in candidatos:
        if c in df.columns:
            return df[c].map(_clean)
    return pd.Series([""] * len(df), index=df.index)


def _clave(df: pd.DataFrame) -> pd.Series:
    return df["clave"] if "clave" in df.columns else _series(df, conv.KEY_COLUMNS)


def _pt(df: pd.DataFrame) -> pd.Series:
    return df["texto_pt"] if "texto_pt" in df.columns else _series(df, conv.PT_COLUMNS)


def _check(nome: str, falhas: list[str]) -> dict:
    return {"nome": nome, "ok": not falhas, "falhas": falhas[:20], "total": len(falhas)}


def verificar_placeholders(rows: pd.DataFrame) -> dict:
    falhas = []
    for clave, pt, corr in zip(_clave(rows), _pt(rows), rows[conv.COL_CORRECAO].map(_clean)):
        alvo = corr or pt
        if sorted(PLACEHOLDER_RE.findall(pt)) != sorted(PLACEHOLDER_RE.findall(alvo)):
            falhas.append(f"{clave}: placeholders alterados")
    return _check("placeholders íntegros", falhas)


def verificar_residuo_pt_pt(rows: pd.DataFrame) -> dict:
    marcador = getattr(conv, "_PT_PT_MARKER", None)
    falhas = []
    if marcador is None:
        return _check("zero resíduo pt-PT", [])
    for clave, corr in zip(_clave(rows), rows[conv.COL_CORRECAO].map(_clean)):
        if not corr:
            continue
        m = marcador.search(corr)
        if m:
            falhas.append(f"{clave}: {m.group(0)!r}")
    return _check("zero resíduo pt-PT nas correções", falhas)


def verificar_ok_sem_correcao(rows: pd.DataFrame) -> dict:
    falhas = []
    for clave, estado, corr in zip(
        _clave(rows), rows[conv.COL_ESTADO].map(_clean), rows[conv.COL_CORRECAO].map(_clean)
    ):
        if estado == conv.ESTADO_OK and corr:
            falhas.append(f"{clave}: OK com correção")
    return _check("OK ⇒ Corrección PT vazia", falhas)


def verificar_duvidas_intocadas(rows: pd.DataFrame) -> dict:
    falhas = []
    for clave, estado, corr in zip(
        _clave(rows), rows[conv.COL_ESTADO].map(_clean), rows[conv.COL_CORRECAO].map(_clean)
    ):
        if clave in conv.DUVIDAS_KEYS and (estado != conv.ESTADO_PENDENTE_PO or corr):
            falhas.append(f"{clave}: dúvida PO alterada")
    return _check("dúvidas PO intocadas", falhas)


def verificar_sem_revisar(rows: pd.DataFrame) -> dict:
    n = int((rows[conv.COL_ESTADO].map(_clean) == "REVISAR").sum())
    return _check("REVISAR eliminado", [] if n == 0 else [f"{n} linhas em REVISAR"])


def verificar_cobertura(rows: pd.DataFrame) -> dict:
    decisoes = conv.load_decisoes(ROOT / conv.OUT_DECISOES)
    if not decisoes:
        return _check("cobertura das worklists", [])
    chaves = set(zip(_clave(rows), _pt(rows)))
    falhas = [f"{k[0]}: {k[1][:40]!r}" for k in decisoes if k not in chaves]
    return _check("cobertura das worklists", falhas)


def verificar_amostra() -> dict:
    """Contra-auditoria: reprovação da amostra de 5% pelo Agente 2."""
    amostra = ROOT / "worklists" / "agente2_amostra.out.jsonl"
    if not amostra.exists():
        return _check("amostra 5% (contra-auditoria)", [])
    total = 0
    reprovadas = 0
    for line in amostra.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        rec = json.loads(line)
        total += 1
        if rec.get("verdict") == "CORRIGIR":
            reprovadas += 1
    taxa = reprovadas / total if total else 0.0
    falhas = []
    if total and taxa > LIMITE_REPROVACAO_AMOSTRA:
        falhas.append(
            f"reprovação {taxa:.1%} > {LIMITE_REPROVACAO_AMOSTRA:.0%} "
            f"({reprovadas}/{total}); Agente 1 deve refazer o lote"
        )
    return _check(f"amostra 5% ({reprovadas}/{total} reprovadas)", falhas)


def relatorio() -> int:
    rows = _carregar_tudo()
    verificacoes = [
        verificar_placeholders(rows),
        verificar_residuo_pt_pt(rows),
        verificar_ok_sem_correcao(rows),
        verificar_duvidas_intocadas(rows),
        verificar_sem_revisar(rows),
        verificar_cobertura(rows),
        verificar_amostra(),
    ]

    print("=" * 62)
    print("RELATÓRIO DE SIGN-OFF — QA pt-PT → pt-BR")
    print("=" * 62)
    for v in verificacoes:
        marca = "PASS" if v["ok"] else "FAIL"
        extra = "" if v["ok"] else f"  ({v['total']} falhas)"
        print(f"[{marca}] {v['nome']}{extra}")
        for f in v["falhas"]:
            print(f"        - {f}")

    contagem = rows[conv.COL_ESTADO].map(_clean).value_counts().to_dict()
    print("-" * 62)
    print("Contagem por estado:")
    for estado in (conv.ESTADO_OK, conv.ESTADO_CORRIGIDO, conv.ESTADO_PENDENTE_PO):
        print(f"  {estado}: {contagem.get(estado, 0)}")
    print(f"  Total: {len(rows)}")

    ok = all(v["ok"] for v in verificacoes)
    print("-" * 62)
    print("RESULTADO GLOBAL:", "PASS ✅" if ok else "FAIL ❌")
    return 0 if ok else 1


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--strict", action="store_true", help="exit 1 em qualquer FAIL")
    args = parser.parse_args(argv)
    code = relatorio()
    return code if args.strict else 0


if __name__ == "__main__":
    raise SystemExit(main())
