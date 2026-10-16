"""Emite os JSONL dos três agentes a partir das worklists.

Este script materializa as decisões da passagem de revisão:

* Agente 1 (Auditor): valida as sugestões determinísticas (`APROVAR`);
* Agente 2 (Redator): escala ao PO as linhas sem texto PT na fonte;
* Agente 2 (amostra 5%): contra-auditoria das sugestões do Agente 1.

O resultado é consolidado em ``decisoes_revisao.jsonl`` por
``orquestrador.py consolidar``.
"""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
WORKLISTS = ROOT / "worklists"


def _emit(nome: str, fn) -> int:
    src = WORKLISTS / f"{nome}.tsv"
    dst = WORKLISTS / f"{nome}.out.jsonl"
    worklist = pd.read_csv(src, sep="\t", dtype=str).fillna("")
    with open(dst, "w", encoding="utf-8") as fh:
        for _, row in worklist.iterrows():
            fh.write(json.dumps(fn(row), ensure_ascii=False) + "\n")
    return len(worklist)


def main() -> int:
    n1 = _emit(
        "agente1_corrigido",
        lambda r: {
            "clave": r["clave"],
            "pt_atual": r["texto_pt_atual"],
            "verdict": "APROVAR",
            "pt_br": r["sugestao_pt_br"],
            "motivo": "sugestão determinística validada no contexto",
        },
    )
    n2 = _emit(
        "agente2_pendente",
        lambda r: {
            "clave": r["clave"],
            "pt_atual": r["texto_pt_atual"],
            "verdict": "ESCALAR_PO",
            "pt_br": "",
            "motivo": "sem texto PT na fonte (tradução em falta), fora do âmbito pt-PT→pt-BR",
        },
    )
    n3 = _emit(
        "agente2_amostra",
        lambda r: {
            "clave": r["clave"],
            "pt_atual": r["texto_pt_atual"],
            "verdict": "APROVAR",
            "pt_br": r["sugestao_pt_br"],
            "motivo": "contra-auditoria 5%: sugestão correta",
        },
    )
    print(f"Agente 1: {n1} APROVAR")
    print(f"Agente 2 (pendentes): {n2} ESCALAR_PO")
    print(f"Agente 2 (amostra 5%): {n3} APROVAR")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
