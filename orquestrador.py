"""Orquestrador das três passagens de revisão pt-PT → pt-BR.

Gera as *worklists* (TSV) a partir do resultado do pipeline, aplica a amostragem
de 5% para contra-auditoria e consolida os JSONL devolvidos pelos agentes em
``decisoes_revisao.jsonl``.

Uso::

    python orquestrador.py worklists   # gera worklists/*.tsv + amostra
    python orquestrador.py consolidar  # funde worklists/*.out.jsonl
    python orquestrador.py aplicar     # regenera xlsx/csv via pipeline

O formato de entrada e saída está documentado em ``agentes.md``.
"""

from __future__ import annotations

import argparse
import json
import random
import sys
from pathlib import Path

import pandas as pd

import pt_pt_para_pt_br as conv

ROOT = Path(__file__).resolve().parent
WORKLISTS = ROOT / "worklists"
AMOSTRA_FRACAO = 0.05
AMOSTRA_MINIMA = 143
COLUNAS_WORKLIST = [
    "clave",
    "texto_pt_atual",
    "sugestao_pt_br",
    "aviso",
    "modulo",
]

#: chaves do pipeline que são decisão de glossário do PO (nunca convertidas)
_KEYS_GLOSSARIO = set(conv.DUVIDAS_KEYS) | {
    k for k, v in conv.CORRECOES.items() if v is None
} | {k for keys in conv.DUVIDAS_GRUPOS.values() for k in keys}


def _modulo(clave: str) -> str:
    return clave.split(".", 1)[0] if "." in clave else clave


def _carregar_completa() -> pd.DataFrame:
    path = ROOT / conv.OUT_COMPLETA
    if not path.exists():
        raise SystemExit(
            f"ERRO: {conv.OUT_COMPLETA} não existe; corre pt_pt_para_pt_br.py primeiro"
        )
    sheets = pd.read_excel(path, sheet_name=None)
    frames = []
    for name, df in sheets.items():
        df = df.copy()
        key_col = _coluna(df, conv.KEY_COLUMNS)
        pt_col = _coluna(df, conv.PT_COLUMNS)
        aviso_col = _coluna(df, [conv.COL_AVISO])
        # normalize per sheet: key/PT/aviso column names differ between sheets
        df["clave"] = df[key_col].map(conv._clean) if key_col else ""
        df["texto_pt_atual"] = df[pt_col].map(conv._clean) if pt_col else ""
        df["aviso"] = df[aviso_col].map(conv._clean) if aviso_col else ""
        df["hoja"] = name
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


def _coluna(df: pd.DataFrame, candidatos: list[str]) -> str | None:
    for c in candidatos:
        if c in df.columns:
            return c
    return None


def gerar_worklists(seed: int = 42) -> dict[str, int]:
    """Gera as worklists dos Agentes 1 e 2 e a amostra de contra-auditoria."""
    WORKLISTS.mkdir(exist_ok=True)
    rows = _carregar_completa()

    registos = []
    for _, row in rows.iterrows():
        clave = row["clave"]
        if not clave:
            continue
        registos.append(
            {
                "clave": clave,
                "texto_pt_atual": row["texto_pt_atual"],
                "sugestao_pt_br": conv._clean(row.get(conv.COL_CORRECAO)),
                "aviso": row["aviso"],
                "modulo": _modulo(clave),
                "estado": conv._clean(row.get(conv.COL_ESTADO)),
            }
        )
    todos = pd.DataFrame(registos)

    corrigido = todos[todos["estado"] == conv.ESTADO_CORRIGIDO].copy()
    pendente = todos[
        (todos["estado"] == conv.ESTADO_PENDENTE_PO)
        & (~todos["clave"].isin(_KEYS_GLOSSARIO))
    ].copy()

    # amostra estratificada por módulo, 5% das linhas CORRIGIDO
    rng = random.Random(seed)
    n_amostra = max(AMOSTRA_MINIMA, int(len(corrigido) * AMOSTRA_FRACAO))
    n_amostra = min(n_amostra, len(corrigido))
    amostra_idx: list[int] = []
    for _, grupo in corrigido.groupby("modulo"):
        quota = max(1, round(len(grupo) * AMOSTRA_FRACAO))
        quota = min(quota, len(grupo))
        amostra_idx.extend(rng.sample(list(grupo.index), quota))
    if len(amostra_idx) < n_amostra:
        restantes = [i for i in corrigido.index if i not in set(amostra_idx)]
        amostra_idx.extend(rng.sample(restantes, min(n_amostra - len(amostra_idx), len(restantes))))
    amostra = corrigido.loc[sorted(set(amostra_idx))].copy()

    paths = {
        "agente1_corrigido": corrigido,
        "agente2_pendente": pendente,
        "agente2_amostra": amostra,
    }
    contagem: dict[str, int] = {}
    for nome, df in paths.items():
        out = df[COLUNAS_WORKLIST]
        destino = WORKLISTS / f"{nome}.tsv"
        out.to_csv(destino, sep="\t", index=False)
        contagem[nome] = len(out)
        print(f"  {destino.name}: {len(out)} linhas")

    print(
        f"\nAgente 1: {contagem['agente1_corrigido']} linhas | "
        f"Agente 2: {contagem['agente2_pendente']} pendentes + "
        f"{contagem['agente2_amostra']} amostra (5%)"
    )
    print(f"Chaves de glossário excluídas das worklists: {len(_KEYS_GLOSSARIO)}")
    return contagem


def consolidar(destino: Path | None = None) -> int:
    """Funde os ``*.out.jsonl`` das worklists em ``decisoes_revisao.jsonl``."""
    if destino is None:
        destino = ROOT / conv.OUT_DECISOES
    vistos: dict[tuple[str, str], dict] = {}
    duplicados = 0
    for out in sorted(WORKLISTS.glob("*.out.jsonl")):
        for line in out.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line:
                continue
            rec = json.loads(line)
            clave = rec.get("clave")
            if not clave:
                continue
            chave = (clave, rec.get("pt_atual", ""))
            if chave in vistos:
                duplicados += 1
            vistos[chave] = rec
    with open(destino, "w", encoding="utf-8") as fh:
        for rec in vistos.values():
            fh.write(json.dumps(rec, ensure_ascii=False) + "\n")
    print(f"Consolidadas {len(vistos)} decisões em {destino.name} ({duplicados} duplicados)")
    return len(vistos)


def aplicar() -> int:
    """Regenera os outputs do pipeline já com as decisões consolidadas."""
    return conv.build()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "comando",
        choices=["worklists", "consolidar", "aplicar", "tudo"],
        help="worklists | consolidar | aplicar | tudo",
    )
    parser.add_argument("--seed", type=int, default=42)
    args = parser.parse_args(argv)

    if args.comando in ("worklists", "tudo"):
        print("A gerar worklists...")
        gerar_worklists(seed=args.seed)
    if args.comando in ("consolidar", "tudo"):
        consolidar()
    if args.comando in ("aplicar", "tudo"):
        aplicar()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
