"""Auditoria independente: procura resíduo pt-PT no texto final de TODAS as linhas.

Não usa o marcador interno do pipeline (lista fechada). Usa uma lista curada de
termos cujo pt-PT difere CLARAMENTE do pt-BR (sem falsos positivos como
"utilizar"/"acesso"/"ligar", que são válidos em pt-BR).
"""
from __future__ import annotations

import re

import pandas as pd

import pt_pt_para_pt_br as conv

#: termo pt-PT -> equivalente pt-BR (para leitura humana)
TERMOS = {
    "utilizador": "usuário", "utilizadores": "usuários", "ficheiro": "arquivo",
    "ficheiros": "arquivos", "guardar": "salvar", "eliminar": "excluir",
    "transferir": "baixar", "ecrã": "tela", "ecrãs": "telas",
    "telemóvel": "celular", "telemóveis": "celulares", "morada": "endereço",
    "moradas": "endereços", "predefinição": "padrão", "predefinições": "padrões",
    "predefinido": "padrão", "predefinida": "padrão",
    "folha de cálculo": "planilha", "fecho": "fechamento", "secção": "seção",
    "secções": "seções", "receção": "recepção", "montante": "valor",
    "faturação": "faturamento", "planeamento": "planejamento",
    "enviámos": "enviamos", "subscrição": "assinatura", "cadeia": "rede",
    "demasiado": "muito", "demasiada": "muita", "Letónia": "Letônia",
    "Roménia": "Romênia", "Iémen": "Iêmen", "Zimbabué": "Zimbábue",
    "acrónimo": "acrônimo", "género": "gênero", "percentagem": "porcentagem",
    "grelha": "grade", "recolha": "coleta", "gestor de conta": "gerente de contas",
    # descobertos na auditoria de 2026-10
    "connosco": "conosco", "convosco": "com você",
    "pequeno-almoço": "café da manhã", "pequenos-almoços": "cafés da manhã",
    "apelido": "sobrenome", "apelidos": "sobrenomes",
    "contacto": "contato", "contactos": "contatos", "contactar": "contatar",
    "contacte": "contate", "contactem": "contatem", "contactado": "contatado",
    "casa de banho": "banheiro", "casas de banho": "banheiros",
    "registar": "registrar", "registe": "registre", "registem": "registrem",
    "registou": "registrou", "registado": "registrado", "registada": "registrada",
    "registadora": "registradora", "registámos": "registramos",
    # grafia AO90 que o pt-BR não usa
    "acção": "ação", "acções": "ações", "actividade": "atividade",
    "actividades": "atividades", "actual": "atual", "objectivo": "objetivo",
    "objectivos": "objetivos", "director": "diretor",
}

PAT = re.compile(
    r"\b(?:" + "|".join(re.escape(t) for t in sorted(TERMOS, key=len, reverse=True)) + r")\b",
    re.IGNORECASE,
)


def _txt(value: object) -> str:
    """Texto limpo de uma célula; ``None``/NaN viram string vazia.

    Cuidado: ``NaN or ""`` devolve ``NaN`` (é truthy), por isso a coerção tem
    de ser explícita — senão as linhas ``OK`` (correção vazia) nunca entram na
    auditoria.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value)


def main() -> int:
    sheets = pd.read_excel(conv.OUT_COMPLETA, sheet_name=None)
    rows = []
    for name, df in sheets.items():
        for _, r in df.iterrows():
            pt = _txt(r.get("Traducción PT actual"))
            corr = _txt(r.get(conv.COL_CORRECAO)).strip()
            rows.append(
                (
                    name,
                    r.get("Clave") or r.get("Clave técnica (no editar)") or r.get("Código"),
                    r.get(conv.COL_ESTADO),
                    pt,
                    corr,
                    corr if corr else pt,
                )
            )
    a = pd.DataFrame(rows, columns=["hoja", "clave", "estado", "pt", "corr", "final"])
    print(f"Total linhas: {len(a)}")
    print(f"Por estado: {a['estado'].value_counts().to_dict()}")
    print()

    hit = a[a["final"].str.contains(PAT, na=False)]
    print(f"=== Resíduo pt-PT no TEXTO FINAL: {len(hit)} ===")
    if len(hit):
        print(hit["estado"].value_counts().to_dict())
        for _, r in hit.iterrows():
            m = sorted({x.lower() for x in PAT.findall(str(r["final"]))})
            print(f"[{r['estado']}] {str(r['clave'])[:42]:42} | {m}")
            print(f"      {str(r['final'])[:110]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
