# AGENTS.md — pipeline pt-PT → pt-BR (Ulyses Suite)

## Fluxo de regeneração (obrigatório)
O pipeline **aplica `decisoes_revisao.jsonl` por cima das regras**. Ao alterar
regras, limpar sempre antes, senão o ficheiro de decisões antigo impõe sugestões
velhas e os testes falham com resíduo pt-PT:

```bash
rm -rf worklists __pycache__ && rm -f decisoes_revisao.jsonl
python3 pt_pt_para_pt_br.py        # gera xlsx/csv
python3 orquestrador.py worklists  # worklists TSV + amostra 5%
python3 agentes_runner.py          # JSONL dos 3 agentes (APROVAR determinístico)
python3 orquestrador.py consolidar # funde em decisoes_revisao.jsonl
python3 orquestrador.py aplicar    # regenera xlsx/csv já com decisões
```

## Testes e QA
```bash
python3 -m unittest test_pt_pt_pt_br   # 79 testes
python3 qa_report.py                   # sign-off (7 checks)
```

## Convenções de código
- `_rule(pt, br)` gera padrões com word boundaries + `re.IGNORECASE`;
  `converter()` aplica `_adapt_case` (espelha capitalização) e `_fix_agreement`
  (concordância de género com "acomodação").
- **Ordem importa**: compostos antes de verbos nus ("a carregar" antes de
  "carregar"); flexões antes de regras genéricas ("anulado" antes de "anular").
- Substituição direta `re.sub(padrao, subst, texto)` — nunca
  `re.sub(r"\\1", m.group(1))`, porque `group(1)` pode ser `None` → `TypeError`.
- **Progressivo pt-PT → gerúndio**: `_fix_progressive` trata "está a carregar";
  `_fix_bare_progressive` trata o progressivo sem auxiliar ("A processar…") e
  só na posição inicial e com lista fechada de verbos — uma regra genérica
  "A <palavra>" corromperia nomes próprios ("A Cuiner", "A Icar") e expressões
  preposicionais ("a partir de", "a receber").
- **`CORRECOES_BR`** (auditoria 2026-10): correções cirúrgicas por chave com
  precedência sobre regras e agentes, para pt-PT que escapava às regras
  genéricas — sobretudo "Descargar/Download" traduzido como "Transferir"
  (pt-BR usa "Baixar").
- **Cuidado com NaN**: `str(valor or "")` devolve `"nan"` porque `NaN` é
  truthy; usar coerção explícita com `pd.isna`, senão as linhas `OK` (correção
  vazia) escapam a qualquer auditoria.
- `_protect`/`_restore` guardam `{{...}}`, `{0}`, tags HTML antes das regras.
- Nunca traduzir ES/EN de raiz; nunca tocar em colunas fora de
  `Estado`/`Corrección PT`/`Comentario`.
- `allotment` mantém-se intacto (decisão de glossário B2B) — não reintroduzir
  regra global `allotment → cota`.
- `factura` aparece só em nomes de marca (BSFactura, FacturacionCL) — não mexer.

## Nomes de abas
`SHEET_NAMES_PT_BR` traduz os rótulos das abas **só na escrita dos ficheiros de
saída**; as chaves internas do pipeline (e os testes que usam
`_review_sheet_names`) mantêm os nomes originais.

## Estado de entrega
Branch `feat/pt-br-duvidas-po` (PR #4). O `GITHUB_TOKEN` do ambiente é
read-only/inválido; push/PR usam o token de escrita que o utilizador forneceu.

## Ficheiro de entrega principal
`Revision_traducciones_PT_pt-BR.xlsx` — **o próprio original**, preenchido só nas
três colunas de revisão (`Estado`, `Corrección PT`, `Comentario`). Preserva
abas, cabeçalhos, formatação, larguras e o dropdown `OK,Corregir`.
`_write_preenchido()` emparelha por **ordem de ocorrência** da chave, porque a
chave não é única (a coluna «Clave técnica (no editar)» vem truncada e repete-se
em 1041 linhas de «Entidades de sistema» e 102 de «Mensajes e informes»).

`Estado` usa os rótulos oficiais do dropdown: **`OK`** / **`Corregir`**
(`ESTADO_CORRIGIDO = "Corregir"`; o valor interno `PENDENTE_PO` só aparece se
houver dúvidas abertas, e hoje não há).

## Comentários
`COMENTARIOS` (dict chave→nota) preenche a coluna `Comentario` **apenas** nos
casos polémicos ou ambíguos (allotment, pensão, fatura anulada, tokens técnicos
OnHold/OutOfOrder/OutOfService, IVA→imposto, baixar vs transferir, etc.). Uma
linha sem nota fica vazia de propósito. Vão também para `correcoes_backend.csv`.

## Decisões do PO (sign-off 2026-10)
As 62 linhas que estavam em `PENDENTE_PO` foram fechadas e vivem em
`CORRECOES_PO`. Regras de precedência em `process_sheet`:

1. `CORRECOES_PO` (decisão do PO) — vence aviso da fonte e agentes.
2. `DECISOES` (agentes) — `_resolve_estado` normal.
3. `CORRECOES` — override cirúrgico.
4. `converter()` — regras genéricas.

`FALSE_POSITIVOS_PT_PT` isenta chaves cujo texto aprovado contém substring que
as regras reescreveriam (`INVCVOID` = "Fatura anulada", distinto de
"cancelada" em BR fiscal). Se acrescentares uma decisão dessas, adiciona-a lá,
senão `_tem_residuo_pt_pt` esvazia a linha no ficheiro final.
`DUVIDAS_KEYS`/`DUVIDAS_GRUPOS` estão vazios; `duvidas_PO.xlsx` sai sem linhas.
