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
python3 -m unittest test_pt_pt_pt_br   # 73 testes
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
Branch `feat/pt-br-conversao`. O `GITHUB_TOKEN` do ambiente é read-only/inválido
(401 Bad credentials) → push/PR exigem credenciais de escrita fornecidas pelo
utilizador.
