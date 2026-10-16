# trad — pipeline pt-PT → pt-BR (Ulyses)

Conversão automática do ficheiro de revisão de traduções `Revision_traducciones_PT.xlsx`
(pt-PT → pt-BR), sem intervenção manual.

## Ficheiros

| Ficheiro | Descrição |
| --- | --- |
| `pt_pt_para_pt_br.py` | Pipeline principal (lê o xlsx, converte, escreve as saídas). |
| `glossario_pt_BR.md` | Glossário oficial pt-BR (vinculativo) + dúvidas para o PO. |
| `test_pt_pt_pt_br.py` | Testes unitários com casos reais do ficheiro (55 testes). |
| `agentes.md` | Contratos dos 3 agentes de revisão (Auditor, Redator, QA). |
| `orquestrador.py` | Distribui worklists, consolida JSONL e aplica as decisões. |
| `agentes_runner.py` | Materializa as decisões dos agentes em JSONL (worklists → `.out.jsonl`). |
| `qa_report.py` | Relatório de sign-off com as validações de aceitação. |
| `revisao_pt_BR_completa.xlsx` | Saída: todas as abas (Estado / Corrección PT / Comentario / Prioridade). |
| `duvidas_PO.xlsx` | Saída: dúvidas para o PO (coluna `Grupo`), nunca convertidas. |
| `correcoes_backend.csv` | Saída: `Clave;Estado;Corrección PT;Comentario` (só `CORRIGIDO`). |
| `decisoes_revisao.jsonl` | Decisões validadas pelos agentes, aplicadas ao texto final. |

## Utilização

```bash
pip install pandas openpyxl
python3 pt_pt_para_pt_br.py            # gera as 3 saídas + contagem por estado
python3 -m unittest test_pt_pt_pt_br   # 55 testes
```

### Fluxo multi-agente (opcional, fecha a fila sem traduzir ES/EN)

```bash
python3 orquestrador.py worklists    # gera as worklists dos Agentes 1 e 2
python3 agentes_runner.py            # decisões dos agentes -> worklists/*.out.jsonl
python3 orquestrador.py consolidar   # funde em decisoes_revisao.jsonl
python3 orquestrador.py aplicar      # reescreve as saídas com as decisões
python3 qa_report.py                 # validações de aceitação (PASS/FAIL)
```

## Estados (coluna `Estado`)

| Estado | Significado |
| --- | --- |
| `OK` | Já é pt-BR válido (cognato, sigla, termo da whitelist). `Corrección PT` fica **vazia**. |
| `CORRIGIDO` | Conversão determinística aplicada (ou decisão de agente validada). |
| `PENDENTE_PO` | Dúvida de glossário ou tradução em falta — nunca convertida; sai em `duvidas_PO.xlsx`. |

O estado `REVISAR` foi eliminado: cada linha fecha em `OK`, `CORRIGIDO` ou
`PENDENTE_PO`. A coluna `Prioridade` acompanha (`P2 - validar sugestão`,
`P1 - decisão manual`, `P0 - dúvida PO`, `Fechado`).

## Estrutura do ficheiro (detetada automaticamente)

* `Instrucciones` — aba informativa, ignorada.
* `Frontend (pantallas)` — chave em `Clave`, texto pt-PT em `Traducción PT actual`.
* `Entidades de sistema` — chave em `Clave técnica (no editar)`.
* `Mensajes e informes` — chave em `Clave`.
* O cabeçalho é localizado procurando `Clave`/`Código` (há linhas de título antes).

As abas de revisão são identificadas por terem uma coluna de chave (`Clave` ou
`Clave técnica (no editar)`) e uma coluna de texto pt-PT. Só são escritas as
colunas `Estado`, `Corrección PT` e `Comentario`; o texto ES/EN nunca é regenerado.

## Regras implementadas

* **Compostas primeiro**: `a carregar` → `carregando` antes de `carregar` → `enviar`.
  Todas as regras usam fronteiras de palavra (`\b`), pelo que `para carregar` não é
  afetado e `Utilizadores` → `Usuários` (plural antes do singular).
* **`carregar` → `enviar` só no sentido de upload** (seguido de ficheiro, imagem,
  documento, planilha, modelo, Excel ou isolado como rótulo de botão). Paginação
  (`carregar mais`, `recarregar`) e erros de carregamento (`não foi possível
  carregar`) mantêm o verbo, válido em pt-BR.
* **`transferir` → `baixar` só em objetos descarregáveis** (ficheiro, fatura,
  log, Excel…). Movimentos de conta/folio e migração de dados mantêm `transferir`.
* **`allotment` mantém-se** (decisão de glossário B2B; sem regra global `→ cota`).
* **Placeholders protegidos**: `{{...}}`, `{0}`, `{1}` e tags HTML (`<br>`,
  `<strong>`) são guardados antes das substituições e repostos depois.
* **Estado**: `OK` ⇒ `Corrección PT` vazia. O aviso `Igual a ES/EN` é
  **falso positivo** para códigos, siglas, números, nomes de produto,
  `lorem ipsum`/`//TODO` e cognatos escritos igual em ES/PT/pt-BR (Total, Menu,
  Grupo, Cliente…), pelo que essas linhas fecham `OK` automaticamente; ficam
  `PENDENTE_PO` as que contêm grafia/vocabulário pt-PT (`_PT_PT_MARKER`), as de
  `NOT_OK_IGUAL_ES_EN` (`Acrónimo`, `Género`…) ou as que têm correção sugerida.
  `Sin traducción PT` (texto PT vazio) ⇒ `PENDENTE_PO`.
* **Colunas mortas** `A1_Léxico`..`A5_Consistência` removidas (quando presentes).
* **Correções cirúrgicas**: `card.hojaRegistro` → `Ficha de registro de hóspedes
  (FNRH)`; `core.button.settlement` → `Cobranças automáticas`;
  `invoice.direction.label` → `Endereço`.
* **Dúvidas** (`DUVIDAS_KEYS`) vão apenas para `duvidas_PO.xlsx` e nunca são
  convertidas. As linhas sem texto PT na fonte não são traduzidas do zero.
* **Testes**: 55 testes unitários com casos reais do ficheiro (placeholders,
  progressivo `a carregar`, `tem a certeza`, `enviámos`, `5000 linhas`,
  `Allotment não deduzido` intacto, upload vs. paginação).

## Critérios de aceitação (verificados por `qa_report.py`)

* Zero `TypeError` na conversão (sem `re.sub(r"\1", m.group(1)…)`).
* Nenhuma linha `OK` com `Corrección PT` preenchida.
* Placeholders `{{...}}`, `{0}`, `{1}` e tags HTML íntegros.
* Dúvidas do PO intocadas; contagem final por estado impressa no stdout.
