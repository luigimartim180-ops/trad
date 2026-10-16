# trad — pipeline pt-PT → pt-BR (Ulyses)

Conversão automática do ficheiro de revisão de traduções `Revision_traducciones_PT.xlsx`
(pt-PT → pt-BR), sem intervenção manual.

## Ficheiros

| Ficheiro | Descrição |
| --- | --- |
| `pt_pt_para_pt_br.py` | Pipeline principal (lê o xlsx, converte, escreve as saídas). |
| `glossario_pt_BR.md` | Glossário oficial pt-BR (vinculativo) + dúvidas para o PO. |
| `test_pt_pt_pt_br.py` | Testes unitários com casos reais do ficheiro (46 testes). |
| `revisao_pt_BR_completa.xlsx` | Saída: todas as abas (Estado / Corrección PT / Comentario / Prioridade). |
| `revisao_pt_BR_REVISAR.xlsx` | Saída: **só as linhas REVISAR**, para percorrer mais depressa. |
| `duvidas_PO.xlsx` | Saída: dúvidas para o PO (coluna `Grupo`), nunca convertidas. |
| `correcoes_backend.csv` | Saída: `Clave;Estado;Corrección PT;Comentario;Prioridade` (só não-OK). |

## Utilização

```bash
pip install pandas openpyxl
python3 pt_pt_para_pt_br.py            # gera as 4 saídas
python3 -m unittest test_pt_pt_pt_br   # 46 testes
```

## Prioridades (coluna `Prioridade`)

| Prioridade | Significado |
| --- | --- |
| `P0 - dúvida PO` | Dúvida aberta (§3 do glossário) — nunca convertida. |
| `P1 - decisão manual` | Sem sugestão automática (texto igual a ES/EN, sem tradução PT). |
| `P2 - validar sugestão` | Conversão automática sugerida — só validar. |
| `Fechado` | Linha revista e fechada pelo revisor (correção aplicada ou termo já válido). |

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
* **`allotment` mantém-se** (decisão de glossário B2B; sem regra global `→ cota`).
* **Placeholders protegidos**: `{{...}}`, `{0}`, `{1}` e tags HTML (`<br>`,
  `<strong>`) são guardados antes das substituições e repostos depois.
* **Estado**: `OK` ⇒ `Corrección PT` vazia; aviso `Igual a ES/EN` ⇒ `REVISAR`
  (exceto whitelist: Ok, Overbooking, RevPar, JSON, Check-in, Check-out,
  Day use, Voucher, Pax, Fiscal, Central, Empresa, Canal); `Sin traducción PT`
  ⇒ `REVISAR`.
* **Colunas mortas** `A1_Léxico`..`A5_Consistência` removidas (quando presentes).
* **Correções cirúrgicas**: `card.hojaRegistro` → `Ficha de registro de hóspedes
  (FNRH)`; `core.button.settlement` → `Cobranças automáticas`;
  `invoice.direction.label` → `Endereço`.
* **Dúvidas** (`DUVIDAS_KEYS`) vão apenas para `duvidas_PO.xlsx` e nunca são
  convertidas.
