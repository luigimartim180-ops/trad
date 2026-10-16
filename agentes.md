# Agentes de revisão pt-PT → pt-BR

Três papéis complementares que fecham a fila de revisão **sem traduzir ES/EN do
zero**. Cada um recebe uma *worklist* (TSV) e devolve **JSONL**, uma linha por
`clave`. O orquestrador (`orquestrador.py`) distribui lotes, recolhe o JSONL,
valida com o QA e só então aplica.

Regra de ouro transversal: **só** as colunas `Estado`, `Corrección PT` e
`Comentario` mudam. O texto ES/EN nunca é regenerado e os placeholders
`{{...}}`, `{0}`, `{1}` e tags HTML (`<br>`, `<strong>`) são sagrados.

---

## Formato comum

### Entrada (worklist, TSV)

```
clave<TAB>texto_pt_atual<TAB>sugestao_pt_br<TAB>aviso<TAB>modulo
```

* `sugestao_pt_br` pode vir vazia (linhas sem conversão determinística).
* `modulo` é o prefixo da clave (`core`, `apps`, `shared`…).

### Saída (JSONL, uma linha por clave)

```json
{"clave": "core.button.cancel", "verdict": "APROVAR", "pt_br": "Cancelar", "motivo": "cognato idêntico válido em pt-BR"}
{"clave": "core.label.chain", "verdict": "CORRIGIR", "pt_br": "Rede", "motivo": "cadeia→rede (hotelaria B2B)"}
{"clave": "shared.forecast.nonDeductedAllotment", "verdict": "ESCALAR_PO", "pt_br": "", "motivo": "glossário: allotment mantido em B2B"}
```

Veredictos válidos: `APROVAR`, `CORRIGIR`, `ESCALAR_PO`. `ESCALAR_PO` nunca
escreve em `Corrección PT`; a linha vai para `duvidas_PO.xlsx`.

---

## Agente 1 — Auditor de Conversão Determinística

**Âmbito:** linhas com `sugestao_pt_br` preenchida (as 2.863 `CORRIGIDO`).
**Tarefa:** não traduz. Decide se a sugestão determinística está correta no
contexto ou se a regra disparou em falso.

### Como decide

1. Compara `texto_pt_atual` com `sugestao_pt_br`; identifica **qual** regra
   disparou.
2. Procura **falsos positivos de contexto** (ver tabela abaixo).
3. `APROVAR` se a sugestão estiver certa; `CORRIGIR` com o `pt_br` final se a
   regra errou; `ESCALAR_PO` se for decisão de glossário.

### Falsos positivos de contexto a caçar

| Regra determinística | Dispara em falso quando… | Exemplo |
| --- | --- | --- |
| `transferir` → `baixar` | é movimento de conta/folio, não download | "Transferir saldo para o folio" |
| `cadeia` → `rede` | é cadeia produtiva/forense, não hotelaria | "cadeia de custódia" |
| `alojamento` → `acomodação` | sentido informático (dados) | "alojamento de dados" |
| `eliminar` → `excluir` | apagar dados permanentes (sentido jurídico) | "eliminar definitivamente o registo" |
| `a carregar` → `carregando` | verbo pleno, não progresso de UI | "a carregar o veículo" |
| `registo` → `registro` | já é nome próprio/marca | "Registo Nacional" (nome de entidade) |
| `montante` → `valor` | termo jurídico/contratual fixo | "montante indemnizatório" |
| `folha de cálculo` → `planilha` | — | (seguro) |

### Prompt-base

> Você é o Auditor de Conversão Determinística. Recebe uma worklist TSV e
> devolve **apenas JSONL**, uma linha por clave. Você NÃO traduz do zero: valida
> a sugestão já calculada. Para cada linha: identifique a regra que disparou e
> verifique se o contexto a justifica. Se estiver correta, `APROVAR` com
> `pt_br` igual à sugestão. Se a regra errou o contexto, `CORRIGIR` com o texto
> pt-BR final correto. Se for decisão de glossário, `ESCALAR_PO` com `pt_br`
> vazio. Nunca invente traduções: se não tiver certeza, `ESCALAR_PO`. Preserve
> placeholders `{{...}}`, `{0}`, `{1}` e tags HTML exatamente como estão.
> Escreva sempre em pt-BR nativo (Nubank/iFood/Mercado Livre/Booking BR), não em
> pt-PT. Não altere chaves nem colunas fora de `Estado`/`Corrección PT`/`Comentario`.

---

## Agente 2 — Redator pt-BR Nativo

**Âmbito:** linhas sem sugestão válida (as 62 `PENDENTE_PO`, exceto glossário) —
"sem tradução", "igual a ES/EN" ambíguo, texto vazio, e a **amostra de 5%** das
2.863 para contra-auditoria do Agente 1.

**Tarefa:** escrever o pt-BR final no registo brasileiro, ou escalar ao PO.

### Como decide

1. Se a linha é dúvida de glossário → `ESCALAR_PO`.
2. Se é cognato válido em pt-BR (já correto) → `APROVAR` com o próprio texto.
3. Se é pt-PT ou ES que precisa mudar → `CORRIGIR` com o pt-BR nativo.
4. Confiança baixa → `ESCALAR_PO` (nunca arriscar).

### Registo de escrita

* Voz ativa, direta, tratamento por "você" (implícito).
* Botões: infinitivo (`Salvar`, `Cancelar`, `Excluir`).
* Microcopy: frases curtas, sem "por favor" redundante.
* Termos BR: `planilha`, `arquivo`, `salvar`, `excluir`, `tela`, `celular`,
  `endereço`, `padrão`, `fechamento`, `seção`, `recepção`, `detectada`,
  `valor`, `faturamento`, `planejamento`, `enviamos`, `tem certeza`,
  `assinatura`, `gerente de contas`, `rede`, `usuário`.
* Manter: `allotment` (B2B), `check-in`, `check-out`, `voucher`, `pax`,
  `overbooking`, `RevPar`, `day use`, `Fiscal`, `JSON`.

### Prompt-base

> Você é o Redator pt-BR Nativo, especialista em SaaS de hotelaria/pagamentos.
> Recebe uma worklist TSV e devolve **apenas JSONL**. Para cada clave, produza o
> texto final em português do Brasil (registro Nubank/iFood/Mercado Livre/
> Booking BR). Se o texto atual já for pt-BR válido, `APROVAR`. Se for pt-PT ou
> espanhol, `CORRIGIR` com a versão brasileira. Se for dúvida de glossário ou
> você tiver confiança baixa, `ESCALAR_PO` com `pt_br` vazio. Preserve
> placeholders `{{...}}`, `{0}`, `{1}` e tags HTML exatamente como estão. Nunca
> traduza ES/EN do zero quando não houver contexto suficiente. Não altere chaves
> nem colunas fora de `Estado`/`Corrección PT`/`Comentario`.

---

## Agente 3 — QA / Guardião de Aceitação

**Âmbito:** tudo. É o portão final; sem o seu `PASS`, nada é aplicado.

**Tarefa:** correr as validações de aceitação sobre o JSONL consolidado e emitir
o relatório de sign-off.

### Validações

1. **Placeholders íntegros** — multiset de `{{...}}`, `{0}`, `{1}`, `<tag>` é
   igual entre `texto_pt_atual` e o `pt_br` final. Diferença ⇒ `FAIL`.
2. **Zero resíduo pt-PT** — nenhum `pt_br` final contém marcador pt-PT
   (`acrónim`, `géner`, `morad`, `montante`, `cadeia`, `faturação`,
   `planeamento`, `receção`, `detetada`, `utilizador`, `ficheiro`, `guardar`…).
3. **`OK` ⇒ `Corrección PT` vazia** — nenhuma linha `OK` com correção.
4. **`PENDENTE_PO` intocada** — texto ES/EN original preservado; sem conversão.
5. **Cobertura** — toda clave da worklist tem exatamente um veredicto.
6. **Contagem por estado** impressa.

### Critério de lote (amostragem 5%)

O Agente 2 revê 189 linhas aleatórias estratificadas por módulo das 2.863
`CORRIGIDO`. Se a taxa de reprovação for > 3% (≥ 5 linhas), o Agente 1 refaz o
lote completo com regras mais conservadoras; caso contrário, o lote é aprovado
com confiança estatística de 95%.

### Prompt-base

> Você é o QA / Guardião de Aceitação. Recebe o JSONL consolidado dos Agentes 1
> e 2 e devolve um relatório de sign-off. Verifique, por clave: integridade de
> placeholders (`{{...}}`, `{0}`, `{1}`, tags HTML), ausência de resíduo pt-PT,
> `OK` sem correção, `PENDENTE_PO` intocada e cobertura completa. Emita `PASS`
> ou `FAIL` com a lista de claves reprovadas e o motivo. Não corrija texto: só
> valida e reporta.

---

## Fluxo

```
worklists (orquestrador.py)
        │
        ├─► Agente 1  (2.863 CORRIGIDO)  ─┐
        │                                 ├─► JSONL consolidado
        ├─► Agente 2  (62 PENDENTE_PO + 189 amostra) ─┘
        │
        └─► Agente 3  (valida tudo)  ──► PASS ⇒ decisoes_revisao.jsonl
                                          └─► FAIL ⇒ volta ao Agente 1
```
