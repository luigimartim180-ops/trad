# GUIA — Como usar e adaptar o pipeline pt-PT → pt-BR

Este guia serve para **qualquer pessoa** que precise de converter um ficheiro de
traduções de português de Portugal (pt-PT) para português do Brasil (pt-BR),
seguindo as regras deste projeto. Aplica-se ao ficheiro Ulyses e a **outros
ficheiros do mesmo tipo** (mesma estrutura de abas com chave + texto).

---

## 1. O que o programa faz

Lê um Excel de revisão, converte o **texto pt-PT → pt-BR** com regras
determinísticas e classificações, e escreve três ficheiros de saída. Nunca
traduz do zero a partir do espanhol/inglês: só converte o que já está em pt-PT.

O resultado de cada linha fecha em **um de três estados**:

| Estado | Significado | O que fica na coluna `Corrección PT` |
| --- | --- | --- |
| `OK` | Já é pt-BR válido (cognato, sigla, termo da whitelist). | **Vazia** (nada a alterar). |
| `CORRIGIDO` | Houve conversão pt-PT → pt-BR. | O texto final em pt-BR. |
| `PENDENTE_PO` | Dúvida de glossário ou tradução em falta. | Vazia; vai para `duvidas_PO.xlsx`. |

Regra de ouro: **só** se alteram as colunas `Estado`, `Corrección PT` e
`Comentario`. O texto ES/EN original nunca é regenerado.

---

## 2. Instalação (uma vez)

```bash
pip install pandas openpyxl
```

---

## 3. Uso normal (o ficheiro Ulyses)

```bash
python3 pt_pt_para_pt_br.py
python3 -m unittest test_pt_pt_pt_br   # 57 testes
python3 qa_report.py                   # validações de aceitação
```

Saídas geradas na mesma pasta:

| Ficheiro | Para que serve |
| --- | --- |
| **`revisao_pt_BR_FINAL.xlsx`** | **O ficheiro que se entrega.** Uma única coluna de texto, `Texto final pt-BR`, sem a coluna pt-PT antiga. |
| `revisao_pt_BR_completa.xlsx` | Todas as abas, com `Estado`/`Corrección PT`/`Comentario`. **Para revisão técnica**, não para entrega. |
| `duvidas_PO.xlsx` | Dúvidas que só o Product Owner pode decidir (glossário, regime de refeição, etc.). |
| `correcoes_backend.csv` | Lista `Clave;Estado;Corrección PT;Comentario` só das linhas `CORRIGIDO`, para aplicar no backend. |

### Qual ficheiro entregar

Entrega **`revisao_pt_BR_FINAL.xlsx`**. Ele já não tem a coluna
`Traducción PT actual` (a antiga, em pt-PT): o texto final está todo numa única
coluna **`Texto final pt-BR`**. Assim não há forma de se entregar, por engano,
o pt-PT antigo. As linhas sem tradução ou com questão de glossário em aberto
ficam com a célula **vazia** e constam em `duvidas_PO.xlsx`.

Os três (`revisao_pt_BR_completa.xlsx`, `correcoes_backend.csv`, `duvidas_PO.xlsx`)
continuam a ser gerados, para auditoria e para o backend.

### Onde baixar

Os ficheiros ficam na raiz do repositório/projeto, na mesma pasta do
`pt_pt_para_pt_br.py`. No GitHub, abre o repositório
[`luigimartim180-ops/trad`](https://github.com/luigimartim180-ops/trad), entra
no branch `main` e descarrega cada ficheiro com o botão **Download raw file**.

---

## 4. Usar noutro ficheiro (reutilização)

O programa aceita o ficheiro de entrada e a pasta de saída por argumento:

```bash
python3 pt_pt_para_pt_br.py outro_ficheiro.xlsx -o saida_outro/
```

Sem argumentos, usa `Revision_traducciones_PT.xlsx` na própria pasta.

### O que o novo ficheiro precisa de ter

1. Uma ou mais **abas de revisão**, cada uma com:
   * uma coluna de chave (`Clave`, `Clave técnica (no editar)` ou `Código`);
   * uma coluna de texto pt-PT (ex.: `Traducción PT actual`, `Texto PT`).
2. As abas sem chave/texto (ex.: `Instrucciones`) são **ignoradas** sozinhas.
3. O cabeçalho pode estar **depois** de linhas de título: o programa procura a
   linha que contém `Clave`/`Código`.

Se a estrutura for diferente, ajusta em `pt_pt_para_pt_br.py`:

```python
KEY_COLUMNS = ["Clave", "Clave técnica (no editar)", "Código"]
# e a lista de colunas de texto PT:
```

---

## 5. Aprender as regras (essencial)

Todas as conversões estão numa única lista: **`SUBSTITUICOES`** em
`pt_pt_para_pt_br.py`. Cada entrada é `(padrão_regex, substituição)`.

### Regras críticas de ordem

* **Compostas primeiro.** `a carregar` → `carregando` tem de vir **antes** de
  `carregar` → `enviar`, senão o gerúndio é destruído.
* **Plurais antes de singulares.** `utilizadores` → `usuários` antes de
  `utilizador` → `usuário`.
* **Fronteiras de palavra** (`\b`) em todas as regras, para não corromper
  palavras maiores (ex.: `para carregar` não é afetado).

### Contexto importa (falsos positivos a evitar)

| Regra | Só disparar quando… | Nunca disparar em… |
| --- | --- | --- |
| `carregar` → `enviar` | é carregamento de ficheiro/imagem/documento/botão | paginação ("carregar mais"), erro de carga |
| `transferir` → `baixar` | é descarregar um objeto (ficheiro, fatura) | movimento de conta/folio, migração de dados |
| `cadeia` → `rede` | hotelaria/B2B | cadeia de custódia, cadeia produtiva |
| `eliminar` → `excluir` | ação de UI | sentido jurídico ("eliminar definitivamente") |
| `montante` → `valor` | uso comum | contrato ("montante indemnizatório") |

### Placeholders são sagrados

`{{...}}`, `{0}`, `{1}` e tags HTML (`<br>`, `<strong>`) são **protegidos**
antes das substituições e repostos depois. Nenhuma regra os pode alterar. O QA
verifica a integridade multiset.

### Nunca traduzir do zero

Se não houver pt-PT para converter (texto vazio, "Sin traducción PT"), a linha
fica `PENDENTE_PO` — nunca se inventa. Dúvidas de glossário também vão para o PO.

---

## 6. Adicionar uma regra nova (receita)

1. Abre `pt_pt_para_pt_br.py` e encontra `SUBSTITUICOES`.
2. Acrescenta a entrada **na posição correta** (compostas antes das simples):

```python
SUBSTITUICOES = [
    # ... compostas ...
    (re.compile(r"\bmorada\b", re.I), "endereço"),
    (re.compile(r"\bpredefinição\b", re.I), "padrão"),
    # ... simples ...
]
```

3. Usa sempre `r"\bpalavra\b"` e `re.I` (ignora maiúsculas).
4. Adiciona um teste em `test_pt_pt_pt_br.py` com um caso real.
5. Corre `python3 -m unittest test_pt_pt_pt_br` e `python3 qa_report.py`.

---

## 7. Fluxo multi-agente (revisão em qualidade, opcional)

Para fechar a fila sem traduzir ES/EN do zero, há 3 papéis descritos em
`agentes.md` (Auditor → Redator → QA):

```bash
python3 orquestrador.py worklists    # gera worklists dos agentes
python3 agentes_runner.py            # decisões dos agentes -> *.out.jsonl
python3 orquestrador.py consolidar   # funde em decisoes_revisao.jsonl
python3 orquestrador.py aplicar      # reescreve as saídas com as decisões
python3 qa_report.py                 # validações de aceitação (PASS/FAIL)
```

---

## 8. Verificação antes de entregar

* `python3 -m unittest test_pt_pt_pt_br` → **OK** (0 falhas).
* `python3 qa_report.py` → **RESULTADO GLOBAL: PASS ✅**.
* `python3 auditoria_pt_pt.py` → confirma que **nenhuma** linha (incluindo as
  `OK`, que nunca foram tocadas) tem resíduo pt-PT.
* Nenhuma linha `OK` com `Corrección PT` preenchida.
* Nenhum placeholder corrompido.

### Importante: ordem de execução depois de mudar uma regra

O pipeline tem uma dependência que engana. O `pt_pt_para_pt_br.py` **aplica**
o `decisoes_revisao.jsonl` (decisões dos agentes) por cima das regras. Se
acrescentar uma regra e voltar a correr só o pipeline, o ficheiro de decisões
**antigo** volta a impor as sugestões velhas — parece que a regra não funciona.

Sempre assim, nesta ordem:

```bash
rm -rf worklists __pycache__ && rm -f decisoes_revisao.jsonl   # 1. limpar estado antigo
python3 pt_pt_para_pt_br.py                                    # 2. gerar SEM decisões antigas
python3 auditoria_pt_pt.py                                     # 3. verificar resíduo
python3 orquestrador.py worklists                              # 4. só agora regenerar decisões
python3 agentes_runner.py
python3 orquestrador.py consolidar
python3 orquestrador.py aplicar                                # 5. reaplicar
python3 qa_report.py                                           # 6. PASS/FAIL
python3 -m unittest test_pt_pt_pt_br                           # 7. testes
```

---

## 9. Ficheiros do projeto (mapa rápido)

| Ficheiro | Papel |
| --- | --- |
| `pt_pt_para_pt_br.py` | Pipeline principal + regras (`SUBSTITUICOES`). |
| `glossario_pt_BR.md` | Glossário oficial (fonte de verdade dos termos). |
| `agentes.md` | Contratos dos 3 agentes (prompts + formato JSONL). |
| `orquestrador.py` | Worklists, consolidação e aplicação das decisões. |
| `agentes_runner.py` | Gera as decisões dos agentes em JSONL. |
| `qa_report.py` | Validações de aceitação e relatório de sign-off. |
| `auditoria_pt_pt.py` | Auditoria independente de resíduo pt-PT (todas as linhas). |
| `test_pt_pt_pt_br.py` | Testes unitários com casos reais. |
| `README.md` | Visão geral técnica do pipeline. |

---

*Documento de apoio criado por um agente de IA (OpenHands) em nome do
utilizador. As regras vinculativas estão em `glossario_pt_BR.md`.*
