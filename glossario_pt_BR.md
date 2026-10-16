# Glossário pt-BR oficial — Ulyses Suite

**Vinculativo para futuras traduções.** Fonte de verdade para conversão pt-PT → pt-BR.

## 1. Glossário de termos

| EN | pt-BR | pt-PT (histórico) | Notas |
| --- | --- | --- | --- |
| user | usuário | utilizador | UI, e-mails, docs |
| password | senha | palavra-passe | Nunca "palavra-passe" em BR |
| file | arquivo | ficheiro | — |
| to upload | fazer upload / enviar | carregar | "Carregar" só como "carregar dados" |
| to download | baixar | transferir | — |
| to save | salvar | guardar | "Guardar" em pt-BR é "guardar (roupa, objeto)" |
| to delete | excluir | eliminar | "Excluir" é o verbo UI |
| screen | tela | ecrã | — |
| mobile phone | celular | telemóvel | — |
| contact | contato | contacto | AO90 em BR |
| register | registro | registo | — |
| activity | atividade | actividade | AO90 |
| action | ação | acção | AO90 |
| goal / aim | objetivo | objectivo | AO90 |
| boarding / meal plan | plano de refeição (preferir) / pensão (técnico) | pensão | Ver dúvida §3 |
| allotment | cota (UI) / allotment (contratos) | allotment | Ver dúvida §3 |
| OnHold | Em espera | OnHold | Só traduzir se aparecer na UI |
| OutOfOrder | Fora de ordem | OutOfOrder | idem |
| OutOfService | Fora de serviço | OutOfService | idem |
| accommodation | acomodação / alojamento | alojamento | "Acomodação" mais BR |
| room | quarto | quarto | — |
| property | propriedade | propriedade | — |
| invoice | fatura | fatura | AO90 |
| check-in | check-in | check-in | Manter |
| check-out | check-out | check-out | Manter |
| guest | hóspede | hóspede | — |
| booking / reservation | reserva | reserva | — |

## 2. Regras gerais (implementadas em `pt_pt_para_pt_br.py`)

- **Compostas primeiro**: `a carregar` → `carregando` antes de `carregar` → `enviar`.
- **Plurais antes de singulares**: `utilizadores` → `usuários`, `contactos` → `contatos`.
- **Fronteiras de palavra** em todas as regras (`\b`), para não corromper `para carregar`.
- **Placeholders protegidos**: `{{...}}`, `{0}`, `{1}` e tags HTML (`<br>`, `<strong>`).
- **AO90 → BR**: `contacto`→`contato`, `actividade`→`atividade`, `acção`→`ação`,
  `objetivo`, `protecção`→`proteção`, etc.
- **`allotment` mantém-se** (decisão B2B); sem regra global `allotment` → `cota`.
- **`transferir` mantém-se** nas ações de conta/folio; só os botões de download
  viram `Baixar` (correções cirúrgicas por chave).

## 3. Dúvidas para o PO

Cada linha precisa de decisão. As chaves correspondentes ficam com
`Estado=REVISAR` e `Corrección PT` vazia — **nunca convertidas automaticamente**.
Estão listadas em `duvidas_PO.xlsx` (coluna `Grupo`).

### Dúvida 1 — "pensão" (board / meal plan)
Contexto: `apps.board.*`, `apps.rate.titleSection.boardAllowed`, `core.label.board*`, `PropertyClassType` Code 9.

- O que está hoje: "pensão" (pt-PT). EN: board, full board, half board.
- Problema: "pensão" em pt-BR é casa de hóspedes, não plano de refeição.
- Opções: **(a)** "plano de refeição"; **(b)** manter "pensão" como jargão interno;
  **(c)** "pensão completa" / "meia pensão" (usado pelas cadeias internacionais em BR).
- **Recomendação: (c)**; trocar "pensão" isolada por "plano de refeição".

### Dúvida 2 — "allotment"
Contexto: `apps.allotment.*`, `core.label.allotment*`, `apps.chainAllotment.*`.

- O que está hoje: "allotment" (anglicismo). EN: allotment.
- Opções: **(a)** "cota" (Microsoft Language Portal); **(b)** manter "allotment".
- **Recomendação: (b)** para B2B com operadores estrangeiros; (a) se a UI for para gestor júnior.

### Dúvida 3 — OnHold / OutOfOrder / OutOfService
Contexto: `apps.chainSpaceLockType.radio.*`.

- O que está hoje: "OnHold (…)", "OutOfOrder (…)", "OutOfService (…)".
- Opções: **(a)** traduzir ("Em espera", "Fora de ordem", "Fora de serviço");
  **(b)** manter como código.
- **Recomendação: (a)** se aparecer na UI; (b) se for só em logs.

### Dúvida 4 — `apps.aiAssistant.label.hotel`
Contexto: `apps.aiAssistant.label.hotel` = "Hotel" (ES) / "Property" (EN).

- Opções: **(a)** "Hotel" (segue o ES); **(b)** "Propriedade" (segue o EN).
- **Recomendação: (a)** se o público é hoteleiro; (b) se é multi-property/CRS.

## 4. Correções cirúrgicas (fonte oficial BR)

| Chave | pt-PT | pt-BR | Fonte |
| --- | --- | --- | --- |
| `card.hojaRegistro` | Folha de registo | Ficha de registro de hóspedes (FNRH) | FNRH — Polícia Federal/Embratur |
| `core.button.settlement` | Liquidação | Cobranças automáticas | ES "Cargos automáticos" |
| `invoice.direction.label` | Morada | Endereço | — |

## 5. Termos protegidos (whitelist)

Termos que podem legitimamente ficar iguais a ES/EN, sem necessidade de revisão:
Ok, Overbooking, RevPar, JSON, Check-in, Check-out, Day use, Voucher, Pax,
Fiscal, Central, Empresa, Canal.
