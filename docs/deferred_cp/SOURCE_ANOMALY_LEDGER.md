# Gate A source anomaly ledger

Status: provisional local evidence; source correction and live reconciliation still required

Evidence workbook: `C:\Users\IFE13213\Downloads\Cell_Log.xlsx`
Evidence SHA-256: `15787AF0AEACCA6D05C8CF31129EF0F33D7554ED16C511BE656F4C00599694E1`
Inspected sheet: `c&p`, evaluated values, rows 4 through 4590

The updater must not choose a first or last duplicate, merge rows, renumber IDs,
or edit the source workbook. Every entry below therefore remains a hard stop
until the source owner corrects the source or approves a different stable key.

## Duplicate positive IDs

| ID | Source rows | Local comparison | Disposition |
|---:|---:|---|---|
| 5327 | 185, 186 | identical A:S values | source correction required |
| 5569 | 428, 429 | differs in E | source correction required |
| 6087 | 947, 948 | identical A:S values | source correction required |
| 7689 | 2550, 2551 | identical A:S values | source correction required |
| 7780 | 2640, 2644 | differs in L | source correction required |
| 7781 | 2641, 2645 | differs in L | source correction required |
| 9145 | 4006, 4007 | differs in B, C, D, E | source correction required |
| 9146 | 4008, 4009 | differs in B, C, D, E | source correction required |
| 9147 | 4010, 4011 | differs in B, C, D, E | source correction required |
| 9488 | 4352, 4360 | differs in B, C, D, E, F, G, L | source correction required |
| 9489 | 4353, 4361 | differs in B, C, D, E, F, G, L | source correction required |
| 9490 | 4354, 4362 | differs in B, C, D, E, F, G, L | source correction required |
| 9491 | 4355, 4363 | differs in B, C, D, E, F, G, L | source correction required |
| 9492 | 4356, 4364 | differs in B, C, D, E, F, G, L | source correction required |
| 9493 | 4357, 4365 | differs in B, C, D, E, F, G, L, M, N, O, P | source correction required |
| 9494 | 4358, 4366 | differs in B, C, D, E, F, G, L, M, N, O, P | source correction required |
| 9495 | 4359, 4367 | differs in B, C, D, E, F, G, L, M, N, O, P | source correction required |

## Invalid key rows

| Source row | Value | Disposition |
|---:|---|---|
| 1168 | `#REF!` | source correction required |
| 2806 | `#REF!` | source correction required |
| 3979 | `#REF!` | source correction required |
| 4589 | `0` | source correction required |
| 4590 | `0` | source correction required |

## Existing database reconciliation

The inspected production workbook contains database ID `8206`, which is not
present in the provisional local source's positive-ID set. The updater must
abort before the first accepted migration until the source owner explains this
discrepancy and records the decision.

Production evidence SHA-256: `774443DF25A935A55EED8A3BDD45264FE9B3616290CB0C458B88ACFFBEA9DF2C`

## Gate decision

Tasks A3 duplicate/invalid resolution and ID `8206` reconciliation are **not
closed** by this ledger. It is the retained evidence and explicit stop record
needed before a live source comparison and first-run manifest can be accepted.
