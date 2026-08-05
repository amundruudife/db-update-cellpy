# Implementation Plan: Candidate-Only Legacy V1

## Current state — verified V1

The current implementation is a controlled local candidate builder. It filters the legacy `log` sheet to the approved projects, preserves existing `Slurry` and `db_table` content, appends new IDs with minimal system values, writes a separate candidate and optional report, and reopens the candidate for verification. Production paths and legacy update flags are unavailable.

Verification basis: `main.py`, `src/candidate_pipeline.py`, `src/contracts.py`, the candidate CLI tests, and the candidate pipeline tests.

## Next: offline `db_table` refinement

Define and test ownership for the remaining `db_table` columns, including source mapping, formula behavior, units, null policy, and preservation rules. Use representative non-production fixtures and keep candidate output separate from production. This step must not enable SharePoint acquisition, Excel automation, Cellpy validation, or production writes.

## Deferred production work

After mapping refinement, separately qualify live SharePoint `c&p` acquisition, Excel recalculation, and Cellpy reads. Only a later reviewed change may add candidate approval, backup/rollback, and production replacement/apply. The preserved source, mapping, gate, and safety materials are archived in [docs/deferred_cp/](docs/deferred_cp/README.md).
