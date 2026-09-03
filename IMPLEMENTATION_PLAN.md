# Implementation Plan: Gated Source Staging and Candidate V1

## Current state — gated source staging and verified V1

The current implementation has two non-production paths. `python main.py acquire --account ...` resolves the approved SharePoint site/default-drive item, reads the fixed `c&p` range through a non-persistent Graph workbook session, and stages validated values at `source_data/Cell_Log_CP.xlsx`. The candidate builder filters the legacy `log` sheet to the approved projects, validates exact Slurry/db_table/Neware identity reconciliation, preserves existing content, appends new IDs with minimal system values, writes a separate candidate and optional report, and reopens the candidate for verification. Staged artifacts publish in ordered restart-recoverable replacements; this is not same-transaction or multi-file/power-loss atomicity. The opt-in `candidate --cellpy-ready` mode also writes the example's 12 formulas for new `db_table` rows, recalculates the temporary candidate through Windows Excel/COM (pywin32>=312), and verifies cached values before replacement. Production paths and legacy update flags are unavailable.

Verification basis: `main.py`, `src/candidate_pipeline.py`, `src/deferred_cp/source_acquisition.py`, `src/deferred_cp/snapshot.py`, `src/contracts.py`, the acquisition/provider tests, and the candidate tests. Live SharePoint execution and first-run approval remain open.

## Cellpy-ready boundary

The formula templates are limited to the example's `A,C,O,P,Q,R,S,U,V,Z,AA,AG` columns. New manual/preservation-only fields remain blank, and existing formulas and `b01:b07` values are not rewritten. Excel/COM failure removes the temporary candidate and leaves source and database inputs unchanged. Missing pywin32 produces an actionable prerequisite error. This is a cached-value/database-reader gate, not full semantic Cellpy qualification.

## Deferred production work

Separately qualify live SharePoint `c&p` acquisition, the operator's Excel/Cellpy environment, and representative business-workbook semantics. Only a later reviewed change may add candidate approval, backup/rollback, and production replacement/apply. The preserved source, mapping, gate, and safety materials are archived in [docs/deferred_cp/](docs/deferred_cp/README.md).
