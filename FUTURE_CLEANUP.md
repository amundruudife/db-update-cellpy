# Future cleanup

Status: candidate ledger only. This file does not authorize archive or deletion
work. Correctness findings must be fixed before cleanup changes are considered.

## Rules

- Preserve source workbooks, production workbooks, candidate outputs, manifests,
  credentials, and all other business data.
- Move a module only with its directly coupled tests and fixtures.
- Confirm there are no runtime, test, documentation, packaging, or operator
  consumers before archiving or deleting a path.
- Run `python -m pytest -q` and `git diff --check` after each cleanup batch.
- Keep `docs/deferred_cp/` as historical evidence unless the business owner
  explicitly approves a retention change.

## Candidates

| ID | State | Target | Future action | Required gate |
| --- | --- | --- | --- | --- |
| CL1 | Keep for now | `config.template.json` | Archive outside the active root, then consider deletion | Remove the stale setup instructions that reference it and confirm no external installer or operator workflow consumes it. |
| CL2 | Keep for now | `src/deferred_cp/anomaly_ledger.py` and `tests/deferred_cp/test_anomaly_ledger.py` | Archive together | Gate A anomaly decisions are formally closed or the implementation is explicitly superseded. |
| CL3 | Keep for now | `src/deferred_cp/source_manifest.py` and `tests/deferred_cp/test_source_manifest.py` | Archive together | The accepted source-manifest workflow is integrated into the active acquisition path or explicitly abandoned. |
| CL4 | Keep for now | `src/deferred_cp/field_mapping.py`, `tests/deferred_cp/test_field_mapping.py`, and `tests/deferred_cp/fixtures/gate_b/` | Archive as one Gate B unit | Mapping ownership, the 73/74-column mismatch, formula ranges, and Cellpy behavior are resolved or explicitly superseded. |
| CL5 | Keep for now | `src/deferred_cp/database_inventory.py`, `tests/deferred_cp/test_database_inventory.py`, and `tests/deferred_cp/fixtures/gate_b_inventory/` | Archive as one inventory unit | Gate B topology evidence is accepted, retained elsewhere, and no regeneration path depends on the module. |
| CL6 | Keep for now | `setup.py` | Delete only after replacement | A minimal packaging replacement provides the console entry point and dependency metadata, stale setup side effects/instructions are removed, and install/CLI checks pass. |
| CL7 | Completed | `.tmp_luna_adversarial_*` | Delete reviewer prompt/log scratch files | The delegated review returned its final report and no process consumes the scratch artifacts. |

## Protected or active paths

Do not archive or delete these as cleanup:

- `main.py`, `src/candidate_pipeline.py`, and `src/neware.py`;
- `src/deferred_cp/credentials.py`, `source_acquisition.py`, `snapshot.py`, and
  `source_validation.py`;
- `docs/NEWARE_MAPPING.md`, `PROJECT_SCOPE.md`, and `IMPLEMENTATION_PLAN.md`;
- active candidate, Neware, acquisition, and Cellpy-integration tests;
- `docs/deferred_cp/` and `docs/deferred_cp/gate_b_artifacts/`;
- `source_data/`, `output/`, real workbooks, manifests, and credentials.

`graphify-out/` is generated navigation output. Regenerate it through the
project Graphify workflow; do not treat it as source truth or manually curate
individual generated files.
