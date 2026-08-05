# Project Scope: Verified Candidate-Only Legacy V1

Status: current authority
Last revised: 2026-08-05

## Purpose

Legacy V1 safely generates a non-production candidate workbook from a local legacy `log` workbook. The candidate is built from a separate database copy and is reopened and checked before it is reported. Source and database inputs remain unchanged.

The implementation is deliberately candidate-only. It does not acquire SharePoint data, alter production, or claim Cellpy or Excel qualification.

## V1 acceptance

V1 is accepted only when the controlled candidate workflow can:

- read a local workbook containing `log`;
- retain rows for the fixed approved project set and reject invalid or duplicate included-source IDs;
- append only new IDs to `Slurry` and minimal system-owned fields to `db_table`;
- preserve existing rows and existing manual `b01:b07` values;
- report retained, new, absent, and duplicate IDs; and
- save and reopen the separate candidate workbook successfully.

The public command is:

```powershell
python main.py candidate --source ... --database ... --output ... [--report ...]
```

The exact production database path is rejected. In-place output is rejected. `python main.py` and `python main.py validate` remain disabled with exit code 3, and retired legacy flags remain rejected.

## Explicitly deferred

The following are not V1 acceptance criteria and must not be described as verified:

- SharePoint `c&p` acquisition, authentication, source manifests, and live source anomaly resolution;
- complete `db_table` field/mapping refinement (`db_table` refinement is the next offline step);
- Cellpy installation, qualification, or semantic validation;
- Excel recalculation, desktop Excel automation, and cached-value qualification; and
- production workbook replacement, apply, backup, rollback, or any other production write.

Preserved planning, mapping, source, gate, and safety artifacts for this deferred work are in [docs/deferred_cp/](docs/deferred_cp/README.md). They are historical future-work material, not active V1 scope.

## Evidence boundary

Current evidence is limited to the implementation and candidate tests in `main.py`, `src/candidate_pipeline.py`, `src/contracts.py`, `tests/test_candidate_cli.py`, and `tests/test_candidate_pipeline.py`. No live SharePoint, business workbook, Cellpy, Excel, or production validation is claimed.

## Next plan

Refine `db_table` ownership and mapping offline against representative fixtures, then define separate Excel/Cellpy qualification gates. Production acquisition and replacement require a later, explicitly reviewed scope and implementation change.
