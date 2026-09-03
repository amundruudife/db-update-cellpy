# Project Scope: Gated Source Staging and Candidate V1

Status: current authority
Last revised: 2026-08-05

## Purpose

The repository safely stages a validated local mirror of the approved SharePoint `c&p` source and generates a non-production candidate workbook from a local legacy `log` workbook. Source acquisition uses a non-persistent read session and replaces the local snapshot only after identity, range cardinality, and values validation. Candidate publication is ordered and restart-recoverable; it is not a same-transaction or multi-file/power-loss atomic operation. The candidate is built from a separate database copy and is reopened and checked before it is reported. The SharePoint source and database inputs remain unchanged.

The transformation remains deliberately candidate-only. The separate `acquire` command may stage validated SharePoint values locally, but it cannot alter the SharePoint source or production and does not claim Cellpy or Excel qualification.

The optional `candidate --cellpy-ready` path adds the example workbook's 12 `db_table` formulas for new rows and asks a private Windows Excel/COM instance (pywin32>=312) to recalculate the temporary candidate. It fails closed with an actionable dependency error when pywin32 is missing, and its cached-value checks do not constitute full semantic Cellpy qualification.

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
python main.py acquire --account ife-user
python main.py candidate --source ... --database ... --output ... [--report ...] [--cellpy-ready]
```

The exact production database path is rejected. In-place output is rejected. `python main.py` and `python main.py validate` remain disabled with exit code 3, and retired legacy flags remain rejected.

## Explicitly deferred

The following are not V1 acceptance criteria and must not be described as verified:

- live SharePoint identity verification, authentication qualification, source manifests, and source anomaly resolution;
- complete `db_table` field/mapping refinement beyond the example formula templates;
- Cellpy installation, qualification, or semantic validation;
- live Excel/COM recalculation qualification beyond the opt-in candidate gate; and
- production workbook replacement, apply, backup, rollback, or any other production write; ordered candidate artifact publication does not claim multi-file atomicity.

Preserved planning, mapping, source, gate, and safety artifacts for the deferred qualification work are in [docs/deferred_cp/](docs/deferred_cp/README.md).

## Evidence boundary

Current evidence includes the acquisition/provider contract tests, candidate tests, formula/cache guards, and a skipped-when-unavailable Cellpy integration test. No live SharePoint permission/path, old-link revocation, business workbook, semantic Cellpy, Excel, or production validation is claimed.

## Next plan

Qualify the opt-in path on the operator's actual Excel/Cellpy installation and representative business workbooks, then review any production replacement separately.
