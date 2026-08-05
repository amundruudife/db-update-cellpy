---
name: run-graphify
description: Refresh or query this repository's Graphify knowledge graph using the recorded project interpreter, incremental extraction, and low-cost default execution. Use for Graphify refreshes, codebase navigation, graph queries, or post-change graph updates in this repository.
---

# Run Graphify

## Overview

Use the bundled PowerShell runner to refresh or query this repository's Graphify outputs with the recorded interpreter and low-cost defaults.

## Commands

Run from the repository root:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\.agents\skills\run-graphify\scripts\run_graphify.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\.agents\skills\run-graphify\scripts\run_graphify.ps1 -Mode full
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\.agents\skills\run-graphify\scripts\run_graphify.ps1 -Mode query -Question "What calls update_slurry?"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\.agents\skills\run-graphify\scripts\run_graphify.ps1 -Mode path -ExtraArgs @('main', 'database')
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\.agents\skills\run-graphify\scripts\run_graphify.ps1 -Mode explain -Question "update_slurry"
```

- Default mode is `update`, which re-extracts only changed files.
- Use `-Mode full` only when a complete rebuild is explicitly needed.
- Do not use deep extraction or optional exports unless requested.
- When semantic extraction requires a worker, use the default general-purpose worker with no model override. The configured low-cost default is Luna; do not escalate for routine Graphify extraction.
- Treat `graphify-out\` as generated navigation output, not source truth. Verify material facts in source files.

## Outputs

The runner writes `GRAPH_REPORT.md`, `graph.json`, and `graph.html` to `graphify-out\`.

