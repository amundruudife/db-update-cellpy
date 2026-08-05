# Repository guidance

## Graphify navigation

This repository has a generated Graphify knowledge graph:

- [Interactive graph](graphify-out/graph.html)
- [Graph report](graphify-out/GRAPH_REPORT.md)
- [Raw graph data](graphify-out/graph.json)
- [Project Graphify skill](.agents/skills/run-graphify/SKILL.md)

For codebase questions, query the existing graph first through the project skill. Use its default incremental refresh when the graph is stale. The default workflow is deliberately low-cost: no deep extraction or optional exports, and semantic workers use the configured Luna/default worker with no model override.

Run the project-local post-change workflow from the repository root:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\.agents\skills\run-graphify\scripts\run_graphify.ps1
```

`graphify-out\` is generated navigation/audit output, not source truth. Verify important claims against the source files.
