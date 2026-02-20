---
name: excel-pivot-distinct-count
description: Create or refresh an Excel PivotTable that counts unique clients per referrer using Distinct Count on the account column, while keeping the source sheet unchanged and without helper columns. Use when a workbook like order-management files needs a reusable add-rows-then-refresh workflow for non-technical users.
---

# Excel Pivot Distinct Count

Use this skill to rebuild a reusable PivotTable where account ID is the unique client key.

## Quick Start

Run:

```powershell
powershell -ExecutionPolicy Bypass -File "C:/Users/Administrator/.codex/skills/excel-pivot-distinct-count/scripts/rebuild_pivot_distinct_count.ps1" -WorkbookPath "f:/ai-vscode/????.xlsx"
```

## Workflow

1. Verify the workbook has source headers for account and referrer.
2. Remove existing `????` sheet if present.
3. Keep source data unchanged; only remove old helper column `??????` if it exists.
4. Build or resize an Excel table for source data.
5. Create a new PivotTable in `????`:
- Row field: `?????`
- Value field: `????`
- Aggregate: Distinct Count
6. Save workbook and output completion info.

## Notes

- Treat `????` as the only unique customer identifier.
- Target environment is Excel desktop on Windows with COM automation.
- After new rows are added, users only need to refresh the pivot sheet.
