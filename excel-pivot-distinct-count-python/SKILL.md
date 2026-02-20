---
name: excel-pivot-distinct-count-python
description: Create or refresh an Excel PivotTable using Python and Excel COM, counting unique clients per referrer with Distinct Count on account IDs, without adding helper columns or changing source data layout. Use when users update rows in workbooks like order-management files and only need refresh to see updated results.
---

# Excel Pivot Distinct Count Python

Use this skill to rebuild a reusable PivotTable where account ID is the unique client key.

## Quick Start

Run:

```powershell
python "C:/Users/Administrator/.codex/skills/excel-pivot-distinct-count-python/scripts/rebuild_pivot_distinct_count.py" --workbook "f:/ai-vscode/????.xlsx"
```

## Workflow

1. Locate a source worksheet containing headers `????` and `?????`.
2. Delete old pivot worksheet `????` if it exists.
3. Keep source data unchanged; remove old helper column `??????` only when present.
4. Build or resize a source Excel table.
5. Create a new PivotTable in `????`:
- Row field: `?????`
- Value field: `????`
- Aggregate: Distinct Count
6. Save the workbook.

## Requirements

- Windows
- Excel desktop installed
- Python with `pywin32` (`pip install pywin32`)
