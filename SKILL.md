---
name: excel-pivot-distinct-count
description: 使用 Excel 数据透视表对客户进行去重统计：按“推荐人姓名”分组，对“资金账号”做非重复计数（Distinct Count），且不改动原始数据结构、不增加辅助列。适用于订单管理类工作簿，需要“新增记录后刷新透视表即可出结果”的非编程场景。
---

# Excel 透视表非重复计数

使用此 Skill 可自动重建一个可复用的数据透视表，以 `资金账号` 作为客户唯一标识，按 `推荐人姓名` 统计客户数。

## 快速开始

执行以下命令：

```powershell
powershell -ExecutionPolicy Bypass -File "C:/Users/Administrator/.codex/skills/excel-pivot-distinct-count/scripts/rebuild_pivot_distinct_count.ps1" -WorkbookPath "f:/ai-vscode/订单管理.xlsx"
```

## 工作流程

1. 校验工作簿中是否存在源数据表头：`资金账号`、`推荐人姓名`。
2. 若已存在 `统计透视` 工作表，先删除再重建。
3. 保持源数据不变；仅在检测到旧版本辅助列 `客户去重标记` 时移除该列。
4. 基于源数据创建或调整 Excel 表格区域（Table）。
5. 在 `统计透视` 中创建数据透视表：
- 行字段：`推荐人姓名`
- 值字段：`资金账号`
- 汇总方式：`非重复计数（Distinct Count）`
6. 保存文件并输出执行结果。

## 注意事项

- `资金账号` 是唯一客户标识，同一账号只计 1 个客户。
- 运行环境为 Windows + Excel 桌面版（COM 自动化）。
- 后续只需新增源数据并在透视表中点击“刷新/全部刷新”。
