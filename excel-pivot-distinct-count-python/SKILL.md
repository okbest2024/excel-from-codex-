---
name: excel-pivot-distinct-count-python
description: 使用 Python + Excel COM 创建或重建 Excel 数据透视表，按推荐人统计客户数，并对资金账号做非重复计数（Distinct Count）；不增加辅助列，不改变原始数据结构。适用于客户只需新增记录并刷新透视表即可得到最新统计结果的场景。
---

# Excel 透视表非重复计数（Python）

使用此 Skill 可自动重建可复用的数据透视表，以 `资金账号` 作为客户唯一标识，按 `推荐人姓名` 统计客户数。

## 快速开始

执行：

```powershell
python "C:/Users/Administrator/.codex/skills/excel-pivot-distinct-count-python/scripts/rebuild_pivot_distinct_count.py" --workbook "f:/ai-vscode/订单管理.xlsx"
```

## 工作流程

1. 在工作簿中定位包含 `资金账号` 和 `推荐人姓名` 表头的源数据工作表。
2. 若存在旧的 `统计透视` 工作表，先删除后重建。
3. 保持源数据不变；仅在检测到历史辅助列 `客户去重标记` 时移除该列。
4. 基于源数据创建或调整 Excel 表格区域（Table）。
5. 在 `统计透视` 中创建数据透视表：
- 行字段：`推荐人姓名`
- 值字段：`资金账号`
- 汇总方式：`非重复计数（Distinct Count）`
6. 保存工作簿。

## 环境要求

- Windows 系统
- 已安装 Excel 桌面版
- Python 环境并安装 `pywin32`（`pip install pywin32`）