# Excel 透视表非重复计数 Skill

该仓库现在包含两个版本的 Skill：

- PowerShell 版本（原版本，保留不变）
- Python 版本（新增）

两个版本都实现同一目标：按 `推荐人姓名` 统计唯一客户数，唯一标识是 `资金账号`，汇总方式为数据透视表的 **非重复计数（Distinct Count）**。

## 版本目录

- PowerShell 版（根目录）
  - `SKILL.md`
  - `agents/openai.yaml`
  - `scripts/rebuild_pivot_distinct_count.ps1`
- Python 版（子目录）
  - `excel-pivot-distinct-count-python/SKILL.md`
  - `excel-pivot-distinct-count-python/agents/openai.yaml`
  - `excel-pivot-distinct-count-python/scripts/rebuild_pivot_distinct_count.py`

## Python 版快速使用

```powershell
python "excel-pivot-distinct-count-python/scripts/rebuild_pivot_distinct_count.py" --workbook "f:/ai-vscode/订单管理.xlsx"
```

## PowerShell 版快速使用

```powershell
powershell -ExecutionPolicy Bypass -File "scripts/rebuild_pivot_distinct_count.ps1" -WorkbookPath "f:/ai-vscode/订单管理.xlsx"
```

## 数据要求

源数据工作表需要包含以下表头：

- `资金账号`
- `推荐人姓名`

脚本会自动识别包含这两个表头的工作表作为源表。

## 客户侧使用流程

1. 在源数据表持续新增记录。
2. 打开透视表页 `统计透视`。
3. 点击“刷新”或“全部刷新”，即可得到最新统计结果。

## 注意事项

- 两个版本都依赖 Windows + Excel 桌面版（COM 自动化）。
- Python 版额外需要 `pywin32`：`pip install pywin32`。
- `资金账号`被视为唯一客户标识，同一账号只计一次。
