# Excel 透视表非重复计数 Skill

这是一个给 Codex 使用的 Skill，用于在 Excel 中自动创建/重建数据透视表，按“推荐人姓名”统计唯一客户数（以“资金账号”为唯一标识）。

## 功能说明

- 保持原始数据表结构不变，不依赖辅助列。
- 自动删除旧的 `统计透视` 工作表并重建。
- 透视表行字段使用 `推荐人姓名`。
- 透视表值字段使用 `资金账号`，汇总方式为 **非重复计数（Distinct Count）**。
- 设置透视缓存为打开文件时可刷新。

## 目录结构

- `SKILL.md`：Skill 触发与使用说明。
- `agents/openai.yaml`：UI 元数据。
- `scripts/rebuild_pivot_distinct_count.ps1`：实际执行脚本。

## 快速使用

在 Windows + Excel 桌面版环境执行：

```powershell
powershell -ExecutionPolicy Bypass -File "scripts/rebuild_pivot_distinct_count.ps1" -WorkbookPath "f:/ai-vscode/订单管理.xlsx"
```

## 脚本参数

- `-WorkbookPath`：必填，目标 Excel 文件路径。
- `-PivotSheetName`：可选，默认 `统计透视`。
- `-PivotTitle`：可选，默认 `推荐人客户数统计（按资金账号非重复计数）`。
- `-TableName`：可选，默认 `订单明细`。

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

- 仅支持 Windows 下通过 Excel COM 自动化。
- 运行前请关闭目标 Excel 文件，避免文件占用导致保存失败。
- `资金账号`被视为唯一客户标识，同一账号只计一次。
