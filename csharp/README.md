# C# 重构版

这是海南售电结算自动化工具的 C# 重构版。当前目标不是立刻替换 Python 版，而是用更工程化的结构逐步迁移。

## 当前范围

- 已搭建 WinForms / Core / Excel 三层项目结构。
- 阶段1最小闭环正在迁移：电量处理表或原始明细 -> 待整理台账 -> JSON 报告。
- 阶段2暂不迁移：代理/居间分表、汇总表生成仍以 Python 版为准。

## 项目结构

```text
csharp/
  HainanSettlementTool.sln
  src/
    HainanSettlementTool.WinForms/   # 桌面界面，只负责输入、日志、调用服务
    HainanSettlementTool.Core/       # 业务模型、业务服务、接口
    HainanSettlementTool.Excel/      # ClosedXML 文件读写实现
```

## 设计原则

- UI 不写业务规则。
- Core 不引用 ClosedXML、WinForms、文件格式细节。
- Excel 层只做文件读写和模板复制。
- 每个阶段都输出可审计报告。
- C# 结果必须用脱敏样例和 Python 版输出对照后，才能视为可替代。

## 运行环境目标

- 开发：Visual Studio 2022 或具备 .NET Framework 4.7.2 targeting pack 的 MSBuild 环境。
- 运行：Windows 7 SP1 及以上，需安装 .NET Framework 4.7.2 或更高版本。

## 重要限制

- C# 第一版暂不直接清洗 `.xls` 原始明细；请先另存为 `.xlsx`，或使用已清洗的电量处理表。
- 公式重算、阶段2分表/汇总表生成还没有迁移。
- 当前机器没有 .NET SDK/MSBuild，代码已按项目结构创建，但尚未在本机完成编译验证。
