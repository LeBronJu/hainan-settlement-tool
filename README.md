# 海南售电结算自动化工具

用于海南售电结算月度处理的本地桌面工具。

## 功能

- 阶段1：清洗零售侧明细，生成电量处理表，并把当月电量、新增客户名称、户号导入台账。
- 阶段2：读取人工整理后的台账、上月代理分表文件夹、上月居间分表文件夹，生成代理/居间分表和代理费汇总表。
- 保留原始文件，结果统一写入输出文件夹。
- 支持运行日志、文件占用提示、Excel 公式重算。

## C# 重构版

C# 重构版位于 `csharp/`，当前在 `codex/csharp-stage1` 分支推进。它按 WinForms / Core / Excel 三层拆分，第一阶段目标是迁移“阶段1：电量导入台账”。

Python 版仍是当前可用主版本；C# 版在脱敏样例和真实工作副本验证完成前，不视为替代版本。

## 项目文档

- `AGENTS.md`：给后续开发 agent 的项目规则、业务口径和安全边界。
- `HANDOFF.md`：当前上下文压缩版，用于长对话后接续工作。
- `csharp/docs/architecture.md`：C# 重构版分层和迁移边界。

## 阶段1需要的文件

- 基础台账（必填）：上月台账或修正版台账。
- 电量处理表或原始零售侧明细（二选一）。
- 当月结算文件夹（可选）：用于自动识别原始明细和默认输出目录。
- 参考台账（可选）：用于补新增客户基础资料。

## 阶段2需要的文件

- 人工整理后的台账（必填）。
- 上月代理分表文件夹（必填）：通常名为 `2026年代理 - 海南`。
- 上月居间分表文件夹（必填）：通常名为 `2026年居间 - 海南`。
- 上月/修正版汇总表（必填）：用于生成当月汇总表。

## 开发运行

```powershell
python -m pip install -r requirements.txt
$env:PYTHONPATH = "$PWD\src"
python -m hainan_settlement_tool
```

## 打包 Windows exe

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_exe.ps1
```

打包结果输出到 `dist\海南售电结算自动化工具.exe`。

## 注意

- 推荐在 Windows 10/11 上运行。
- 阶段2如需刷新 Excel 公式，电脑上需要安装 Microsoft Excel。
- 不要提交真实结算台账、客户数据、生成结果或个人/公司敏感文件到 GitHub。
