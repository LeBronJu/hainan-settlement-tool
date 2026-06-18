# Handoff

Last updated: 2026-06-18

## Current Goal

Refactor the Hainan settlement automation from a working Python desktop tool into a more maintainable C# WinForms application while preserving the Python version as the full-featured baseline.

The user wants the C# version to be treated as a real project: clear layering, robust error handling, documentation, and GitHub-ready source without real workbook data.

## Current State

- Repository: `D:\Document\文件处理\hainan-settlement-tool`
- GitHub: `https://github.com/LeBronJu/hainan-settlement-tool`
- Active branch: `codex/csharp-stage1`
- Working tree was clean before this documentation update.
- Python version remains the more complete application.
- C# version currently implements the first stage only.

## C# Implementation Status

C# project exists under `csharp/`:

- `HainanSettlementTool.WinForms`: UI, file pickers, logs, status label.
- `HainanSettlementTool.Core`: models and `Stage1Service`.
- `HainanSettlementTool.Excel`: ClosedXML adapters for reading/writing workbooks.

The C# app currently supports stage 1:

- Read an existing power workbook, or build one from `.xlsx`/`.csv` raw detail.
- Copy/update the ledger with current month power.
- Add newly discovered customer names and customer codes where possible.
- Emit a JSON report.

The C# app intentionally does not yet support:

- Direct `.xls` raw-detail cleaning.
- Stage 2 summary generation.
- Agent/intermediary split workbook generation.
- Formula recalculation through installed Excel.

## Environment Status

Installed during this project:

- .NET SDK 8 and 9 are visible through `dotnet --list-sdks`.
- Visual Studio Build Tools 2022 with managed desktop build tools.
- .NET Framework 4.7.2 targeting pack and SDK.

Verified build command:

```powershell
& "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe" "D:\Document\文件处理\hainan-settlement-tool\csharp\HainanSettlementTool.sln" /restore /p:Configuration=Debug /m
```

Last known build result: success, 0 warnings, 0 errors.

Debug exe:

```text
D:\Document\文件处理\hainan-settlement-tool\csharp\src\HainanSettlementTool.WinForms\bin\Debug\net472\海南售电结算自动化工具.exe
```

## Recent Work

Recent commits on `codex/csharp-stage1`:

- `9866b00 Fix WinForms input layout alignment`
- `6537e26 Modernize WinForms stage1 UI`
- `0b07906 Start C# stage1 rewrite`

UI notes:

- The first C# UI looked too old-fashioned.
- It was modernized into a header, stage input panel, workflow panel, and dark log panel.
- The month selector changed from `NumericUpDown` to a fixed dropdown from `2026年2月` through `2026年12月`.
- A layout bug caused labels/buttons to misalign; it was fixed by explicit row creation in the `TableLayoutPanel`.

## Business Context

The user handles monthly Hainan retail electricity settlement work. Current intended workflow:

1. Clean monthly retail-side detail data and generate/import monthly power data.
2. Update the Hainan settlement ledger.
3. User manually reviews volatile ledger fields.
4. Generate agent/intermediary split workbooks.
5. Generate monthly summary workbook.
6. Cross-check outputs against ledger.

The user decided the app should be staged:

- Stage 1: power cleaning/import into ledger plus new customer name/code import.
- Stage 2: after manual ledger cleanup, generate summary and split workbooks.

## High-Risk Business Rules

- Do not overwrite original workbooks.
- Do not rename ledger customers to match summary/payment-account names.
- Units should use `万千瓦时` in the ledger; older wording saying `兆瓦时` was wrong.
- Historical January/February data may contain special irregular cases; avoid encoding those as normal rules.
- New customers may lack负责人/项目开发人; stage 1 may leave those blank for user review.
- If a new customer's invoice/payment note is unknown, default to `走平台扣13%`, except when that agent already has a historical rule.
- `项目开发人` is an agent/intermediary under the salesperson, not the same thing as负责人.

## Suggested Next Steps

1. Review the updated C# UI visually on the user's machine.
2. Add a small sanitized sample workbook set for repeatable tests. Do not use real customer data.
3. Add Core unit tests for matching and stage 1 report semantics.
4. Improve C# file-lock detection and user-facing errors.
5. Port Python stage 2 only after stage 1 has been validated.
6. Add a release packaging script for the C# app that copies the exe and required DLLs into a clean folder.

## Useful Skills For Next Session

- `spreadsheets:Spreadsheets` when modifying or validating `.xlsx` behavior.
- `diagnose` when debugging UI, workbook, or packaging failures.
- `handoff` when compacting long context again.

## Git Notes

Use non-destructive Git operations. Prefer:

```powershell
git status -sb
git diff --stat
git add <specific files>
git commit -m "<message>"
git -c http.sslBackend=openssl push origin codex/csharp-stage1
```

The OpenSSL backend was used because regular Git HTTPS push once failed with a Windows schannel TLS handshake error.

