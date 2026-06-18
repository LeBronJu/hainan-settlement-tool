# Agent Instructions

This repository contains the local desktop automation tool for Hainan retail electricity settlement work. Treat it as a real engineering project, not a one-off script.

## Safety Rules

- Never commit real Excel ledgers, customer lists, settlement outputs, screenshots containing sensitive data, or personal/company financial data.
- Do not overwrite user workbooks in place. All generated workbooks must be written to an output folder or a clearly named copy.
- Preserve user-made changes. If a workbook or source file has unexpected edits, inspect and work with them instead of reverting.
- Prefer explicit validation reports over silent best guesses.
- If business meaning is unclear and cannot be inferred from existing completed month files, ask the user before encoding the rule.

## Repository Layout

- `src/hainan_settlement_tool/`: Python production tool. This is still the practical baseline for full stage 1 and stage 2 behavior.
- `scripts/`: packaging and helper scripts for the Python version.
- `csharp/`: C# rewrite in progress.
- `csharp/src/HainanSettlementTool.WinForms/`: desktop UI only.
- `csharp/src/HainanSettlementTool.Core/`: business models, services, interfaces.
- `csharp/src/HainanSettlementTool.Excel/`: ClosedXML workbook implementation.
- `csharp/docs/architecture.md`: C# layering and migration boundary.
- `HANDOFF.md`: current project status and next-step notes.

## Current Branch

Main C# rewrite work is on `codex/csharp-stage1`.

Recent relevant commits:

- `9866b00 Fix WinForms input layout alignment`
- `6537e26 Modernize WinForms stage1 UI`
- `0b07906 Start C# stage1 rewrite`
- `3007857 Select split template folders separately`
- `08108a9 Initial hainan settlement tool`

## Business Workflow Summary

The monthly settlement workflow is intentionally staged:

1. Clean/import monthly power data into the Hainan settlement ledger.
2. User manually reviews/fills volatile ledger fields such as负责人, 项目开发人, 代理/居间 details, and special cases.
3. Generate salesperson/agent/intermediary split workbooks.
4. Generate the monthly agent-fee summary workbook.
5. Cross-check generated summary and split totals against the ledger.

Stage 1 should be robust and conservative. Stage 2 should not assume missing business fields.

## Important Business Rules

- Ledger power unit should be `万千瓦时`; the older `合同年用电量（兆瓦时）` wording was wrong and should be treated as `万千瓦时` in new outputs.
- Electricity reward calculation uses `0.0001 元/千瓦时`; convert units carefully when reports use `万千瓦时`.
- `项目开发人` is the agent/intermediary relationship under a负责人, not the salesperson themselves.
- Customer legal/account names may differ between the ledger and summary workbook. Do not rename ledger customer names just to match payment account names.
- For new customers, default invoice/payment note is `走平台扣13%` unless the same agent already has an established historical rule, in which case follow that agent's historical rule.
- Special historical months before March 2026 may be irregular. Do not overfit C# rules to January/February quirks unless the user asks.
- March 2026 matters as an active review/correction month; February and earlier should generally be treated as settled historical reference.

## C# Rewrite Principles

- UI must not contain Excel parsing, matching, amount calculation, or template rules.
- Core must not reference ClosedXML, WinForms, or workbook-specific details.
- Excel layer owns workbook reading/writing and template copying.
- Keep stage boundaries explicit. Stage 1 currently ends at a ledger needing human review plus a JSON report.
- Do not migrate stage 2 until stage 1 is stable against real working copies and/or sanitized samples.

## Build Commands

Preferred C# build command:

```powershell
& "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe" "D:\Document\文件处理\hainan-settlement-tool\csharp\HainanSettlementTool.sln" /restore /p:Configuration=Debug /m
```

Python development command:

```powershell
python -m pip install -r requirements.txt
$env:PYTHONPATH = "$PWD\src"
python -m hainan_settlement_tool
```

Python packaging command:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_exe.ps1
```

## Known Compatibility Direction

- C# target is `net472` for Windows 7 SP1 compatibility.
- Running C# builds requires .NET Framework 4.7.2 or newer on the target machine.
- Development machine now has .NET SDK 8/9 and VS Build Tools 2022 installed.
- C# app currently builds successfully with 0 warnings and 0 errors.

