# Agent Instructions

This repository contains the Python desktop automation tool for Hainan retail electricity settlement work. Treat it as a real engineering project, not a one-off script.

## Safety Rules

- Never commit real Excel ledgers, customer lists, settlement outputs, screenshots containing sensitive data, or personal/company financial data.
- Do not overwrite user workbooks in place. All generated workbooks must be written to an output folder or a clearly named copy.
- Preserve user-made changes. If a workbook or source file has unexpected edits, inspect and work with them instead of reverting.
- Prefer explicit validation reports over silent best guesses.
- If business meaning is unclear and cannot be inferred from existing completed month files, ask the user before encoding the rule.

## Repository Layout

- `src/hainan_settlement_tool/`: Python production tool and current full-featured baseline.
- `scripts/`: packaging and helper scripts for the Python version.
- `HANDOFF.md`: current project status and next-step notes.

## Related Repositories

The C# desktop rewrite has been split into a standalone repository:

- `https://github.com/LeBronJu/hainan-settlement-desktop`

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

## Build Commands

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

## Notes

- This repository should not regain a `csharp/` project directory. Put C# work in `hainan-settlement-desktop`.
- Python remains useful as the behavior reference when porting features to C#.
