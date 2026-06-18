# Handoff

Last updated: 2026-06-18

## Current State

This repository is now the Python baseline for the Hainan retail electricity settlement automation tool.

Local path:

```text
D:\Document\文件处理\hainan-settlement-tool
```

GitHub:

```text
https://github.com/LeBronJu/hainan-settlement-tool
```

The C# WinForms rewrite has been split out into a standalone project and repository.

## Split-Out C# Project

Local path:

```text
D:\Document\文件处理\hainan-settlement-desktop
```

GitHub:

```text
https://github.com/LeBronJu/hainan-settlement-desktop
```

The old `csharp/` directory was removed from this repository after the standalone project was created, built, committed, and pushed.

## Python Repository Role

Use this repository for:

- Existing full-featured Python GUI behavior.
- Stage 1 and stage 2 reference implementation.
- Packaging the Python exe through `scripts/build_exe.ps1`.
- Comparing behavior while porting features to C#.

Do not add new C# source here. C# work belongs in `hainan-settlement-desktop`.

## Business Context

The user handles monthly Hainan retail electricity settlement work. Current intended workflow:

1. Clean monthly retail-side detail data and generate/import monthly power data.
2. Update the Hainan settlement ledger.
3. User manually reviews volatile ledger fields.
4. Generate agent/intermediary split workbooks.
5. Generate monthly summary workbook.
6. Cross-check outputs against ledger.

The app should stay staged:

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

## Environment Notes

Installed during this project:

- .NET SDK 8/9.
- Visual Studio Build Tools 2022.
- .NET Framework 4.7.2 targeting pack / SDK.

These are primarily for the standalone C# repository.

## Suggested Next Steps

1. Keep Python fixes and behavior-reference work in this repository.
2. Put all C# development in `D:\Document\文件处理\hainan-settlement-desktop`.
3. Add sanitized sample workbooks in the C# repository before expanding its automated tests.
4. Do not commit real workbook data to either repository.

## Useful Skills For Future Sessions

- `spreadsheets:Spreadsheets` for workbook behavior.
- `diagnose` for debugging workbook or packaging failures.
- `handoff` when compacting context again.

## Git Notes

Use non-destructive Git operations. Prefer:

```powershell
git status -sb
git diff --stat
git add <specific files>
git commit -m "<message>"
git -c http.sslBackend=openssl push
```

The OpenSSL backend was used because regular Git HTTPS push once failed with a Windows schannel TLS handshake error.

