# Architectural Decision Record: PowerShell SDK 7.6.6 and .NET 10 Baseline

## Context

OnlyExo365 executes PowerShell commands out-of-process in `OnlyExo365.Worker` while driving the user interface in `OnlyExo365.Shell`. The worker requires direct compatibility with modern Exchange Online and Graph cmdlets.

## Decision

1. Pin runtime framework to `.NET 10` across all projects (`net10.0` for Contracts/Worker, `net10.0-windows` for Shell/Tests).
2. Pin build toolchain to `.NET SDK 10.0.401` via `global.json` with disabled roll-forward.
3. Use `Microsoft.PowerShell.SDK 7.6.6` inside `OnlyExo365.Worker`.
4. Target `win-x64` exclusively for packaging and distribution.

## Rationale

- Ensures end-to-end toolchain closure between the desktop runtime, PowerShell SDK, and packaging scripts.
- Prevents runtime assembly version drift between Exchange cmdlets and the PowerShell host engine.
- Consolidates repository boundaries into Shell, Worker, and Contracts without redundant abstraction layers.

## Maintenance Revision — 2026-10-10

Updated the Worker SDK to 7.6.6 while retaining .NET 10 and SDK 10.0.401. Approved ExchangeOnlineManagement is 3.10.1 and the four Graph modules are 2.41.1. External hosts importing Exchange require PowerShell 7.6+. System.Security.Cryptography.ProtectedData and System.Security.Cryptography.Xml are 10.0.12; OpenXml and test dependencies remain unchanged. Lockfiles are regenerated through NuGet. See the [audit report](../audits/2026-10-10-critical-audit.md) for current primary sources and verification scope.
