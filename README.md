![OnlyExo365 app icon](src/ExchangeAdmin.Presentation/Assets/AppIcon.png)

# OnlyExo365

OnlyExo365 is a Windows-first WPF desktop application for Exchange Online administration. The repository ships a desktop operator shell (`ExchangeAdmin.Presentation`) and a separate PowerShell-backed worker process (`ExchangeAdmin.Worker`) connected through shared contracts and named-pipe IPC.

## Verified Repository Scope

Repository evidence currently shows:

- dashboard and worker, Exchange, and Graph connection state
- mailbox workflows for listing, details, storage, restore, permissions, settings, licensing, and access reporting
- recipient and directory areas for contacts, resource mailboxes, distribution lists, public folders, mobile devices, and role groups
- mail operations for mail flow, mail security, message trace, migration, and compliance
- local operator tooling for prerequisite checks, runtime configuration loading, persistent logs, and Microsoft 365 SKU catalog handling

## Windows-First Setup

Minimum repository-backed local baseline:

- Windows
- PowerShell 7
- .NET SDK `10.0.202` from `global.json`

Verified local command path from the repository root:

```powershell
pwsh ./scripts/doctor.ps1
pwsh ./scripts/bootstrap.ps1 -RuntimeIdentifier win-x64
pwsh ./scripts/compile.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBootstrap
pwsh ./scripts/test.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBootstrap
```

Packaging prerequisites, runtime configuration, and release validation stay in the technical documentation linked below.

## Current Project Status

- Verified in this repository context on 2026-04-18: locked restore through `bootstrap.ps1` for `win-x64` and `win-x86`, Debug compile, and the automated test suite.
- Automated tests passed locally with 713 tests.
- Real release validation still depends on checks that cannot be completed from this local context alone: tenant-backed validation, real 32-bit installer validation, and upgrade validation from a legacy `ExchangeAdmin` installation.
- Residual release-validation work is tracked in [PROJECT_STATUS.json](PROJECT_STATUS.json).

## Technical Documentation

- [Architecture overview](docs/architecture/overview.md)
- [Development, build, and test workflow](docs/development/build-test.md)
- [Operator setup and runtime configuration](docs/operations/operator-setup.md)
- [Release and verification](docs/releases/verification.md)
- [Microsoft 365 SKU catalog maintenance](docs/maintenance/licensing-catalog.md)
- [PowerShell SDK major upgrade decision](docs/decisions/2026-04-11-powershell-sdk-major-upgrade-strategy.md)

## License

OnlyExo365 is distributed under the proprietary terms in [LICENSE](LICENSE).
