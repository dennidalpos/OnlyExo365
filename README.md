![OnlyExo365 app icon](src/ExchangeAdmin.Presentation/Assets/AppIcon.png)

# OnlyExo365

OnlyExo365 is a Windows-first WPF desktop application for Exchange Online administration. The repository contains a desktop operator shell (`ExchangeAdmin.Presentation`) and a separate PowerShell-backed worker process (`ExchangeAdmin.Worker`) connected through shared contracts and named-pipe IPC.

## Verified Feature Set

The current repository evidence shows these operator areas in the shipped shell and worker:

- dashboards and connection state
- mailbox listing, details, storage, restore, permissions, settings, and license workflows
- deleted mailboxes, distribution lists, contacts, resources, and public folders
- mail flow, mail security, message trace, migration, compliance, mobile devices, and permissions
- prerequisite checks in `Tools` for PowerShell 7, execution policy, Exchange Online Management, and Microsoft Graph modules
- local diagnostics and persistent logs

## Windows-First Setup

Minimum verified setup path:

1. Install Windows 10 or Windows 11.
2. Install PowerShell 7 and the .NET SDK `10.0.202` for repository builds, or the .NET Desktop Runtime 10 for running the framework-dependent app payload.
3. Restore and validate the repository from the root:

```powershell
pwsh ./scripts/doctor.ps1
pwsh ./scripts/bootstrap.ps1
pwsh ./scripts/compile.ps1 -Configuration Debug -NoBootstrap
pwsh ./scripts/test.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBootstrap
```

For packaged app installation, setup, runtime configuration, and operator prerequisites, see [Operator setup and runtime configuration](docs/operations/operator-setup.md).

## Current Status

- Local verification completed for the Windows toolchain check, locked restore, Debug compile, and automated test suite.
- The automated test suite passed with 690 tests in this repository context.
- Packaging, signed release, tenant-backed validation, and real upgrade or 32-bit install validation remain separate release checks and are not implied by the local build/test result.
- Concrete remaining release-validation gaps are tracked in [PROJECT_STATUS.json](PROJECT_STATUS.json).

## Technical Documentation

- [Architecture overview](docs/architecture/overview.md)
- [Development, build, and test workflow](docs/development/build-test.md)
- [Operator setup and runtime configuration](docs/operations/operator-setup.md)
- [Release and verification](docs/releases/verification.md)
- [Microsoft 365 SKU catalog maintenance](docs/maintenance/licensing-catalog.md)
- [Architecture decision: PowerShell SDK major upgrade strategy](docs/decisions/2026-04-11-powershell-sdk-major-upgrade-strategy.md)

## License

This repository is distributed under a proprietary license. See [LICENSE](LICENSE).
