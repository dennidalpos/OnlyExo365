![OnlyExo365 app icon](src/OnlyExo365.Shell/Assets/Generated/AppIcon.png)

# OnlyExo365

Windows x64 desktop application (WPF) for administrators and help desks working with Exchange Online and Microsoft Graph.

## Purpose

Inspect tenant data, edit supported settings and export reports: mailboxes, recipients, message trace, mail security and flow, compliance and diagnostics. Availability depends on tenant permissions and discovered capabilities. The interface is in Italian and English.

## Requirements

- Windows x64; the .NET 10 Desktop Runtime for framework-dependent packages (self-contained ones include it).
- External PowerShell 7.6+ for the interactive Exchange bootstrap, and network access for module installation and cloud operations. Approved module versions are in the [development baseline](docs/development/workflow.md#pinned-baseline).
- To build: .NET SDK 10.0.403 exactly (`global.json` disables roll-forward) and PowerShell 7.

## First setup

Get an installer from the maintainer (or build one: [packaging](docs/releases/packaging.md)), then:

1. Run `OnlyExo365.Setup.exe` as administrator, or open `OnlyExo365.Shell.exe` from an extracted package.
2. Open **Tools** to check prerequisites.
3. Configure [tenant and authentication settings](docs/configuration/settings.md) and use **Connect Exchange**; check Worker, Exchange and Graph status separately.

From source:

```powershell
pwsh ./scripts/bootstrap.ps1 -RuntimeIdentifier win-x64
pwsh ./scripts/build.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -Clean:$false -LockedMode
pwsh ./scripts/start.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBuild
```

## Usage

Read the [operator guide](docs/operations/operator-guide.md) before applying changes. Cloud writes are never retried automatically: after a timeout or partial failure, refresh and inspect the tenant before resubmitting, as there is no rollback. Diagnostic exports can contain tenant identifiers: review them before sharing.

## Build and test

See the [development workflow](docs/development/workflow.md) for validation and the [qualification notes](docs/development/qualification.md): offline tests do not qualify real tenant authentication or the installer.

## Documentation and security

[Documentation index](docs/README.md) · [Authentication](docs/operations/authentication.md) · [Diagnostics](docs/operations/diagnostics.md) · [Architecture](docs/architecture/overview.md) · [Security policy](SECURITY.md) · [License](LICENSE) (proprietary).
