![OnlyExo365 app icon](src/OnlyExo365.Shell/Assets/Generated/AppIcon.png)

# OnlyExo365

Windows x64 WPF application for Exchange Online and Microsoft Graph administration. The Shell handles presentation; a separate Worker executes PowerShell; Contracts defines their shared protocol.

Start from the repository root:

```powershell
pwsh ./scripts/bootstrap.ps1 -RuntimeIdentifier win-x64
pwsh ./scripts/build.ps1 -Configuration Debug -RuntimeIdentifier win-x64
pwsh ./scripts/start.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBuild
```

Requirements and verification commands: [development workflow](docs/development/workflow.md).

## Documentation

| Area | Reference |
| --- | --- |
| Runtime and source responsibilities | [Architecture](docs/architecture/overview.md) |
| Transport, authentication, cancellation | [IPC](docs/architecture/ipc.md) |
| Exchange/Graph settings and environment | [Configuration](docs/configuration/settings.md) |
| Files, retention and preservation | [Local storage](docs/configuration/storage.md) |
| Install and first connection | [Operator guide](docs/operations/operator-guide.md) |
| Approved modules and authentication limits | [Modules and authentication](docs/operations/authentication.md) |
| Logs, console and troubleshooting | [Diagnostics](docs/operations/diagnostics.md) |
| SKU cache and embedded catalog maintenance | [License catalog](docs/licensing/catalog.md) |
| Build, clean, tests and local gate | [Development workflow](docs/development/workflow.md) |
| Isolated installer, tenant and UI checks | [Qualification](docs/development/qualification.md) |
| Installer and release asset generation | [Packaging](docs/releases/packaging.md) |
| Payload signing and verification | [Signing](docs/releases/signing.md) |
| Reliability / UI / script evidence | [Critical audit](docs/audits/2026-10-10-critical-audit.md), [UI audit](docs/audits/2026-10-10-ui-ux-audit.md), [PowerShell audit](docs/audits/2026-10-10-powershell-audit.md) |
| Retained fixture cleanup | [Cleanup record](docs/audits/2026-10-10-retained-fixtures.md) |
| Documentation/code reconciliation | [Review](docs/audits/2026-10-10-documentation-review.md) |

Open work and current qualification status: [PROJECT_STATUS.json](PROJECT_STATUS.json). Audit results describe their recorded revisions; they do not qualify later changes.

Distributed under the proprietary [LICENSE](LICENSE).
