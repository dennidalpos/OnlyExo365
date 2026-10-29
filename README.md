![OnlyExo365 app icon](src/OnlyExo365.Shell/Assets/Generated/AppIcon.png)

# OnlyExo365

Windows x64 WPF application for Exchange Online and Microsoft Graph administration. The Shell handles presentation; a separate Worker executes PowerShell; Contracts defines their shared protocol.

Start from the repository root:

```powershell
pwsh ./scripts/bootstrap.ps1 -RuntimeIdentifier win-x64
pwsh ./scripts/build.ps1 -Configuration Debug -RuntimeIdentifier win-x64
pwsh ./scripts/start.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBuild
```

Requirements and verification commands: [development workflow](docs/development/workflow.md). Detailed index by domain: [Documentation Index](docs/README.md).

## Documentation

| Area | Focus | Reference |
| --- | --- | --- |
| **Architecture** | Runtime & project responsibilities | [Architecture overview](docs/architecture/overview.md) |
| | Named pipe transport, limits, cancellation | [IPC protocol](docs/architecture/ipc.md) |
| **Configuration** | Settings precedence & environment overrides | [Configuration](docs/configuration/settings.md) |
| | Paths, logs, atomic reports & data preservation | [Local storage](docs/configuration/storage.md) |
| **Licensing** | SKU cache & embedded catalog update | [License catalog](docs/licensing/catalog.md) |
| **Operations** | Installation, first connection & Message Trace | [Operator guide](docs/operations/operator-guide.md) |
| | Visual user workflows & Mermaid diagrams | [Visual user guide](docs/operations/visual-guide.md) |
| | Exact module policy & authentication modes | [Modules and authentication](docs/operations/authentication.md) |
| | Console toggle, persistent logs & culture | [Diagnostics](docs/operations/diagnostics.md) |
| **Development** | Pinned toolchain, build, clean & local gate | [Development workflow](docs/development/workflow.md) |
| | Isolated VM installer, tenant & UI checks | [Qualification](docs/development/qualification.md) |
| **Releases** | Inno Setup packaging & release bundle | [Packaging](docs/releases/packaging.md) |
| | Windows SDK Authenticode signing pipeline | [Signing](docs/releases/signing.md) |
| **Audits & Evidence** | Reliability, UI, PowerShell, fixtures, review | [Critical audit](docs/audits/2026-10-10-critical-audit.md), [UI audit](docs/audits/2026-10-10-ui-ux-audit.md), [PowerShell audit](docs/audits/2026-10-10-powershell-audit.md), [Retained fixtures](docs/audits/2026-10-10-retained-fixtures.md), [Documentation review](docs/audits/2026-10-10-documentation-review.md) |
| | Bounded IPC read, atomic exports, runspaces | [Local remediation](docs/audits/2026-10-11-local-remediation.md) |
| | Application audit, dependency updates and confirmed residuals | [Application audit](docs/audits/2026-10-27-application-audit.md) |

Open work and current qualification status: 02-areas/products/OnlyExo365/tracker.json (nel dossier SecondBrain) (SecondBrain: `02-areas/products/OnlyExo365/tracker.json`). Audit results describe their recorded revisions; they do not qualify later changes.

Distributed under the proprietary [LICENSE](LICENSE).
